"""The built answer's sheets moved into the user's open workbook and finished there (the COM side of the file render).

move_in inserts every sheet of the temporary file (render/xlsx.build) with ONE Sheets.Add(Type:=<its path>): no
workbook is opened, nothing is copied or closed, and the insert leaves no link to the file (scripts/check_insert.py
proved it the same as the Open + Copy + Close it replaced, live 20.11.2026). Names are checked first: the insert
would silently rename a taken 'Q5 Data' to 'Q5 Data (2)' and every step by name would then hit the wrong sheet
(docs/audit/2026-11-19-render-map.md section 10 item 3). A sheet holding an Excel Table is never moved between
workbooks (Excel can crash), so the file holds none (its zip is checked before the insert) and finish makes the
tables in the workbook, then writes the formula blocks (the formulas name the tables), then sets the per-point chart
labels the file cannot carry (LabelFix). quiet keeps the user's Excel as it
was found around all of this and lands on the answer when it succeeds; show lets Excel paint once in between; park /
restore / drop rename a question's old sheets out of the way for a follow-up and delete or restore them afterwards.

Book is the job's picture of the workbook's sheets: read once (read_book) and kept up to date by the steps that add,
rename, move or delete sheets, so none of them reads the sheets again.

Every COM call goes through jobs.com (busy retry, abort check, the call count). A call that must not run twice
(Sheets.Add, ListObjects.Add, Delete) is alone in its call, as jobs.com reruns the whole call while Excel is busy;
the other calls batch only absolute sets, so a rerun sets the same. Late-bound COM ignores keyword arguments, so
every argument is positional (see xl/sheets.add_sheet)."""
from __future__ import annotations

import contextlib
import logging
import re
import zipfile
from dataclasses import dataclass, field

from ..render.xlsx import Base, BuiltFile, TableSpec
from ..render.xlsx_charts import LabelFix
from . import jobs
from .host import HostError, InputError
from .sheetorder import ANSWER_TAB, DATA_TAB

log = logging.getLogger("smartview.movein")

MANUAL = -4135                          # xlCalculationManual
QUIET = (("ScreenUpdating", False), ("EnableEvents", False), ("DisplayAlerts", False), ("Calculation", MANUAL))
HIDDEN = (0, 2)                         # xlSheetHidden, xlSheetVeryHidden
SRC_RANGE, HAS_HEADERS = 1, 1           # xlSrcRange, xlYes (ListObjects.Add)
TABLE_STYLE = "TableStyleLight1"
POSITION = {"right": -4152, "above": 0, "below": 1, "left": -4131}       # xlLabelPositionRight, Above, Below, Left
SHEET_NAME_MAX = 31
INSERT_WAIT_S = 120                     # the insert reads the file (as Workbooks.Open did, which had this limit)
LABEL_WAIT_S = 10                       # a label is cosmetic: a busy Excel is not waited for as long as for the rest
STANDARD_PX = 64                        # XlsxWriter's standard column (Calibri 11), when the destination's is unknown
DATE_SETTINGS = (16, 31, 40, 41, 42)    # 0-based Application.International: xlDateSeparator, xlDateOrder,
                                        # xlMonthLeadingZero, xlDayLeadingZero, xl4DigitYears
SHORT_DATE = 14                         # Excel's built-in short date (the system's)


class NameTaken(InputError):
    """A sheet or table name of the answer already exists in the workbook."""


@dataclass
class Quiet:
    """What quiet yields: active - the workbook was its Excel's active one when the job began, visible or not;
    shown - and that Excel was visible (show may paint it); land - a sheet of that workbook to end on when the body
    succeeds (else the user's sheet)."""
    shown: bool = False
    land: str | None = None
    active: bool = False


def _active(app, wb) -> tuple[bool, bool]:
    """(active, shown): wb is its Excel's active workbook; and that Excel is visible."""
    if wb is None:
        return False, False
    active = app.ActiveWorkbook
    is_active = active is not None and str(active.FullName) == str(wb.FullName)
    return is_active, is_active and bool(app.Visible)


@contextlib.contextmanager
def quiet(app, wb=None):
    """ScreenUpdating, EnableEvents and DisplayAlerts off and Calculation manual while the answer is moved in and
    finished. Afterwards (also when a setting or the body fails) the sheet to land on is activated - land when the
    body succeeded, else the user's active sheet - and the four settings are put back in reverse order: in one call,
    each setting on its own inside it (and each in a call of its own when that call fails). A restore that fails is
    logged and never replaces the error that is on its way out. Three calls around a body that succeeds."""
    before, sheet, (active, shown) = jobs.com(lambda: ([getattr(app, name) for name, _ in QUIET], app.ActiveSheet,
                                                       _active(app, wb)))
    state, ok = Quiet(shown, active=active), False

    def off():
        for name, value in QUIET:
            setattr(app, name, value)
    try:
        jobs.com(off)
        yield state
        ok = True
    finally:
        land = (lambda: wb.Worksheets(state.land).Activate()) if ok and state.land else (
            (lambda: (sheet.Parent.Activate(), sheet.Activate())) if sheet is not None else None)
        settings = list(reversed(list(zip((n for n, _ in QUIET), before))))
        try:
            jobs.com(lambda: _put_back(app, land, settings), limit=30, check=False)
        except Exception:   # noqa: BLE001 - Excel stayed busy: each on its own, as before
            for name, value in settings:
                try:
                    jobs.com(lambda name=name, value=value: setattr(app, name, value), limit=30, check=False)
                except Exception as e:   # noqa: BLE001
                    log.warning("could not restore Application.%s=%s: %s", name, value, e)


def _put_back(app, land, settings) -> None:
    """quiet's way out inside one call: a busy Excel is raised (the call is rerun: every set is absolute); any other
    failure is logged and the rest goes on."""
    if land is not None:
        try:
            land()
        except Exception as e:   # noqa: BLE001 - e.g. the user's active sheet was a parked one, now deleted
            _raise_busy(e)
            log.warning("could not re-activate the sheet to land on: %s", e)
    for name, value in settings:
        try:
            setattr(app, name, value)
        except Exception as e:   # noqa: BLE001
            _raise_busy(e)
            log.warning("could not restore Application.%s=%s: %s", name, value, e)


def _raise_busy(e: BaseException) -> None:
    """Inside a batched call: a busy Excel or an abandoned job goes on up (jobs.com reruns or stops the call)."""
    if isinstance(e, jobs.JobAborted) or jobs.is_busy(e):
        raise e


def show(app, wb, name: str) -> None:
    """Let Excel paint once with the sheet name of wb active, screen updating off again afterwards: two calls
    (activate + screen updating on, then off). Only for a workbook quiet found shown (never steal focus from another
    workbook). Cosmetic: a failure is logged."""
    for call in (lambda: (wb.Worksheets(name).Activate(), setattr(app, "ScreenUpdating", True)),
                 lambda: setattr(app, "ScreenUpdating", False)):     # off again, also when the first call failed
        try:
            jobs.com(call)
        except jobs.JobAborted:
            raise
        except Exception as e:   # noqa: BLE001
            log.warning("could not show %s: %s", name, e)


@dataclass
class Sheet:
    name: str
    ws: object = None                    # the COM worksheet (None for a sheet the job inserted: found by name)
    visible: int = -1                    # Worksheet.Visible: -1 visible, 0 hidden, 2 very hidden
    tab: object = None                   # Tab.Color (False: none; None: Excel would not say)
    tables: list = field(default_factory=list)      # [(table name, ListObject)]


@dataclass
class Book:
    """The workbook's sheets in order, its chart sheets and the active sheet (when it is the active workbook)."""
    sheets: list[Sheet]
    charts: list[str] = field(default_factory=list)
    active: str | None = None

    def names(self) -> list[str]:
        return [s.name for s in self.sheets]

    def tables(self) -> list[tuple]:
        """Every table as (name, ListObject, sheet name), in workbook order."""
        return [(n, lo, s.name) for s in self.sheets for n, lo in s.tables]

    def order(self, names: list[str]) -> None:
        """The sheets in this order now (after moves)."""
        by = {s.name: s for s in self.sheets}
        self.sheets = [by[n] for n in names]


def _cosmetic(read):
    try:
        return read()
    except Exception as e:   # noqa: BLE001 - e.g. Excel would not give a tab colour
        _raise_busy(e)
        return None


def _read_book(wb) -> Book:
    sheets = [Sheet(str(ws.Name), ws, int(ws.Visible), _cosmetic(lambda ws=ws: ws.Tab.Color),
                    [(str(lo.Name), lo) for lo in ws.ListObjects]) for ws in wb.Worksheets]
    app = wb.Application
    active = _cosmetic(lambda: str(wb.ActiveSheet.Name) if str(app.ActiveWorkbook.FullName) == str(wb.FullName)
                       else None)
    return Book(sheets, [str(c.Name) for c in wb.Charts], active)


def read_book(t) -> Book:
    """The workbook's Book in ONE call: each worksheet's name, visibility, tab colour and tables, the chart sheets'
    names and the active sheet."""
    return jobs.com(lambda: _read_book(t.wb))


def _inventory(wb) -> tuple[list, list]:
    """Every worksheet as (name, sheet) and every table as (name, table, sheet name), in workbook order."""
    sheets, tables = [], []
    for ws in wb.Worksheets:
        name = str(ws.Name)
        sheets.append((name, ws))
        tables += [(str(lo.Name), lo, name) for lo in ws.ListObjects]
    return sheets, tables


def _names_defined(wb, names: list[str]) -> list[str]:
    """Which of these are a defined name of the workbook (a table cannot take one), inside one call."""
    out = []
    for n in names:
        try:
            out.append(str(wb.Names(n).Name))
        except Exception as e:   # noqa: BLE001 - no such name
            _raise_busy(e)
    return out


def taken(wb, sheets: list[str], tables: list[str], book: Book | None = None) -> list[str]:
    """What of these names the workbook already has, each as it reads in a message ("the sheet 'Q5 Data'"): a sheet
    or chart sheet of an answer sheet's name, a table or defined name of a table's name (names compare without case,
    as in Excel; a table name is taken on any sheet). book: the workbook's sheets as the job knows them (else one
    read); the defined names are one call."""
    book = book or jobs.com(lambda: _read_book(wb))
    want_sheets, want_tables = {s.casefold() for s in sheets}, {n.casefold() for n in tables}
    names = jobs.com(lambda: _names_defined(wb, list(tables))) if tables else []
    return ([f"the sheet '{n}'" for n in book.names() if n.casefold() in want_sheets]
            + [f"the chart sheet '{n}'" for n in book.charts if n.casefold() in want_sheets]
            + [f"the table '{n}' on sheet '{s}'" for n, _, s in book.tables() if n.casefold() in want_tables]
            + [f"the name '{n}'" for n in names])


def _taken_text(found: list[str], book: str) -> str:
    if len(found) == 1:
        return f"{found[0]} already exists in {book}: rename or delete it, then ask again"
    return f"{', '.join(found[:-1])} and {found[-1]} already exist in {book}: rename or delete them, then ask again"


def _anchor(book: Book, after: str | None, first: bool = False) -> tuple:
    """(Before, After) Sheets of the Book for the insert: after the sheet named after (None: the workbook's last
    sheet), or after the last visible sheet before it when it is hidden (an insert after a hidden sheet is not
    verified);
    before the first visible sheet when none comes before, or when first."""
    sheets = book.sheets
    if first:
        return next((s for s in sheets if s.visible not in HIDDEN), sheets[0]), None
    i = len(sheets) - 1 if after is None else book.names().index(after)
    for s in reversed(sheets[:i + 1]):
        if s.visible not in HIDDEN:
            return None, s
    for s in sheets[i + 1:]:
        if s.visible not in HIDDEN:
            return s, None
    return None, sheets[i]


_TABLE_NAME = re.compile(rb'\bdisplayName="([^"]*)"')


def file_tables(path) -> list[str]:
    """The Excel Tables in an .xlsx (its xl/tables parts), read from the zip: none in a file render/xlsx built."""
    with zipfile.ZipFile(path) as z:
        parts = [n for n in z.namelist() if n.startswith("xl/tables/") and n.endswith(".xml")]
        return [m.group(1).decode() if (m := _TABLE_NAME.search(z.read(n))) else n for n in parts]


def move_in(t, built: BuiltFile, after: str | None, first: bool = False, book: Book | None = None) -> None:
    """Insert the built file's sheets into t.wb after the sheet named after (None: after the last sheet; first:
    before the first visible sheet instead, where sheets.add_sheet puts a sheet when no sheet is there to follow)
    with ONE Sheets.Add(Type:=path), alone in its call. A name of the answer already in the workbook is NameTaken and
    a file holding a table a ValueError, both before the insert. The names are added to t.sheets_added (what a
    rollback deletes) just before it, so an insert that fails half-way is rolled back too. book (else read): the
    workbook's sheets, told where the insert landed. Two calls (the defined names, the insert)."""
    book = book or read_book(t)
    found = taken(t.wb, built.sheets, [spec.name for spec in built.tables], book)
    if found:
        raise NameTaken(_taken_text(found, jobs.com(lambda: str(t.wb.Name))))
    tables = file_tables(built.path)
    if tables:
        raise ValueError(f"the built file {built.path.name} holds an Excel Table ({', '.join(tables)}): a sheet "
                         "with a table is never moved between workbooks; the tables are made after the insert")
    before, after_sheet = _anchor(book, after, first)
    path, names = str(built.path), list(built.sheets)
    t.sheets_added.extend(names)
    if before is not None:
        jobs.com(lambda: t.wb.Sheets.Add(_ws(t, before), None, None, path), limit=INSERT_WAIT_S)
    else:
        jobs.com(lambda: t.wb.Sheets.Add(None, _ws(t, after_sheet), None, path), limit=INSERT_WAIT_S)
    k = book.sheets.index(before) if before is not None else book.sheets.index(after_sheet) + 1
    book.sheets[k:k] = [Sheet(n, None, -1, ANSWER_TAB if i == 0 else DATA_TAB) for i, n in enumerate(names)]


def _ws(t, sheet: Sheet):
    """The COM worksheet of a Book sheet (one the job inserted is found by name), inside a call."""
    return sheet.ws if sheet.ws is not None else t.wb.Worksheets(sheet.name)


def _read_base(wb) -> tuple:
    font = wb.Styles("Normal").Font
    ws = wb.Worksheets(1)
    last = ws.Columns(ws.Columns.Count)
    intl = wb.Application.International
    return (str(font.Name), float(font.Size), float(ws.StandardHeight), float(ws.StandardWidth),
            float(last.ColumnWidth), float(last.Width), tuple(intl[i] for i in DATE_SETTINGS))


def base(t) -> Base:
    """The destination's Normal font, standard sizes and date format, the base the file is built in
    (render/xlsx.Base), in one call: the first sheet's standard row height, and its last column's width in pixels
    when that column has the standard width (else XlsxWriter's 64 px, logged: the inserted widths are then off by
    the digit widths' ratio)."""
    name, size, height, standard, chars, points, dates = jobs.com(lambda: _read_base(t.wb))
    col_px = round(points / 0.75)
    if chars != standard:
        log.warning("no standard column to measure in %s (%s, standard %s characters): columns in %s px units",
                    t.wb.Name, chars, standard, STANDARD_PX)
        col_px = STANDARD_PX
    return Base(name, int(size) if size.is_integer() else size, round(height / 0.75), col_px, date_format(*dates))


def date_format(sep: str, order, month_zero: bool, day_zero: bool, century: bool) -> str | int:
    """How Excel stores drill_xl's local date format 'dd.mm.yyyy' (set through NumberFormatLocal): as its built-in
    short date where that is the locale's short date (day first, '.' between, leading zeros, four-digit years: nb-NO;
    live, 19.11.2026), else as a code whose dots are the date separator ('dd/mm/yyyy') or literal."""
    if sep == "." and int(order) == 1 and month_zero and day_zero and century:
        return SHORT_DATE
    return "dd/mm/yyyy" if sep == "." else "dd.mm.yyyy"


def _add_table(wb, spec: TableSpec):
    ws = wb.Worksheets(spec.sheet)
    return ws.ListObjects.Add(SRC_RANGE, ws.Range(spec.ref), None, HAS_HEADERS)


def _looks(tables: list) -> None:
    """Every new table's name and look, inside one call (absolute sets)."""
    for lo, spec in tables:
        lo.Name = spec.name
        lo.TableStyle = TABLE_STYLE
        if not spec.stripes:
            lo.ShowTableStyleRowStripes = False


def _bgr(colour: str) -> int:
    """'#RRGGBB' as a COM colour (red + 256 green + 65536 blue)."""
    r, g, b = (int(colour[i:i + 2], 16) for i in (1, 3, 5))
    return r + (g << 8) + (b << 16)


def _find_label(ws, fix: LabelFix) -> tuple:
    """The point's DataLabel, and its Top when the fix lifts it (read first, so a rerun sets the same value): None
    when the label is not more than the lift below the chart's top, which charts._labels leaves where it is (the
    empty points of a waterfall's other step series)."""
    lab = ws.ChartObjects(fix.chart).Chart.SeriesCollection(fix.series).Points(fix.point).DataLabel
    top = lab.Top if fix.lift_pt else None
    return lab, (top if isinstance(top, (int, float)) and top > fix.lift_pt else None)


def _skipped(fix: LabelFix, e: BaseException) -> None:
    log.warning("chart label fix skipped (chart %s, series %s, point %s): %s", fix.chart, fix.series, fix.point, e)


def _find_labels(wb, sheet: str, fixes: list[LabelFix]) -> list[tuple]:
    """[(fix, label, top)] of the fixes whose label Excel gives (the others skipped with a log line), in one call."""
    ws, out = wb.Worksheets(sheet), []
    for fix in fixes:
        try:
            out.append((fix, *_find_label(ws, fix)))
        except Exception as e:   # noqa: BLE001
            _raise_busy(e)
            _skipped(fix, e)
    return out


def _set_labels(found: list[tuple]) -> None:
    """Each found label's position, colour and lift (from the top read before: a rerun sets the same), in one call."""
    for fix, lab, top in found:
        try:
            if fix.position:
                lab.Position = POSITION[fix.position]
            if fix.font_color:
                lab.Font.Color = _bgr(fix.font_color)
            if top is not None:
                lab.Top = top - fix.lift_pt
        except Exception as e:   # noqa: BLE001
            _raise_busy(e)
            _skipped(fix, e)


def _formulas(wb, blocks) -> None:
    for block in blocks:
        wb.Worksheets(block.sheet).Range(block.ref).Formula2 = tuple(tuple(row) for row in block.grid)


def finish(t, built: BuiltFile) -> None:
    """After move_in: each table (ListObjects.Add alone, then every table's name and TableStyleLight1, row stripes
    off where the spec says, in one call), then every formula block in one call (its grid as it is: text keeps its
    leading apostrophe), then the chart label fixes on the Answer sheet (one call reads the labels, one sets them).
    A label fix that fails is skipped with a log line: a label is cosmetic."""
    added = [(jobs.com(lambda spec=spec: _add_table(t.wb, spec)), spec) for spec in built.tables]
    if added:
        jobs.com(lambda: _looks(added))
    if built.formulas:
        jobs.com(lambda: _formulas(t.wb, built.formulas))
    if not built.label_fixes:
        return
    try:
        found = jobs.com(lambda: _find_labels(t.wb, built.sheets[0], built.label_fixes), limit=LABEL_WAIT_S)
        jobs.com(lambda: _set_labels(found), limit=LABEL_WAIT_S)
    except jobs.JobAborted:
        raise
    except Exception as e:   # noqa: BLE001 - Excel stayed busy: the labels stay as the file drew them
        log.warning("chart label fixes skipped: %s", e)


def calculate(t, sheets: list[str]) -> None:
    """Worksheet.Calculate per sheet in one call, the Answer sheets last (they read Data and Lines); the other
    sheets in the order given."""
    order = sorted(sheets, key=lambda s: s.endswith(" Answer"))
    jobs.com(lambda: [t.wb.Worksheets(name).Calculate() for name in order])


@dataclass
class Parked:
    sheets: dict[str, str] = field(default_factory=dict)        # old name -> parked name ("old Q3 Answer")
    tables: dict[str, str] = field(default_factory=dict)        # old name -> parked name ("Data_Q3_old")


def _free(base: str, used: set[str], limit: int | None = None, sep: str = " ") -> str:
    """base, or base with ' 2', ' 3' ... (sep '' for tables: 'Data_Q3_old2'), cut to limit, not in used (casefold)."""
    name, i = base[:limit], 2
    while name.casefold() in used:
        suffix = f"{sep}{i}"
        name = (base[:limit - len(suffix)] if limit else base) + suffix
        i += 1
    used.add(name.casefold())
    return name


def park(t, sheets: list[str], tables: list[str], book: Book | None = None) -> Parked:
    """Rename a question's sheets to 'old <name>' and its tables to '<name>_old' (a free name when that is taken),
    so the new answer can take their names: every rename in one call (a rerun sets the same names). A name the
    workbook does not have is an InputError, before any rename. All or nothing: a park that fails puts back what it
    renamed. book (else read): the workbook's sheets, told the new names."""
    book = book or read_book(t)
    by_sheet = {s.name.casefold(): s for s in book.sheets}
    by_table = {n.casefold(): (lo, s) for s in book.sheets for n, lo in s.tables}
    missing = ([f"the sheet '{n}'" for n in sheets if n.casefold() not in by_sheet]
               + [f"the table '{n}'" for n in tables if n.casefold() not in by_table])
    if missing:
        raise InputError(f"{' and '.join(missing)} to replace {'is' if len(missing) == 1 else 'are'} not in the "
                         "workbook any more: ask the question again as a new one")
    used_sheets, used_tables = set(by_sheet), set(by_table)
    parked, plan = Parked(), []                                  # plan: (COM object, new name)
    for name in sheets:
        parked.sheets[name] = _free("old " + name, used_sheets, SHEET_NAME_MAX)
        plan.append((by_sheet[name.casefold()].ws, parked.sheets[name]))
    for name in tables:
        parked.tables[name] = _free(name + "_old", used_tables, sep="")
        plan.append((by_table[name.casefold()][0], parked.tables[name]))
    try:
        jobs.com(lambda: [setattr(obj, "Name", new) for obj, new in plan])
    except BaseException:
        try:
            restore(t, parked)
        except Exception as e:   # noqa: BLE001 - the park's own error is the one to raise
            log.warning("a failed park could not put every name back: %s", e)
        raise
    for name, new in parked.sheets.items():
        by_sheet[name.casefold()].name = new
        if book.active is not None and book.active.casefold() == name.casefold():
            book.active = new                                    # the active sheet is the parked one now
    for name, new in parked.tables.items():
        lo, s = by_table[name.casefold()]
        s.tables = [(new if x is lo else n, x) for n, x in s.tables]
    return parked


def restore(t, parked: Parked) -> None:
    """After a failure (the new sheets already removed): the parked sheets and tables get their names back. Every
    name is tried; a HostError then names those that could not go back (their name is taken or the sheet is gone)."""
    have_sheets, have_tables = jobs.com(lambda: _inventory(t.wb), check=False)
    left: list[str] = []
    for kind, items, have in (("sheet", parked.sheets, [(n, ws) for n, ws in have_sheets]),
                              ("table", parked.tables, [(n, lo) for n, lo, _ in have_tables])):
        by = {n.casefold(): obj for n, obj in have}
        for old, new in items.items():
            obj = by.get(new.casefold())
            if obj is None and old.casefold() in by:
                continue                                         # already back
            if obj is None or old.casefold() in by:
                left.append(f"{kind} '{new}' -> '{old}'")
                continue
            try:
                jobs.com(lambda obj=obj, old=old: setattr(obj, "Name", old), limit=30, check=False)
            except Exception as e:   # noqa: BLE001
                log.warning("could not rename %s '%s' back to '%s': %s", kind, new, old, e)
                left.append(f"{kind} '{new}' -> '{old}'")
    if left:
        raise HostError("could not put back the names of " + ", ".join(left))


def drop(t, parked: Parked, book: Book | None = None, alerts_off: bool = False) -> None:
    """After success: the parked sheets (and the tables on them) deleted in ONE call, DisplayAlerts off meanwhile
    (alerts_off: it is already, inside quiet - then the Delete is the only call). book: told they are gone."""
    names = list(parked.sheets.values())
    if not names:
        return
    if alerts_off:
        jobs.com(lambda: t.wb.Worksheets(names).Delete())
    else:
        prev = jobs.com(lambda: t.app.DisplayAlerts)
        jobs.com(lambda: setattr(t.app, "DisplayAlerts", False))
        try:
            jobs.com(lambda: t.wb.Worksheets(names).Delete())
        finally:
            jobs.com(lambda: setattr(t.app, "DisplayAlerts", prev), limit=30, check=False)
    if book is not None:
        gone = {n.casefold() for n in names}
        book.sheets = [s for s in book.sheets if s.name.casefold() not in gone]
        if book.active is not None and book.active.casefold() in gone:
            book.active = None
