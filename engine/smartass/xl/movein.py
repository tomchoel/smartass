"""The built answer's sheets moved into the user's open workbook and finished there (the COM side of the file render).

move_in copies every sheet of the temporary file (render/xlsx.build) in ONE Worksheets.Copy: the file is opened
read-only with its window hidden, copied, closed, and the links the copy may leave to it are broken. Names are
checked first: Copy would silently rename a taken 'Q5 Data' to 'Q5 Data (2)' and every step by name would then hit
the wrong sheet (docs/audit/2026-11-19-render-map.md section 10 item 3). A sheet holding an Excel Table is never
copied between workbooks (Excel can crash), so the file holds none and finish makes the tables in the workbook, then
writes each formula block with one Range.Formula2 (the formulas name the tables), then sets the per-point chart labels
the file cannot carry (LabelFix). quiet keeps the user's Excel as it was found around all of this; park / restore /
drop rename a question's old sheets out of the way for a follow-up and delete or restore them afterwards.

Every COM call goes through jobs.com (busy retry, abort check, the call count). A call that must not run twice
(Open, Copy, ListObjects.Add, Delete) is alone in its call, as jobs.com reruns the whole call while Excel is busy.
Late-bound COM ignores keyword arguments, so every argument is positional (see xl/sheets.add_sheet)."""
from __future__ import annotations

import contextlib
import logging
import re
from dataclasses import dataclass, field

from ..render.xlsx import BuiltFile, TableSpec
from ..render.xlsx_charts import LabelFix
from . import jobs
from .host import HostError, InputError

log = logging.getLogger("smartview.movein")

MANUAL = -4135                          # xlCalculationManual
QUIET = (("ScreenUpdating", False), ("EnableEvents", False), ("DisplayAlerts", False), ("Calculation", MANUAL))
HIDDEN = (0, 2)                         # xlSheetHidden, xlSheetVeryHidden
LINKS = 1                               # xlExcelLinks (LinkSources, BreakLink)
SRC_RANGE, HAS_HEADERS = 1, 1           # xlSrcRange, xlYes (ListObjects.Add)
TABLE_STYLE = "TableStyleLight1"
POSITION = {"right": -4152, "above": 0, "below": 1, "left": -4131}       # xlLabelPositionRight, Above, Below, Left
SHEET_NAME_MAX = 31
LABEL_WAIT_S = 10                       # a label is cosmetic: a busy Excel is not waited for as long as for the rest


class NameTaken(InputError):
    """A sheet or table name of the answer already exists in the workbook."""


@contextlib.contextmanager
def quiet(app):
    """ScreenUpdating, EnableEvents and DisplayAlerts off and Calculation manual while the answer is moved in and
    finished. Afterwards (also when a setting or the body fails) the active sheet is re-activated and the four
    settings are put back in reverse order, each on its own: a restore that fails is logged and never replaces the
    error that is on its way out."""
    before, sheet = jobs.com(lambda: ([getattr(app, name) for name, _ in QUIET], app.ActiveSheet))

    def off():
        for name, value in QUIET:
            setattr(app, name, value)
    try:
        jobs.com(off)
        yield
    finally:
        if sheet is not None:
            try:
                jobs.com(lambda: (sheet.Parent.Activate(), sheet.Activate()), limit=30, check=False)
            except Exception as e:   # noqa: BLE001 - e.g. the user's active sheet was a parked one, now deleted
                log.warning("could not re-activate the user's sheet: %s", e)
        for (name, _), value in reversed(list(zip(QUIET, before))):
            try:
                jobs.com(lambda name=name, value=value: setattr(app, name, value), limit=30, check=False)
            except Exception as e:   # noqa: BLE001
                log.warning("could not restore Application.%s=%s: %s", name, value, e)


def _inventory(wb) -> tuple[list, list]:
    """Every worksheet as (name, sheet) and every table as (name, table, sheet name), in workbook order."""
    sheets, tables = [], []
    for ws in wb.Worksheets:
        name = str(ws.Name)
        sheets.append((name, ws))
        tables += [(str(lo.Name), lo, name) for lo in ws.ListObjects]
    return sheets, tables


def taken(wb, sheets: list[str], tables: list[str]) -> list[str]:
    """What of these names the workbook already has (sheet names and table names compare without case, as in Excel;
    a table name is taken on any sheet), each as it reads in a message: "the sheet 'Q5 Data'"."""
    have_sheets, have_tables = jobs.com(lambda: _inventory(wb))
    want_sheets, want_tables = {s.casefold() for s in sheets}, {n.casefold() for n in tables}
    return ([f"the sheet '{n}'" for n, _ in have_sheets if n.casefold() in want_sheets]
            + [f"the table '{n}' on sheet '{s}'" for n, _, s in have_tables if n.casefold() in want_tables])


def _taken_text(found: list[str], book: str) -> str:
    if len(found) == 1:
        return f"{found[0]} already exists in {book}: rename or delete it, then ask again"
    return f"{', '.join(found[:-1])} and {found[-1]} already exist in {book}: rename or delete them, then ask again"


def _anchor(wb, after, first: bool = False) -> tuple:
    """(Before, After) for Copy: after 'after' (None: the workbook's last sheet), or after the last visible sheet
    before it when it is hidden (a copy after a hidden sheet is not verified); before the first visible sheet when
    none comes before, or when first."""
    sheets = list(wb.Worksheets)
    if first:
        return next((s for s in sheets if s.Visible not in HIDDEN), sheets[0]), None
    names = [str(s.Name) for s in sheets]
    i = len(sheets) - 1 if after is None else names.index(str(after.Name))
    for s in reversed(sheets[:i + 1]):
        if s.Visible not in HIDDEN:
            return None, s
    for s in sheets[i + 1:]:
        if s.Visible not in HIDDEN:
            return s, None
    return None, sheets[i]


def _hide_unless_tabled(src) -> list[str]:
    """The tables in the opened file; none: its window hidden, so the copy does not show it."""
    tables = [f"{lo.Name} on {ws.Name}" for ws in src.Worksheets for lo in ws.ListObjects]
    if not tables:
        src.Windows(1).Visible = False
    return tables


def _break_links(wb, file_name: str) -> None:
    """Break the links the copy left to the temporary file, by its (unique) file name; the user's own links to other
    workbooks stay."""
    for link in wb.LinkSources(LINKS) or ():
        if re.split(r"[\\/]", str(link))[-1].casefold() == file_name.casefold():
            wb.BreakLink(link, LINKS)


def move_in(t, built: BuiltFile, after, first: bool = False) -> None:
    """Copy the built file's sheets into t.wb after the worksheet 'after' (None: after the last sheet; first: before
    the first visible sheet instead, where sheets.add_sheet puts a sheet when no sheet is there to follow), in one call.
    A name of the answer already in the workbook is NameTaken and a file holding a table a ValueError, both before
    anything is copied. The copied names are added to t.sheets_added (what a rollback deletes) just before the copy,
    so a copy that fails half-way is rolled back too. The temporary workbook is closed whatever happens; only the
    links to it are broken."""
    found = taken(t.wb, built.sheets, [spec.name for spec in built.tables])
    if found:
        raise NameTaken(_taken_text(found, jobs.com(lambda: str(t.wb.Name))))
    before, after_sheet = jobs.com(lambda: _anchor(t.wb, after, first))
    path, names = str(built.path), list(built.sheets)
    src = jobs.com(lambda: t.app.Workbooks.Open(path, 0, True), limit=120)       # no link update, read-only
    try:
        tables = jobs.com(lambda: _hide_unless_tabled(src))
        if tables:
            raise ValueError(f"the built file {built.path.name} holds an Excel Table ({', '.join(tables)}): a sheet "
                             "with a table is never copied between workbooks; the tables are made after the copy")
        t.sheets_added.extend(names)
        if before is not None:
            jobs.com(lambda: src.Worksheets(names).Copy(before))
        else:
            jobs.com(lambda: src.Worksheets(names).Copy(None, after_sheet))
    finally:
        try:
            jobs.com(lambda: src.Close(False), limit=30, check=False)
        except Exception as e:   # noqa: BLE001 - the answer is in; a failure on the way out stays the error
            log.warning("could not close the temporary workbook %s: %s", path, e)
    jobs.com(lambda: _break_links(t.wb, built.path.name))


def _add_table(wb, spec: TableSpec):
    ws = wb.Worksheets(spec.sheet)
    return ws.ListObjects.Add(SRC_RANGE, ws.Range(spec.ref), None, HAS_HEADERS)


def _table(t, spec: TableSpec) -> None:
    lo = jobs.com(lambda: _add_table(t.wb, spec))

    def look():
        lo.Name = spec.name
        lo.TableStyle = TABLE_STYLE
        if not spec.stripes:
            lo.ShowTableStyleRowStripes = False
    jobs.com(look)


def _bgr(colour: str) -> int:
    """'#RRGGBB' as a COM colour (red + 256 green + 65536 blue)."""
    r, g, b = (int(colour[i:i + 2], 16) for i in (1, 3, 5))
    return r + (g << 8) + (b << 16)


def _find_label(ws, fix: LabelFix) -> tuple:
    """The point's DataLabel, and its Top when the fix lifts it (read first, so a rerun sets the same value)."""
    lab = ws.ChartObjects(fix.chart).Chart.SeriesCollection(fix.series).Points(fix.point).DataLabel
    return lab, (lab.Top if fix.lift_pt else None)


def _label(ws, fix: LabelFix) -> None:
    lab, top = jobs.com(lambda: _find_label(ws, fix), limit=LABEL_WAIT_S)

    def apply():
        if fix.position:
            lab.Position = POSITION[fix.position]
        if fix.font_color:
            lab.Font.Color = _bgr(fix.font_color)
        if top is not None:
            lab.Top = top - fix.lift_pt
    jobs.com(apply, limit=LABEL_WAIT_S)


def finish(t, built: BuiltFile) -> None:
    """After move_in: each table (TableStyleLight1, row stripes off where the spec says), then each formula block
    with one Formula2 (its grid as it is: text keeps its leading apostrophe), then the chart label fixes on the
    Answer sheet. A label fix that fails is skipped with a log line: a label is cosmetic."""
    for spec in built.tables:
        _table(t, spec)
    for block in built.formulas:
        grid = tuple(tuple(row) for row in block.grid)
        jobs.com(lambda block=block, grid=grid: setattr(t.wb.Worksheets(block.sheet).Range(block.ref), "Formula2",
                                                        grid))
    if not built.label_fixes:
        return
    ws = jobs.com(lambda: t.wb.Worksheets(built.sheets[0]))
    for fix in built.label_fixes:
        try:
            _label(ws, fix)
        except jobs.JobAborted:
            raise
        except Exception as e:   # noqa: BLE001
            log.warning("chart label fix skipped (chart %s, series %s, point %s): %s", fix.chart, fix.series,
                        fix.point, e)


def calculate(t, sheets: list[str]) -> None:
    """Worksheet.Calculate per sheet, the Answer sheets last (they read Data and Lines); the other sheets in the
    order given."""
    for name in sorted(sheets, key=lambda s: s.endswith(" Answer")):
        jobs.com(lambda name=name: t.wb.Worksheets(name).Calculate())


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


def park(t, sheets: list[str], tables: list[str]) -> Parked:
    """Rename a question's sheets to 'old <name>' and its tables to '<name>_old' (a free name when that is taken),
    so the new answer can take their names. A name the workbook does not have is an InputError, before any rename.
    All or nothing: a park that fails puts back what it renamed."""
    have_sheets, have_tables = jobs.com(lambda: _inventory(t.wb))
    by_sheet = {n.casefold(): ws for n, ws in have_sheets}
    by_table = {n.casefold(): lo for n, lo, _ in have_tables}
    missing = ([f"the sheet '{n}'" for n in sheets if n.casefold() not in by_sheet]
               + [f"the table '{n}'" for n in tables if n.casefold() not in by_table])
    if missing:
        raise InputError(f"{' and '.join(missing)} to replace {'is' if len(missing) == 1 else 'are'} not in the "
                         "workbook any more: ask the question again as a new one")
    used_sheets, used_tables = set(by_sheet), set(by_table)
    parked = Parked()
    try:
        for name in sheets:
            ws, new = by_sheet[name.casefold()], _free("old " + name, used_sheets, SHEET_NAME_MAX)
            jobs.com(lambda ws=ws, new=new: setattr(ws, "Name", new))
            parked.sheets[name] = new
        for name in tables:
            lo, new = by_table[name.casefold()], _free(name + "_old", used_tables, sep="")
            jobs.com(lambda lo=lo, new=new: setattr(lo, "Name", new))
            parked.tables[name] = new
    except BaseException:
        try:
            restore(t, parked)
        except Exception as e:   # noqa: BLE001 - the park's own error is the one to raise
            log.warning("a failed park could not put every name back: %s", e)
        raise
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


def drop(t, parked: Parked) -> None:
    """After success: the parked sheets (and the tables on them) deleted in ONE call, DisplayAlerts off meanwhile."""
    names = list(parked.sheets.values())
    if not names:
        return
    prev = jobs.com(lambda: t.app.DisplayAlerts)
    jobs.com(lambda: setattr(t.app, "DisplayAlerts", False))
    try:
        jobs.com(lambda: t.wb.Worksheets(names).Delete())
    finally:
        jobs.com(lambda: setattr(t.app, "DisplayAlerts", prev), limit=30, check=False)
