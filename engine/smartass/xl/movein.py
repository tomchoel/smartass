"""The built answer's sheets moved into the user's open workbook and finished there (the Excel side of the file render),
as operations recorded into batches (xl/inproc/batch.py) that run in one go inside Excel (xl/inproc/runner.py).

The file job's batches (pipeline._file_draw): look - the reads the build needs (record_book, record_screen,
record_base); bring in - the answer's table names checked against the workbook's defined names (record_free: a
lookup of those names, the batch's first operation), (a follow-up: park) the insert, the moves, then finish: the
tables, the formula blocks, the labels' position and colour (fix_labels), the calculation and the labels' lifts
(lift_labels), and the reads the close-out needs; close out - status, heights, (a follow-up: drop), Sources, spec,
sheet order. A function given a batch records into it; given none it builds and runs its own (read_book, base,
move_in, finish, calculate, park, drop: the tests and scripts/check_insert.py use them so). restore and land always
run their own: they clean up after a failure, also for an abandoned job.

move_in inserts every sheet of the temporary file (render/xlsx.build) with ONE Sheets.Add(Type:=<its path>): no
workbook is opened, nothing is copied or closed, and the insert leaves no link to the file (scripts/check_insert.py
proved it the same as the Open + Copy + Close it replaced, live 20.11.2026). Names are checked first: the insert
would silently rename a taken 'Q5 Data' to 'Q5 Data (2)' and every step by name would then hit the wrong sheet
(docs/audit/2026-11-19-render-map.md section 10 item 3). A sheet holding an Excel Table is never moved between
workbooks (Excel can crash), so the file holds none, and as the insert brings every sheet of the file, it must hold
the answer's sheets and only those, in order (its zip is checked for both before the insert); finish makes the
tables in the workbook, then writes the formula blocks (the formulas name the tables); the per-point chart labels the
file cannot carry (LabelFix) are positioned with the formulas and lifted once the answer is calculated (fix_labels,
lift_labels). park / restore / drop
rename a question's old sheets out of the way for a follow-up and delete or restore them afterwards.

Book is the job's picture of the workbook's sheets: read once (record_book) and kept up to date by the steps that add,
rename, move or delete sheets, so none of them reads the sheets again. Nothing comes back from a batch as a COM object:
every step addresses sheets and tables by name (a batch takes a sheet's node before it renames it).

A quiet batch turns off and puts back the settings itself (Calculation only in a batch that writes formulas); Excel
paints between two batches, so the job pauses PAINT_PAUSE after the batch that brings the answer in, when the workbook
is the active one of a visible Excel (Screen.shown): the user first sees the answer with its numbers and charts,
fully painted. Late-bound COM
ignores keyword arguments: every argument is positional, an omitted one MISSING."""
from __future__ import annotations

import logging
import re
import zipfile
from collections.abc import Callable
from dataclasses import dataclass, field
from xml.etree import ElementTree

from ..core import config
from ..render.xlsx import Base, BuiltFile
from ..render.xlsx_charts import LabelFix
from .host import HostError, InputError
from .inproc.batch import MISSING, Batch
from .jobs import book_name
from .inproc.runner import Result, run
from .sheetorder import ANSWER_TAB, DATA_TAB

log = logging.getLogger("smartview.movein")

HIDDEN = (0, 2)                         # xlSheetHidden, xlSheetVeryHidden
SRC_RANGE, HAS_HEADERS = 1, 1           # xlSrcRange, xlYes (ListObjects.Add)
TABLE_STYLE = "TableStyleLight1"
POSITION = {"right": -4152, "above": 0, "below": 1, "left": -4131}       # xlLabelPositionRight, Above, Below, Left
SHEET_NAME_MAX = 31
INSERT_WAIT_S = 120                     # the insert reads the file (as Workbooks.Open did, which had this limit)
STANDARD_PX = 64                        # XlsxWriter's standard column (Calibri 11), when the destination's is unknown
DATE_SETTINGS = (17, 32, 41, 42, 43)    # Application.International(i), 1-based: xlDateSeparator, xlDateOrder,
                                        # xlMonthLeadingZero, xlDayLeadingZero, xl4DigitYears
SHORT_DATE = 14                         # Excel's built-in short date (the system's)
PAINT_PAUSE = config.PAINT_PAUSE_S       # s (SV_PAINT_PAUSE, default 0.15): the owner's Excel paints a new sheet and
                                        # its chart whole in 0.15 s after the batch that brings it in (0.04: half
                                        # painted, measured by window captures 20.11.2026)
BOOK_PATHS = ["Name", "Visible", "Tab.Color", "ListObjects.*.Name"]


class NameTaken(InputError):
    """A sheet or table name of the answer already exists in the workbook."""


@dataclass
class Sheet:
    name: str
    visible: int = -1                    # Worksheet.Visible: -1 visible, 0 hidden, 2 very hidden
    tab: object = None                   # Tab.Color (False: none; None: Excel would not say)
    tables: list[str] = field(default_factory=list)      # the names of its tables


@dataclass
class Book:
    """The workbook's sheets in order, its chart sheets and the active sheet (when it is the active workbook)."""
    sheets: list[Sheet]
    charts: list[str] = field(default_factory=list)
    active: str | None = None

    def names(self) -> list[str]:
        return [s.name for s in self.sheets]

    def tables(self) -> list[tuple[str, str]]:
        """Every table as (name, sheet name), in workbook order."""
        return [(n, s.name) for s in self.sheets for n in s.tables]

    def order(self, names: list[str]) -> None:
        """The sheets in this order now (after moves)."""
        by = {s.name: s for s in self.sheets}
        self.sheets = [by[n] for n in names]


@dataclass
class Screen:
    """The Excel around the job, as look found it: active - the workbook is its Excel's active one (visible or not);
    shown - and that Excel is visible (the job pauses for a paint); user_book / user_sheet - the active workbook and
    sheet (landed on again after a failure); window - the workbook's first window is visible (a new Sources log's
    frozen header needs it)."""
    active: bool = False
    shown: bool = False
    user_book: str | None = None
    user_sheet: str | None = None
    window: bool = False


def _own(t, record, quiet: bool = False):
    """record(b, workbook node) -> reader, run in a batch of its own; what the reader gives."""
    b = Batch(quiet=quiet)
    reader = record(b, b.workbook(book_name(t)))
    return reader(run(t.app, b))


# ------------------------------------------------------------------ look
def record_book(b: Batch, wbn, key: str = "book") -> Callable[[Result], Book]:
    """The workbook's Book read in b: every worksheet's name, visibility, tab colour and tables (the verb each), the
    chart sheets' names, and the active sheet when the workbook is its Excel's active one."""
    b.ret(b.verb("each", wbn.Worksheets, BOOK_PATHS), f"{key}.sheets")
    b.ret(b.verb("each", wbn.Charts, ["Name"]), f"{key}.charts")
    b.ret(wbn.FullName, f"{key}.full name")
    with b.attempt():                                # no active workbook: no active sheet to know
        b.ret(wbn.ActiveSheet.Name, f"{key}.active")
        b.ret(b.app.ActiveWorkbook.FullName, f"{key}.active book")

    def book(res: Result) -> Book:
        v = res.values
        sheets = [Sheet(str(n), -1 if vis is None else int(vis), tab, [str(x) for x in tables or () if x is not None])
                  for n, vis, tab, tables in v[f"{key}.sheets"]]
        mine = v.get(f"{key}.active book") is not None and v.get(f"{key}.active book") == v[f"{key}.full name"]
        active = v.get(f"{key}.active") if mine else None
        return Book(sheets, [str(n) for (n,) in v[f"{key}.charts"] if n is not None],
                    None if active is None else str(active))
    return book


def read_book(t) -> Book:
    """The workbook's Book in a batch of its own (record_book)."""
    return _own(t, record_book)


def record_screen(b: Batch, wbn, key: str = "screen") -> Callable[[Result], Screen]:
    """The Excel around the job (Screen), read in b."""
    b.ret(wbn.FullName, f"{key}.full name")
    with b.attempt():                                # no active workbook, or no window: left at their defaults
        b.ret(b.app.Visible, f"{key}.visible")
        b.ret(b.app.ActiveWorkbook.FullName, f"{key}.active book")
        b.ret(b.app.ActiveWorkbook.Name, f"{key}.book")
        b.ret(b.app.ActiveSheet.Name, f"{key}.sheet")
        b.ret(wbn.Windows(1).Visible, f"{key}.window")

    def screen(res: Result) -> Screen:
        v = res.values
        active = v.get(f"{key}.active book") is not None and v.get(f"{key}.active book") == v[f"{key}.full name"]
        return Screen(active, active and bool(v.get(f"{key}.visible")), v.get(f"{key}.book"), v.get(f"{key}.sheet"),
                      bool(v.get(f"{key}.window")))
    return screen


def record_base(b: Batch, wbn, key: str = "base") -> Callable[[Result], Base]:
    """The destination's Normal font, standard sizes and date format (the base the file is built in, render/xlsx.Base),
    read in b: the first sheet's standard row height, and its last column's width in pixels when that column has the
    standard width (else XlsxWriter's 64 px, logged: the inserted widths are then off by the digit widths' ratio)."""
    font = wbn.Styles("Normal").Font
    ws = wbn.Worksheets(1)
    last = ws.Columns(ws.Columns.Count)
    for name, node in (("font", font.Name), ("size", font.Size), ("height", ws.StandardHeight),
                       ("standard", ws.StandardWidth), ("chars", last.ColumnWidth), ("points", last.Width)):
        b.ret(node, f"{key}.{name}")
    for i in DATE_SETTINGS:
        b.ret(b.app.International(i), f"{key}.international {i}")

    def base(res: Result) -> Base:
        v = {k[len(key) + 1:]: x for k, x in res.values.items() if k.startswith(key + ".")}
        size, chars, standard = float(v["size"]), float(v["chars"]), float(v["standard"])
        col_px = round(float(v["points"]) / 0.75)
        if chars != standard:
            log.warning("no standard column to measure (%s, standard %s characters): columns in %s px units",
                        chars, standard, STANDARD_PX)
            col_px = STANDARD_PX
        dates = (v[f"international {i}"] for i in DATE_SETTINGS)
        return Base(str(v["font"]), int(size) if size.is_integer() else size, round(float(v["height"]) / 0.75),
                    col_px, date_format(*dates))
    return base


def base(t) -> Base:
    """The destination's base in a batch of its own (record_base)."""
    return _own(t, record_base)


def date_format(sep: str, order, month_zero: bool, day_zero: bool, century: bool) -> str | int:
    """How Excel stores drill_xl's local date format 'dd.mm.yyyy' (set through NumberFormatLocal): as its built-in
    short date where that is the locale's short date (day first, '.' between, leading zeros, four-digit years: nb-NO;
    live, 19.11.2026), else as a code whose dots are the date separator ('dd/mm/yyyy') or literal."""
    if sep == "." and int(order) == 1 and month_zero and day_zero and century:
        return SHORT_DATE
    return "dd/mm/yyyy" if sep == "." else "dd.mm.yyyy"


def record_defined(b: Batch, wbn, tables: list[str], key: str = "defined") -> Callable[[Result], list[str]]:
    """Which of these table names the workbook has as defined names (a table cannot take one), read in b: a lookup
    of each (the verb exist), not a walk over every name - 18 s over COM for 2 000 names (re-review M5)."""
    b.ret(b.verb("exist", wbn.Names, list(tables), False), key)
    return lambda res: [str(n) for n in res.values[key]]


def record_free(b: Batch, wbn, tables: list[str]) -> int | None:
    """The check that none of these table names is a defined name of the workbook, recorded in b (the verb exist
    with refuse): one that is fails the batch right there, so it goes before anything the batch renames or inserts.
    Returns its index for name_taken (None: no names, nothing recorded)."""
    if not tables:
        return None
    b.verb("exist", wbn.Names, list(tables), True)
    return b.position - 1


def name_taken(e: BaseException, at: int | None, book: str) -> NameTaken | None:
    """The NameTaken of a batch that failed at the check record_free put at index at (the names Excel found, from
    the refusal's text, which both executors write the same: runner.TAKEN and the names); None for any other
    failure."""
    from .inproc.runner import TAKEN, BatchError
    if at is None or not (isinstance(e, BatchError) and e.index == at and TAKEN in e.message):
        return None
    names = e.message.rpartition(TAKEN)[2].split(", ")
    return NameTaken(_taken_text([f"the name '{n}'" for n in names], book))


def taken(book: Book, sheets: list[str], tables: list[str], names: list[str]) -> list[str]:
    """What of these names the workbook already has, each as it reads in a message ("the sheet 'Q5 Data'"): a sheet
    or chart sheet of an answer sheet's name, a table or defined name (names: record_defined) of a table's name
    (names compare without case, as in Excel; a table name is taken on any sheet)."""
    want_sheets, want_tables = {s.casefold() for s in sheets}, {n.casefold() for n in tables}
    return ([f"the sheet '{n}'" for n in book.names() if n.casefold() in want_sheets]
            + [f"the chart sheet '{n}'" for n in book.charts if n.casefold() in want_sheets]
            + [f"the table '{n}' on sheet '{s}'" for n, s in book.tables() if n.casefold() in want_tables]
            + [f"the name '{n}'" for n in names if n.casefold() in want_tables])


def _taken_text(found: list[str], book: str) -> str:
    if len(found) == 1:
        return f"{found[0]} already exists in {book}: rename or delete it, then ask again"
    return f"{', '.join(found[:-1])} and {found[-1]} already exist in {book}: rename or delete them, then ask again"


# ------------------------------------------------------------------ move in
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
_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def file_parts(path) -> tuple[list[str], list[str]]:
    """(the sheet names in the file's order, from xl/workbook.xml; the Excel Tables, from its xl/tables parts), in one
    read of the zip. A file render/xlsx built holds exactly its BuiltFile.sheets and no table."""
    with zipfile.ZipFile(path) as z:
        book = ElementTree.fromstring(z.read("xl/workbook.xml"))
        sheets = [s.get("name") for s in book.iter(_MAIN + "sheet")]
        parts = [n for n in z.namelist() if n.startswith("xl/tables/") and n.endswith(".xml")]
        return sheets, [m.group(1).decode() if (m := _TABLE_NAME.search(z.read(n))) else n for n in parts]


def move_in(t, built: BuiltFile, after: str | None, first: bool = False, book: Book | None = None,
            names: list[str] | None = None, b: Batch | None = None, free_at: int | None = None) -> int | None:
    """Insert the built file's sheets into t.wb after the sheet named after (None: after the last sheet; first:
    before the first visible sheet instead, where sheets.add_sheet puts a sheet when no sheet is there to follow)
    with ONE Sheets.Add(Type:=path). A name of the answer already in the workbook is NameTaken, and a file holding
    other sheets than built.sheets in that order (the insert brings every sheet of the file) or a table is a
    ValueError, all before anything is recorded. The names are added to t.sheets_added (what a rollback deletes)
    when the insert is recorded, so an insert that fails half-way is rolled back too. book: the workbook's sheets,
    told where the insert landed; names: which of the answer's table names are defined names (record_defined). Each
    is read in a batch of its own when not given - except the names when the insert is recorded into b: then the
    caller passes them, or records the check in b first (record_free, before anything it renames) and passes its
    index as free_at (pipeline._bring_in); neither is a ValueError, before anything is recorded. Recorded into b,
    the index of the insert operation there is returned (a failure before it inserted nothing); without b the insert
    runs in a quiet batch of its own (None)."""
    wanted = [spec.name for spec in built.tables]
    if b is not None and names is None and wanted and free_at is None:
        raise ValueError("move_in into a batch needs the defined names (names) or the index of the check recorded "
                         "before it (record_free's, as free_at)")
    if book is None or (names is None and b is None):
        rb = Batch(quiet=False)
        wbn = rb.workbook(book_name(t))
        get_book = record_book(rb, wbn) if book is None else None
        get_names = record_defined(rb, wbn, wanted) if names is None and b is None and wanted else None
        res = run(t.app, rb)
        book = book if book is not None else get_book(res)
        names = names if get_names is None else get_names(res)
    found = taken(book, built.sheets, wanted, names or [])
    if found:
        raise NameTaken(_taken_text(found, book_name(t)))
    sheets, tables = file_parts(built.path)
    if sheets != list(built.sheets):
        raise ValueError(f"the built file {built.path.name} holds the sheets {sheets}, not {list(built.sheets)}: the "
                         "insert brings every sheet of the file, so it must hold the answer's sheets and only them, "
                         "in that order")
    if tables:
        raise ValueError(f"the built file {built.path.name} holds an Excel Table ({', '.join(tables)}): a sheet "
                         "with a table is never moved between workbooks; the tables are made after the insert")
    before, after_sheet = _anchor(book, after, first)
    own = b is None
    b = Batch(calculation=False) if own else b
    wbn = b.workbook(book_name(t))
    path, new = str(built.path), list(built.sheets)
    t.sheets_added.extend(new)
    if before is not None:
        wbn.Sheets.Add(wbn.Worksheets(before.name), MISSING, MISSING, path)
    else:
        wbn.Sheets.Add(MISSING, wbn.Worksheets(after_sheet.name), MISSING, path)
    at = b.position - 1
    if own:
        run(t.app, b, limit=INSERT_WAIT_S)
    k = book.sheets.index(before) if before is not None else book.sheets.index(after_sheet) + 1
    book.sheets[k:k] = [Sheet(n, -1, ANSWER_TAB if i == 0 else DATA_TAB) for i, n in enumerate(new)]
    return None if own else at


# ------------------------------------------------------------------ finish
def _bgr(colour: str) -> int:
    """'#RRGGBB' as a COM colour (red + 256 green + 65536 blue)."""
    r, g, b = (int(colour[i:i + 2], 16) for i in (1, 3, 5))
    return r + (g << 8) + (b << 16)


def finish(t, built: BuiltFile, b: Batch | None = None) -> list[tuple]:
    """After move_in: each table (ListObjects.Add, its name, TableStyleLight1, row stripes off where the spec says),
    then every formula block (its grid as it is: text keeps its leading apostrophe), the chart labels' position and
    colour (fix_labels), the calculation (calculate), and the labels' lifts right after it (lift_labels) - so a label
    never shows unlifted between two paints. Returns the label operations' spans for label_skips. Without b in a
    quiet batch of its own (Calculation manual meanwhile: it writes the formulas), the same order."""
    own = b is None
    b = Batch() if own else b
    wbn = b.workbook(book_name(t))
    for spec in built.tables:
        ws = wbn.Worksheets(spec.sheet)
        lo = ws.ListObjects.Add(SRC_RANGE, ws.Range(spec.ref), MISSING, HAS_HEADERS)
        lo.Name = spec.name
        lo.TableStyle = TABLE_STYLE
        if not spec.stripes:
            lo.ShowTableStyleRowStripes = False
    for block in built.formulas:
        wbn.Worksheets(block.sheet).Range(block.ref).Formula2 = tuple(tuple(row) for row in block.grid)
    spans = fix_labels(b, wbn, built)
    calculate(t, built.sheets, b)
    spans += lift_labels(b, wbn, built)
    if own:
        label_skips(run(t.app, b), spans)
        return []
    return spans


def calculate(t, sheets: list[str], b: Batch | None = None) -> None:
    """Worksheet.Calculate per sheet, the Answer sheets last (they read Data and Lines); the other sheets in the order
    given. Without b in a batch of its own."""
    own = b is None
    b = Batch(calculation=False) if own else b
    wbn = b.workbook(book_name(t))
    for name in sorted(sheets, key=lambda s: s.endswith(" Answer")):
        wbn.Worksheets(name).Calculate()
    if own:
        run(t.app, b)


def _label(ws, fix: LabelFix):
    return ws.ChartObjects(fix.chart).Chart.SeriesCollection(fix.series).Points(fix.point).DataLabel


def fix_labels(b: Batch, wbn, built: BuiltFile) -> list[tuple]:
    """Each fix's position and colour recorded in b, before the calculation (a point label positioned after it lands
    12-24 pt elsewhere, measured in a hidden Excel 20.11.2026), every operation one that may fail (a label is
    cosmetic). Returns [(fix, first operation, end)] for label_skips."""
    spans = []
    if not built.label_fixes:
        return spans
    ws = wbn.Worksheets(built.sheets[0])
    with b.attempt():
        for fix in built.label_fixes:
            if not (fix.position or fix.font_color):
                continue
            start = b.position
            lab = _label(ws, fix)
            if fix.position:
                lab.Position = POSITION[fix.position]
            if fix.font_color:
                lab.Font.Color = _bgr(fix.font_color)
            spans.append((fix, start, b.position))
    return spans


def lift_labels(b: Batch, wbn, built: BuiltFile) -> list[tuple]:
    """Each lift recorded in b right after the calculation, inside Excel (the verb nudge: Top - lift, only for a
    label more than the lift below the chart's top - charts._labels leaves the empty points of a waterfall's other
    step series where they are). A lift set before the calculation is lost (the label ends where Excel puts it,
    measured in a hidden Excel 20.11.2026). Operations that may fail; returns [(fix, first operation, end)]."""
    spans = []
    lifts = [fix for fix in built.label_fixes if fix.lift_pt]
    if not lifts:
        return spans
    ws = wbn.Worksheets(built.sheets[0])
    with b.attempt():
        for fix in lifts:
            start = b.position
            b.verb("nudge", _label(ws, fix), "Top", -float(fix.lift_pt), float(fix.lift_pt))
            spans.append((fix, start, b.position))
    return spans


def label_skips(res: Result, spans: list[tuple]) -> None:
    """One log line per label fix with an operation Excel refused (skipped: cosmetic)."""
    for fix, start, end in spans:
        why = next((m for i, _, m in res.skipped if start <= i < end), None)
        if why is not None:
            log.warning("chart label fix skipped (chart %s, series %s, point %s): %s", fix.chart, fix.series,
                        fix.point, why)


# ------------------------------------------------------------------ park, restore, drop
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


def parking(book: Book, sheets: list[str], tables: list[str]) -> Parked:
    """The names a question's sheets ('old <name>') and tables ('<name>_old') are parked under (a free name when that
    is taken), so the new answer can take theirs. A name the workbook does not have is an InputError."""
    by_sheet = {s.name.casefold() for s in book.sheets}
    by_table = {n.casefold() for s in book.sheets for n in s.tables}
    missing = ([f"the sheet '{n}'" for n in sheets if n.casefold() not in by_sheet]
               + [f"the table '{n}'" for n in tables if n.casefold() not in by_table])
    if missing:
        raise InputError(f"{' and '.join(missing)} to replace {'is' if len(missing) == 1 else 'are'} not in the "
                         "workbook any more: ask the question again as a new one")
    parked = Parked()
    for name in sheets:
        parked.sheets[name] = _free("old " + name, by_sheet, SHEET_NAME_MAX)
    for name in tables:
        parked.tables[name] = _free(name + "_old", by_table, sep="")
    return parked


def record_park(b: Batch, wbn, parked: Parked, book: Book, tell: bool = True) -> None:
    """The renames of parked recorded in b, every sheet's and table's node taken before the first rename (a table is
    found on its sheet, from the Book); tell: the Book is told the new names at once (a batch that fails fails the
    job, whose clean-up reads the workbook)."""
    sheet_of = {n.casefold(): s.name for s in book.sheets for n in s.tables}
    nodes = [(wbn.Worksheets(n), new) for n, new in parked.sheets.items()]
    nodes += [(wbn.Worksheets(sheet_of[n.casefold()]).ListObjects(n), new) for n, new in parked.tables.items()]
    for node, new in nodes:
        node.Name = new
    if tell:
        _told(book, parked)


def _told(book: Book, parked: Parked) -> None:
    """The Book after the park: the sheets' and tables' new names, the active sheet's too."""
    by_sheet = {s.name.casefold(): s for s in book.sheets}
    for name, new in parked.sheets.items():
        by_sheet[name.casefold()].name = new
        if book.active is not None and book.active.casefold() == name.casefold():
            book.active = new                                    # the active sheet is the parked one now
    for name, new in parked.tables.items():
        for s in book.sheets:
            s.tables = [new if x.casefold() == name.casefold() else x for x in s.tables]


def park(t, sheets: list[str], tables: list[str], book: Book | None = None) -> Parked:
    """parking and record_park in a batch of its own, all or nothing: a park that fails puts back what it renamed
    (restore). book (else read): the workbook's sheets, told the new names."""
    book = book or read_book(t)
    parked = parking(book, sheets, tables)
    b = Batch(calculation=False)
    record_park(b, b.workbook(book_name(t)), parked, book, tell=False)
    try:
        run(t.app, b)
    except BaseException:
        try:
            restore(t, parked)
        except Exception as e:   # noqa: BLE001 - the park's own error is the one to raise
            log.warning("a failed park could not put every name back: %s", e)
        raise
    _told(book, parked)
    return parked


def restore(t, parked: Parked) -> None:
    """After a failure (the new sheets already removed): the parked sheets and tables get their names back, in two
    batches of their own (the workbook's sheets and tables read, then the renames), also for an abandoned job. Every
    name is tried; a HostError then names those that could not go back (their name is taken or the sheet is gone)."""
    rb = Batch(quiet=False)
    rb.ret(rb.verb("each", rb.workbook(book_name(t)).Worksheets, ["Name", "ListObjects.*.Name"]), "inventory")
    rows = run(t.app, rb, check=False).values["inventory"]
    have_sheets = {str(n).casefold(): str(n) for n, _ in rows}
    have_tables = {str(x).casefold(): str(n) for n, tables in rows for x in tables or () if x is not None}
    left: list[str] = []
    plan = []                                                    # (kind, parked name, old name, the table's sheet)
    for kind, items, have in (("sheet", parked.sheets, have_sheets), ("table", parked.tables, have_tables)):
        for old, new in items.items():
            there = new.casefold() in have
            if not there and old.casefold() in have:
                continue                                         # already back
            if not there or old.casefold() in have:
                left.append(f"{kind} '{new}' -> '{old}'")
                continue
            plan.append((kind, new, old, have_tables.get(new.casefold())))
    if plan:
        b = Batch(calculation=False)
        wbn = b.workbook(book_name(t))
        with b.attempt():                            # every name is tried; each node taken before the first rename
            taken_at = []
            for kind, new, _, sheet in plan:
                start = b.position
                node = wbn.Worksheets(new) if kind == "sheet" else wbn.Worksheets(sheet).ListObjects(new)
                taken_at.append((node, range(start, b.position)))
            spans = []
            for (node, took), (kind, new, old, _) in zip(taken_at, plan):
                start = b.position
                node.Name = old
                spans.append((took, range(start, b.position), kind, new, old))
        try:
            skipped = run(t.app, b, check=False).skipped
        except Exception as e:   # noqa: BLE001 - nothing could be renamed
            skipped = [(i, "", str(e)) for i in range(b.position)]
        for took, renamed, kind, new, old in spans:
            why = next((m for i, _, m in skipped if i in took or i in renamed), None)
            if why is not None:
                log.warning("could not rename %s '%s' back to '%s': %s", kind, new, old, why)
                left.append(f"{kind} '{new}' -> '{old}'")
    if left:
        raise HostError("could not put back the names of " + ", ".join(left))


def drop(t, parked: Parked, book: Book | None = None, b: Batch | None = None) -> int | None:
    """After success: the parked sheets (and the tables on them) deleted with ONE Delete (DisplayAlerts off: the batch
    is quiet). book: told they are gone. Recorded into b, the index of the Delete there is returned; without b in a
    quiet batch of its own (None)."""
    names = list(parked.sheets.values())
    if not names:
        return None
    own = b is None
    b = Batch(calculation=False) if own else b
    b.workbook(book_name(t)).Worksheets(names).Delete()
    at = b.position - 1
    if own:
        run(t.app, b)
    if book is not None:
        gone = {n.casefold() for n in names}
        book.sheets = [s for s in book.sheets if s.name.casefold() not in gone]
        if book.active is not None and book.active.casefold() in gone:
            book.active = None
    return None if own else at


def land(t, screen: Screen) -> None:
    """After a failure: the workbook and sheet the job found active are the active ones again (a batch of its own,
    also for an abandoned job; a sheet that is gone is skipped). Cosmetic: a failure is logged."""
    if not screen.user_book:
        return
    b = Batch(quiet=False)
    with b.attempt():
        wb = b.app.Workbooks(screen.user_book)
        wb.Activate()
        if screen.user_sheet:
            wb.Sheets(screen.user_sheet).Activate()
    try:
        res = run(t.app, b, check=False)
    except Exception as e:   # noqa: BLE001
        log.warning("could not re-activate the sheet the job found active: %s", e)
        return
    for _, _, why in res.skipped:
        log.warning("could not re-activate the sheet the job found active: %s", why)
