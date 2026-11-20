"""The v2 Sources sheet (sub-spec 5 §4): an append-only log with one row per query pull and per ledger-lines drill,
linked both ways with the Q<n> Data header, and a Known limits block to its right.

Layout (fixed: rows never move, so every 'row <r>' text and link target stays right):
  A1:N1    the log header (LOG_COLUMNS), frozen below, AutoFilter on A:N
  A2:N..   the log: append only. A row for a (Q, query id) that is already logged says 'row <r>' (the latest earlier
           row) in Supersedes; earlier rows are never edited. A drill (lines block) logs under its alias (b<i>),
           cube FinLite, route 'drill'.
  P1       'Known limits' (LIMITS_COL: two columns right of the log)
  P2..PK   one line per distinct warn/block caveat text of the workbook's answers (what is already there, merged
           with this answer's caveats; NO_LIMITS while there is none). Rewritten in place, column P only, and only
           when the answer adds a line.
Sheet name: 'Sources', or 'Sources SV' when 'Sources' is not a v2 log (the user's own sheet, or a v1 log - left as
it is; _kind tells them apart).

The furniture is set once, when the log is created: the header row, the whole columns' TN font 9 and the control
total's number format, fixed column widths (WIDTHS, Known limits LIMITS_WIDTH), the AutoFilter (Excel takes the rows
added below it in when it filters), the frozen header row, the print setup of A1:N1048575 (a finite last row:
whole columns are stored as $A:$N, which openpyxl cannot read) and the sheet-level name
FURNITURE_NAME that says all of it is there. A log of an earlier version (real hyperlinks, AutoFit widths, a print
area up to its last row then) gets the column formats, the print setup and the name once, on its next append; its
rows, links, widths and filter stay as they are.

In the file job (pipeline._file_draw) the log is read and written inside its batches: the kind of a 'Sources' sheet
in the look (record_kind: which sheet is the log), the log itself in the finish (record_read: its last row, the rows
already logged, Known limits, the furniture's name), and everything written in the close-out (record_write: a new
sheet, the furniture, the new rows as one Formula2 rectangle whose links (Data sheet, Supersedes) are HYPERLINK
formulas, Known limits when they grew, and the Data header's Sources cells), every operation one that may fail: the
log is provenance, not the answer. write_sources does the same in batches of its own."""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field

LIMITS_TITLE = "Known limits"
NO_LIMITS = "None recorded for the answers in this workbook."
OURS_A1 = ("Fetched",)                   # A1 of a Sources sheet smartass wrote (v1 and v2 logs; _kind tells them apart)
SEVERITIES = ("warn", "block")
LOG_COLUMNS = ("Fetched", "Q", "Query id", "Question", "Cube", "Scenario/Version", "POV", "Periods", "Route", "Rows",
               "Control total", "Status", "Data sheet", "Supersedes")
DATA_COL, SUPERSEDES_COL, QUESTION_COL, TOTAL_COL = 13, 14, 4, 11
LIMITS_COL = len(LOG_COLUMNS) + 2        # P: one empty column (O) between the log and Known limits
LIMITS_WIDTH = 80
QUESTION_MAX = 60
WIDTHS = (24.5, 3.4, 8.2, 45, 5.9, 25.5, 36, 23, 6.4, 6.3, 11.5, 6.9, 12, 10.9)   # A:N: the AutoFit widths of the logs
                                         # in the analyser folder (19.11.2026), D at most QUESTION_MAX
FURNITURE_NAME, FURNITURE_VERSION = "smartass_sources", "1"
SCAN = 500                               # Known limits lines read back at most
XL_UP = -4162
MAX_ROW = 1048576
PRINT_ROWS = MAX_ROW - 1                 # the print area's last row (A1:N1048575, not the whole columns)
V1_NOTE = ("the Sources sheet holds a v1 log, left as it is; the v2 log is on 'Sources SV'")


def _kind(a1, b1, used) -> str:
    """'v2' (A1:B1 = Fetched, Q), 'v1' (A1 Fetched, B1 'Question no'), 'empty' or 'user', from A1, B1 and the used
    range's address of a sheet named 'Sources'."""
    if a1 == "Fetched":
        return "v2" if b1 == "Q" else "v1"
    if a1 in (None, "") and used == "$A$1":
        return "empty"
    return "user"


def limits(caveats, old=()) -> list[str]:
    """The distinct warn/block caveat texts: the lines already on the sheet first, then the new ones (order kept)."""
    out = [x for x in old if x and x != NO_LIMITS]
    for c in caveats or ():
        text = str(c.get("text") or "").strip()
        if c.get("severity") in SEVERITIES and text:
            out.append(text)
    return list(dict.fromkeys(out))


def _text(v):
    return "'" + v if isinstance(v, str) and v else ("" if v is None else v)


def link(sheet: str, cell: str, text: str) -> str:
    """A link to <sheet>!<cell> of this workbook as a formula: =HYPERLINK("#'<sheet>'!<cell>","<text>"); an
    apostrophe in the sheet name is doubled (as in any sheet reference), a double quote in either text too (as in
    any formula string). The cell shows text."""
    target = "#'" + sheet.replace("'", "''") + "'!" + cell
    return '=HYPERLINK("' + target.replace('"', '""') + '","' + text.replace('"', '""') + '")'


def _col(c: int) -> str:
    from ..render.layout import col_letter
    return col_letter(c)


@dataclass
class Log:
    """What the reads tell of the log (record_kind, record_read)."""
    name: str                            # 'Sources' or 'Sources SV'
    exists: bool = False                 # the sheet is there (else record_write makes it)
    v1: bool = False                     # 'Sources' is a v1 log (the v2 log goes on 'Sources SV')
    header: bool = False                 # row 1 holds the log header (a log of this or an earlier version)
    furnished: bool = False              # this version's furniture is there (FURNITURE_NAME)
    filtered: bool = False               # the sheet has its AutoFilter
    last: int = 1                        # the last used row of column A
    logged: list = field(default_factory=list)      # [(row, Q, query id)] of the rows already logged
    limits: list = field(default_factory=list)      # the Known limits lines on the sheet
    titled: bool = False                 # P1 says Known limits
    data: str = ""                       # the Data sheet's name
    setup: str = ""                      # the print setup's key on the sheet (styling.PAGE_SETUP_NAME), "" for none


def record_kind(b, wbn, key: str = "sources kind") -> Callable:
    """The kind of the workbook's 'Sources' sheet (_kind), read in b; the reader gives it, or None when the sheet
    cannot be read (there is none)."""
    with b.attempt():
        ws = wbn.Worksheets("Sources")
        b.ret(ws.Range("A1:B1").Value2, f"{key}.a1")
        b.ret(ws.UsedRange.Address, f"{key}.used")

    def kind(res) -> str | None:
        v = res.values
        if f"{key}.a1" not in v:
            return None
        (a1, b1), = v[f"{key}.a1"]
        return _kind(a1, b1, v.get(f"{key}.used"))
    return kind


def log_of(kind: str | None, names, data: str) -> Log:
    """The log's sheet from the kind of 'Sources' (None: no sheet of that name) and the workbook's sheet names:
    'Sources', or 'Sources SV' when 'Sources' is not a v2 log (the user's own sheet, or a v1 log)."""
    log = Log("Sources SV" if kind in ("v1", "user") else "Sources", v1=kind == "v1", data=data)
    log.exists = log.name in names
    return log


def record_read(b, wbn, log: Log, key: str = "sources") -> Callable:
    """The log's own reads recorded in b (nothing for a log that is not there yet): row 1, the furniture's name, the
    AutoFilter, the last row, the rows already logged (Q and query id), Known limits, the print setup's name. The
    reader fills log in and gives it back."""
    from ..xl.styling import PAGE_SETUP_NAME
    if log.exists:
        p = _col(LIMITS_COL)
        ws = wbn.Worksheets(log.name)
        last = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        b.ret(ws.Range("A1:B1").Value2, f"{key}.head")
        b.ret(ws.AutoFilterMode, f"{key}.filtered")
        b.ret(last, f"{key}.last")
        b.ret(ws.Range(ws.Cells(2, 2), ws.Cells(last, 3)).Value2, f"{key}.logged")
        b.ret(ws.Range(f"{p}1").Value2, f"{key}.title")
        b.ret(ws.Range(f"{p}2:{p}{SCAN + 1}").Value2, f"{key}.limits")
        with b.attempt():                            # a name not there: an earlier version's log, or a new sheet
            b.ret(ws.Names(FURNITURE_NAME).RefersTo, f"{key}.furnished")
            b.ret(ws.Names(PAGE_SETUP_NAME).RefersTo, f"{key}.setup")

    def read(res) -> Log:
        if not log.exists:
            return log
        v = {k[len(key) + 1:]: x for k, x in res.values.items() if k.startswith(key + ".")}
        (a1, b1), = v["head"]
        log.header = (a1, b1) == ("Fetched", "Q")
        log.furnished = str(v.get("furnished")) == f'="{FURNITURE_VERSION}"'
        log.filtered = bool(v["filtered"])
        log.setup = str(v.get("setup") or "")
        log.last = max(int(v["last"]), 1)
        if log.last > 1:
            log.logged = [(2 + i, r[0], r[1]) for i, r in enumerate(v["logged"])]
        log.titled = v["title"] == LIMITS_TITLE
        if log.titled:
            for (x,) in v["limits"]:
                if x is None or x == "":
                    break
                log.limits.append(str(x))
        return log
    return read


def lines_info(ls, q: int) -> dict:
    """What one drill pulled, as the Sources log shows it: POV = the drilled cell, rows = the ledger lines, control
    total = the lines total (the AMOUNT column the sheet's G3 sums)."""
    from ..model.lines import AMOUNT
    cell = dict(ls.cell or {})
    period = cell.get("All Time Periods", "")
    pov = " · ".join(f"{k}: {v}" for k, v in cell.items() if k != "All Time Periods")
    total = sum(v for r in ls.rows if isinstance(v := r.get(AMOUNT), (int, float)) and not isinstance(v, bool))
    return {"scenario": cell.get("Scenario", ""), "pov": pov, "periods": period, "route": "drill",
            "rows": len(ls.rows), "total": float(total), "sheet": f"Q{q} Lines {ls.alias}"}


@dataclass
class Written:
    """What record_write recorded: rows - {query id or alias: Sources row}; span - the range of its operations (a
    skipped one there: a warning); setup - (the sheet's print setup key, the range of its operations) when the
    furniture set it up (styling.setup_name follows once none of them was skipped), else None; sheet - the log's."""
    rows: dict
    span: range
    setup: tuple | None = None
    sheet: str = ""                      # the log's sheet


def _furnish(b, wbn, ws, log: Log, separators: tuple, screen) -> tuple | None:
    """The furniture recorded (see the module doc): all of it on a new log, the column formats, print setup and name
    on a log of an earlier version (it keeps its header, widths, filter and frozen row). The header row is frozen only
    when the workbook's window is visible (styling._freeze), with the log's sheet active meanwhile (the workbook
    activated first when it is not the active one, and the user's workbook and sheet again after)."""
    from ..render.excel import SHEET_FOOTER, VALUE_FORMAT
    from . import styling
    tn = styling.tn
    new, width, p = not log.header, len(LOG_COLUMNS), _col(LIMITS_COL)
    last = _col(width)
    cols = ws.Range(f"A:{last}")
    for rng in (cols, ws.Columns(LIMITS_COL)):
        rng.Font.Name = tn.FONT                      # the log and Known limits: TN font 9, as the Data sheet
        rng.Font.Size = 9
    ws.Columns(TOTAL_COL).NumberFormat = tn.to_local(VALUE_FORMAT, *separators)    # the user's codes (styling.local_format)
    if new:
        ws.Range(ws.Cells(1, 1), ws.Cells(1, width)).Value = (tuple(LOG_COLUMNS),)
        for rng in (ws.Range(ws.Cells(1, 1), ws.Cells(1, width)), ws.Range(f"{p}1")):
            rng.Font.Bold = True
            rng.Font.Color = styling._rgb(styling.tn.NAVY)
        for c, w in enumerate(WIDTHS, 1):
            ws.Columns(c).ColumnWidth = w
        ws.Columns(LIMITS_COL).ColumnWidth = LIMITS_WIDTH
        if not log.filtered:                           # AutoFilter() without arguments toggles
            cols.AutoFilter()
    ws.Names.Add(FURNITURE_NAME, f'="{FURNITURE_VERSION}"', False)      # positional: Name, RefersTo, Visible
    if new and screen.window:                          # styling._freeze: the header row, on the log's own sheet
        if not screen.active:
            wbn.Activate()
        ws.Activate()
        win = wbn.Windows(1)
        for name, value in (("FreezePanes", False), ("ScrollRow", 1), ("ScrollColumn", 1),
                            ("DisplayGridlines", False), ("SplitColumn", 0), ("SplitRow", 1), ("FreezePanes", True)):
            setattr(win, name, value)
        if not screen.active and screen.user_book:
            user = b.app.Workbooks(screen.user_book)
            user.Activate()
            if screen.user_sheet:
                user.Sheets(screen.user_sheet).Activate()
    return styling.record_print_setup(b, ws, last, PRINT_ROWS, SHEET_FOOTER, log.setup)


def record_write(b, wbn, t, q: int, pre, log: Log, header, fetched: str, status: dict, lines_status: dict,
                 separators: tuple, screen, notes: list | None = None) -> Written:
    """One log row per query and per drill of the answer recorded into b, the links both ways (the Data header's
    Sources cell -> the row, the row's Data sheet -> 'Q<q> Data'!A1 or the lines sheet), Known limits merged; a new
    log's sheet (after the last sheet; its name added to t.sheets_added) and its furniture first (_furnish). log: as
    record_read gave it; header: the Data header (datasheet.Header; .sources holds each query's Sources cell);
    status / lines_status: query id / alias -> the calculated OK/DIFF; separators: the decimal and thousands
    separators (Application.International(3), (4)) for the total's number format; screen: movein.Screen; notes: a
    warning is appended there (a v1 log left as it is). Every operation may fail (Written.span)."""
    from ..render.datasheet import query_info
    from . import styling
    from .inproc.batch import MISSING
    start = b.position
    setup = None
    with b.attempt():
        if not log.exists:
            wss = wbn.Worksheets
            ws = wss.Add(MISSING, wbn.Worksheets(wss.Count))       # positional: after the last sheet
            ws.Name = log.name
            added = getattr(t, "sheets_added", None)
            if added is not None:
                added.append(log.name)
            if log.v1 and notes is not None:
                notes.append(V1_NOTE)
        else:
            ws = wbn.Worksheets(log.name)
        if not log.furnished:
            setup = _furnish(b, wbn, ws, log, separators, screen)
        kname, data, width = log.name, log.data, len(LOG_COLUMNS)
        entries = []                                   # (id, values after Q/id/question, sheet, data link cell)
        for query in pre.spec.queries:
            info = query_info(query, pre)
            entries.append((query.id, (query.cube, info["scenario"], info["pov"], info["periods"], info["route"],
                                       info["rows"], info["total"], status.get(query.id, "")), data,
                            header.sources.get(query.id)))
        for ls in (getattr(pre.frame, "lines", None) or {}).values():
            info = lines_info(ls, q)
            entries.append((ls.alias, ("FinLite", info["scenario"], info["pov"], info["periods"], info["route"],
                                       info["rows"], info["total"], lines_status.get(ls.alias, "")), info["sheet"],
                            None))
        rows: dict[str, int] = {}
        grid, backs = [], []
        for qid, vals, sheet, back in entries:
            r = log.last + 1 + len(rows)
            prev = [row for row, qq, x in log.logged if qq == f"Q{q}" and x == qid]
            row = [_text(x) for x in (fetched, f"Q{q}", qid, pre.spec.question, *vals)]
            row += [link(sheet, "A1", sheet), link(kname, f"A{prev[-1]}", f"row {prev[-1]}") if prev else ""]
            grid.append(tuple(row))
            if back:
                backs.append((back, link(kname, f"A{r}", f"{kname} row {r}")))
            rows[qid] = r
        lines = limits(pre.caveats, log.limits) or [NO_LIMITS]
        p = _col(LIMITS_COL)
        if grid:
            first, end = log.last + 1, log.last + len(grid)
            ws.Range(ws.Cells(first, 1), ws.Cells(end, width)).Formula2 = tuple(grid)
            links = ws.Range(ws.Cells(first, DATA_COL), ws.Cells(end, SUPERSEDES_COL))
            links.Font.Name = styling.tn.FONT              # a link takes the Hyperlink style (11 pt): back to the log's
            links.Font.Size = 9
        if not log.titled or lines != log.limits:
            n = max(len(lines), len(log.limits)) + 1     # a shorter list (the placeholder replaced) clears its tail
            text = [LIMITS_TITLE, *lines] + [""] * (n - len(lines) - 1)
            ws.Range(f"{p}1:{p}{n}").Value = tuple((_text(x),) for x in text)
        if backs:
            ws_data = wbn.Worksheets(data)               # a new sheet (the file holds no hyperlink to clear first)
            for cell, formula in backs:
                ws_data.Range(cell).Formula2 = formula
    return Written(rows, range(start, b.position), setup, log.name)


def write_sources(t, q: int, pre, ws_data, header, fetched: str, status: dict | None = None,
                  lines_status: dict | None = None, notes: list | None = None) -> dict[str, int]:
    """record_kind, record_read and record_write in batches of their own (one read, one quiet write, and the print
    setup's name when the furniture set it up): what the file job does inside its batches, for a caller outside it.
    ws_data: the Data sheet. Raises on any operation that failed. Returns {query id or alias: Sources row}."""
    from ..xl import movein
    from .inproc.batch import Batch
    from .inproc.runner import run
    from .host import HostError
    from .jobs import book_name, com
    from .styling import setup_name
    name = book_name(t)
    data = com(lambda: str(ws_data.Name))
    rb = Batch(quiet=False)
    wbn = rb.workbook(name)
    get_kind = record_kind(rb, wbn)
    get_book = movein.record_book(rb, wbn)
    get_screen = movein.record_screen(rb, wbn)
    rb.ret(rb.app.International(3), "decimal")
    rb.ret(rb.app.International(4), "thousands")
    res = run(t.app, rb)
    log = log_of(get_kind(res), get_book(res).names(), data)
    separators = (res.values["decimal"], res.values["thousands"])
    rb = Batch(quiet=False)
    log = record_read(rb, rb.workbook(name), log)(run(t.app, rb))
    b = Batch()
    written = record_write(b, b.workbook(name), t, q, pre, log, header, fetched, status or {}, lines_status or {},
                           separators, get_screen(res), notes)
    res = run(t.app, b)
    if res.skipped:
        raise HostError(f"the Sources log was not written: {res.skipped[0][2]}")
    if written.setup is not None:
        b = Batch(quiet=False)
        setup_name(b.workbook(name).Worksheets(log.name), written.setup[0])
        run(t.app, b)
    return written.rows
