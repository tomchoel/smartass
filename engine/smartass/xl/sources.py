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
it is; log_kind tells them apart).

The furniture is set once, when the log is created: the header row, the whole columns' TN font 9 and the control
total's number format, fixed column widths (WIDTHS, Known limits LIMITS_WIDTH), the AutoFilter (Excel takes the rows
added below it in when it filters), the frozen header row, the print setup of A:N and the sheet-level name
FURNITURE_NAME that says all of it is there. A log of an earlier version (real hyperlinks, AutoFit widths, a print
area up to its last row then) gets the column formats, the print setup and the name once, on its next append; its
rows, links, widths and filter stay as they are. Per answer the log costs one read (the sheet, its last row, the
rows of this question, Known limits, the name) and one write: the new rows as one Formula2 rectangle whose links
(Data sheet, Supersedes) are HYPERLINK formulas, Known limits when they grew, and the Data header's Sources cells."""
from __future__ import annotations

from dataclasses import dataclass, field

LIMITS_TITLE = "Known limits"
NO_LIMITS = "None recorded for the answers in this workbook."
OURS_A1 = ("Fetched",)                   # A1 of a Sources sheet smartass wrote (v1 and v2 logs; log_kind tells them apart)
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
V1_NOTE = ("the Sources sheet holds a v1 log, left as it is; the v2 log is on 'Sources SV'")


def _kind(ws) -> str:
    """log_kind inside a COM call."""
    a1, b1 = ws.Range("A1:B1").Value[0]
    if a1 == "Fetched":
        return "v2" if b1 == "Q" else "v1"
    if a1 in (None, "") and ws.UsedRange.Address == "$A$1":
        return "empty"
    return "user"


def log_kind(ws) -> str:
    """'v2' (A1:B1 = Fetched, Q), 'v1' (A1 Fetched, B1 'Question no'), 'empty' or 'user' - one call."""
    from .jobs import com
    return com(lambda: _kind(ws))


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
    """What one read tells of the log (_read)."""
    name: str                            # 'Sources' or 'Sources SV'
    ws: object = None                    # None: the sheet is not there yet
    v1: bool = False                     # 'Sources' is a v1 log (the v2 log goes on 'Sources SV')
    header: bool = False                 # row 1 holds the log header (a log of this or an earlier version)
    furnished: bool = False              # this version's furniture is there (FURNITURE_NAME)
    last: int = 1                        # the last used row of column A
    logged: list = field(default_factory=list)      # [(row, Q, query id)] of the rows already logged
    limits: list = field(default_factory=list)      # the Known limits lines on the sheet
    titled: bool = False                 # P1 says Known limits
    data: str = ""                       # the Data sheet's name


def _busy_raised(e: BaseException) -> None:
    from .jobs import is_busy
    if is_busy(e):
        raise e


def _read(wb, ws_data) -> Log:
    """The whole read of write_sources, inside ONE COM call."""
    names = [str(s.Name) for s in wb.Worksheets]
    kind = _kind(wb.Worksheets("Sources")) if "Sources" in names else None
    log = Log("Sources SV" if kind in ("v1", "user") else "Sources", v1=kind == "v1",
              data=str(ws_data.Name) if ws_data is not None else "")
    if log.name not in names:
        return log
    ws = log.ws = wb.Worksheets(log.name)
    a1, b1 = ws.Range("A1:B1").Value[0]
    log.header = (a1, b1) == ("Fetched", "Q")
    try:
        log.furnished = str(ws.Names(FURNITURE_NAME).RefersTo) == f'="{FURNITURE_VERSION}"'
    except Exception as e:   # noqa: BLE001 - no such name: an earlier version's log, or a new sheet
        _busy_raised(e)
    log.last = max(int(ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row), 1)
    if log.last > 1:
        v = ws.Range(ws.Cells(2, 1), ws.Cells(log.last, 3)).Value
        v = v if isinstance(v, tuple) and isinstance(v[0], tuple) else ((v,),)
        log.logged = [(2 + i, r[1], r[2]) for i, r in enumerate(v) if len(r) >= 3]
    p = _col(LIMITS_COL)
    log.titled = ws.Range(f"{p}1").Value == LIMITS_TITLE
    if log.titled:
        for (x,) in ws.Range(f"{p}2:{p}{SCAN + 1}").Value:
            if x is None or x == "":
                break
            log.limits.append(str(x))
    return log


def _furnish(t, log: Log) -> None:
    """The furniture, set once (see the module doc): all of it on a new log, the column formats, print setup and
    name on a log of an earlier version (it keeps its header, widths, filter and frozen row)."""
    from ..render.excel import VALUE_FORMAT, sheet_print
    from . import styling
    from .jobs import com
    ws, new = log.ws, not log.header
    width, p = len(LOG_COLUMNS), _col(LIMITS_COL)
    last = _col(width)

    def furnish():
        cols = ws.Range(f"A:{last}")
        for rng in (cols, ws.Columns(LIMITS_COL)):
            rng.Font.Name = styling.tn.FONT              # the log and Known limits: TN font 9, as the Data sheet
            rng.Font.Size = 9
        tot = ws.Columns(TOTAL_COL)
        tot.NumberFormat = styling.local_format(tot, VALUE_FORMAT)
        if new:
            head = ws.Range(ws.Cells(1, 1), ws.Cells(1, width))
            head.Value = (tuple(LOG_COLUMNS),)
            for rng in (head, ws.Range(f"{p}1")):
                rng.Font.Bold = True
                rng.Font.Color = styling._rgb(styling.tn.NAVY)
            for c, w in enumerate(WIDTHS, 1):
                ws.Columns(c).ColumnWidth = w
            ws.Columns(LIMITS_COL).ColumnWidth = LIMITS_WIDTH
            if not ws.AutoFilterMode:                      # AutoFilter() without arguments toggles
                cols.AutoFilter()
        ws.Names.Add(FURNITURE_NAME, f'="{FURNITURE_VERSION}"', False)      # positional: Name, RefersTo, Visible
    com(furnish)
    if new:
        try:
            styling._freeze(ws, 1, 0)
        except Exception:   # noqa: BLE001 - freezing is cosmetic
            pass
    sheet_print(ws, last, MAX_ROW)             # the log A:N, whatever it grows to; Known limits stays on screen only


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


def _link_back(cell, formula: str) -> None:
    """A Data header Sources cell gets its HYPERLINK formula; a real hyperlink an earlier version put there is
    removed first (it would win the click), keeping the cell's alignment (Hyperlinks.Delete clears the formats)."""
    if cell.Hyperlinks.Count:
        align = cell.HorizontalAlignment
        cell.Hyperlinks.Delete()
        cell.HorizontalAlignment = align
    cell.Formula2 = formula


def write_sources(t, q: int, pre, ws_data, header, fetched: str, status: dict | None = None,
                  lines_status: dict | None = None, notes: list | None = None) -> dict[str, int]:
    """One log row appended per query and per drill of the answer, the links both ways (the Data header's Sources
    cell -> the row, the row's Data sheet -> 'Q<q> Data'!A1 or the lines sheet), Known limits merged. header: the
    Data header (datasheet.Header; .sources holds each query's Sources cell). status / lines_status: query id /
    alias -> the calculated OK/DIFF (Data header, lines sheet tie-out). notes: a warning is appended there (a v1 log
    left as it is). Into a log that is there: two COM calls (read, write); a new log also gets its furniture. The
    caller keeps Excel quiet around it. Returns {query id or alias: Sources row}."""
    from ..render.datasheet import query_info
    from . import styling
    from .jobs import com
    status, lines_status = status or {}, lines_status or {}
    log = com(lambda: _read(t.wb, ws_data))
    if log.ws is None:
        log.ws = com(lambda: t.wb.Worksheets.Add(None, t.wb.Worksheets(t.wb.Worksheets.Count)))   # positional
        com(lambda: setattr(log.ws, "Name", log.name))
        added = getattr(t, "sheets_added", None)
        if added is not None:
            added.append(log.name)
        if log.v1 and notes is not None:
            notes.append(V1_NOTE)
    if not log.furnished:
        _furnish(t, log)
    ws, kname, data, width = log.ws, log.name, log.data, len(LOG_COLUMNS)
    entries = []                                       # (id, values after Q/id/question, sheet, data link cell)
    for query in pre.spec.queries:
        info = query_info(query, pre)
        entries.append((query.id, (query.cube, info["scenario"], info["pov"], info["periods"], info["route"],
                                   info["rows"], info["total"], status.get(query.id, "")), data,
                        header.sources.get(query.id)))
    for ls in (getattr(pre.frame, "lines", None) or {}).values():
        info = lines_info(ls, q)
        entries.append((ls.alias, ("FinLite", info["scenario"], info["pov"], info["periods"], info["route"],
                                   info["rows"], info["total"], lines_status.get(ls.alias, "")), info["sheet"], None))
    out: dict[str, int] = {}
    grid, backs = [], []
    for qid, vals, sheet, back in entries:
        r = log.last + 1 + len(out)
        prev = [row for row, qq, x in log.logged if qq == f"Q{q}" and x == qid]
        row = [_text(x) for x in (fetched, f"Q{q}", qid, pre.spec.question, *vals)]
        row += [link(sheet, "A1", sheet), link(kname, f"A{prev[-1]}", f"row {prev[-1]}") if prev else ""]
        grid.append(tuple(row))
        if back:
            backs.append((back, link(kname, f"A{r}", f"{kname} row {r}")))
        out[qid] = r
    lines = limits(pre.caveats, log.limits) or [NO_LIMITS]
    grown = not log.titled or lines != log.limits
    p = _col(LIMITS_COL)

    def write():
        if grid:
            first, end = log.last + 1, log.last + len(grid)
            ws.Range(ws.Cells(first, 1), ws.Cells(end, width)).Formula2 = tuple(grid)
            links = ws.Range(ws.Cells(first, DATA_COL), ws.Cells(end, SUPERSEDES_COL))
            links.Font.Name = styling.tn.FONT              # a link takes the Hyperlink style (11 pt): back to the log's
            links.Font.Size = 9
        if grown:
            n = max(len(lines), len(log.limits)) + 1     # a shorter list (the placeholder replaced) clears its tail
            text = [LIMITS_TITLE, *lines] + [""] * (n - len(lines) - 1)
            ws.Range(f"{p}1:{p}{n}").Value = tuple((_text(x),) for x in text)
        for cell, formula in backs:
            _link_back(ws_data.Range(cell), formula)
    if grid or grown or backs:
        com(write)
    return out
