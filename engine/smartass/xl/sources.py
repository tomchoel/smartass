"""The v2 Sources sheet (sub-spec 5 §4): an append-only log with one row per query pull and per ledger-lines drill,
linked both ways with the Q<n> Data header, and a Known limits block to its right.

Layout (fixed: rows never move, so every 'row <r>' text and hyperlink SubAddress stays right):
  A1:N1    the log header (LOG_COLUMNS), frozen below, AutoFilter on A:N
  A2:N..   the log: append only. A row for a (Q, query id) that is already logged says 'row <r>' (the latest earlier
           row) in Supersedes; earlier rows are never edited. A drill (lines block) logs under its alias (b<i>),
           cube FinLite, route 'drill'.
  P1       'Known limits' (LIMITS_COL: two columns right of the log)
  P2..PK   one line per distinct warn/block caveat text of the workbook's answers (what is already there, merged
           with this answer's caveats; NO_LIMITS while there is none). Rewritten in place, column P only.
Sheet name: 'Sources', or 'Sources SV' when 'Sources' is not a v2 log (the user's own sheet, or a v1 log - left as
it is; sheets._sources_name via log_kind)."""
from __future__ import annotations

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
SCAN = 500                               # Known limits lines read back at most
XL_UP = -4162
V1_NOTE = ("the Sources sheet holds a v1 log, left as it is; the v2 log is on 'Sources SV'")


def log_kind(ws) -> str:
    """'v2' (A1:B1 = Fetched, Q), 'v1' (A1 Fetched, B1 'Question no'), 'empty' or 'user' - one read."""
    from .jobs import com
    a1, b1 = com(lambda: ws.Range("A1:B1").Value)[0]
    if a1 == "Fetched":
        return "v2" if b1 == "Q" else "v1"
    if a1 in (None, "") and com(lambda: ws.UsedRange.Address) == "$A$1":
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


def _sheet(t, notes: list | None):
    from . import jobs, sheets
    name = sheets._sources_name(t)
    if not sheets._sheet_exists(t, name):
        ws = jobs.com(lambda: t.wb.Worksheets.Add(None, t.wb.Worksheets(t.wb.Worksheets.Count)))   # positional
        jobs.com(lambda: setattr(ws, "Name", name))
        added = getattr(t, "sheets_added", None)
        if added is not None:
            added.append(name)
        if name == "Sources SV" and notes is not None and log_kind(jobs.com(lambda: t.wb.Worksheets("Sources"))) == "v1":
            notes.append(V1_NOTE)
    return name, jobs.com(lambda: t.wb.Worksheets(name))


def _col(c: int) -> str:
    from ..render.layout import col_letter
    return col_letter(c)


def _read_limits(ws) -> list[str]:
    """The Known limits lines already on the sheet (P2 down to the first empty cell); none on a new sheet."""
    from .jobs import com
    p = _col(LIMITS_COL)
    if com(lambda: ws.Range(f"{p}1").Value) != LIMITS_TITLE:
        return []
    col = com(lambda: ws.Range(f"{p}2:{p}{SCAN + 1}").Value)
    lines = []
    for (v,) in col:
        if v is None or v == "":
            break
        lines.append(str(v))
    return lines


def _logged(ws) -> tuple[int, list[tuple]]:
    """(last used row of the log, [(row, Q, query id)] of the rows already logged)."""
    from .jobs import com
    last = max(com(lambda: ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row), 1)
    if last <= 1:
        return 1, []
    v = com(lambda: ws.Range(ws.Cells(2, 1), ws.Cells(last, 3)).Value)
    v = v if isinstance(v, tuple) and isinstance(v[0], tuple) else ((v,),)
    return last, [(2 + i, r[1], r[2]) for i, r in enumerate(v) if len(r) >= 3]


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


def write_sources(t, q: int, pre, ws_data, header, fetched: str, status: dict | None = None,
                  lines_status: dict | None = None, notes: list | None = None) -> dict[str, int]:
    """Known limits rewritten, one log row appended per query and per drill of the answer, the links both ways (the
    Data header's Sources cell -> the row, the row's Data sheet -> 'Q<q> Data'!A1 or the lines sheet). header: the
    Data header (datasheet.Header; .sources holds each query's Sources cell). status / lines_status: query id /
    alias -> the calculated OK/DIFF (Data header, lines sheet tie-out). notes: a warning is appended there (a v1 log
    left as it is). Returns {query id or alias: Sources row}."""
    from ..render.datasheet import query_info
    from ..render.excel import VALUE_FORMAT
    from . import styling
    from .jobs import com, fast_excel
    status, lines_status = status or {}, lines_status or {}
    kname, ws = _sheet(t, notes)
    data = com(lambda: ws_data.Name)
    width = len(LOG_COLUMNS)
    with fast_excel(t.app):
        old = _read_limits(ws)
        lines = limits(pre.caveats, old) or [NO_LIMITS]
        p = _col(LIMITS_COL)
        block = ws.Range(f"{p}1:{p}{max(len(old), len(lines)) + 1}")
        com(lambda: block.ClearContents())
        top = ws.Range(f"{p}1:{p}{len(lines) + 1}")
        com(lambda: setattr(top, "Value", tuple((_text(x),) for x in [LIMITS_TITLE, *lines])))
        head = ws.Range(ws.Cells(1, 1), ws.Cells(1, width))
        com(lambda: setattr(head, "Value", (tuple(LOG_COLUMNS),)))
        last, logged = _logged(ws)
        entries = []                                       # (id, values after Q/id/question, sheet, data link)
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
        out: dict[str, int] = {}
        for qid, vals, sheet, back in entries:
            r = last + 1 + len(out)
            prev = [row for row, qq, x in logged if qq == f"Q{q}" and x == qid]
            row = (fetched, f"Q{q}", qid, pre.spec.question, *vals, sheet, f"row {prev[-1]}" if prev else "")
            rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, width))
            com(lambda rng=rng, row=row: setattr(rng, "Value", (tuple(_text(x) for x in row),)))
            # positional (Anchor, Address, SubAddress, ScreenTip, TextToDisplay): keyword TextToDisplay= was ignored
            com(lambda r=r, s=sheet: ws.Hyperlinks.Add(ws.Cells(r, DATA_COL), "", f"'{s}'!A1", s, s))
            if back:
                com(lambda r=r, cell=back: ws_data.Hyperlinks.Add(ws_data.Range(cell), "", f"'{kname}'!A{r}",
                                                                  f"This pull on the {kname} sheet", f"{kname} row {r}"))
            out[qid] = r
        end = max(last + len(out), 2)

        def look():
            for rng in (head, ws.Range(f"{p}1")):
                rng.Font.Bold = True
                rng.Font.Color = styling._rgb(styling.tn.NAVY)
            tot = ws.Range(ws.Cells(2, TOTAL_COL), ws.Cells(end, TOTAL_COL))
            tot.NumberFormat = styling.local_format(tot, VALUE_FORMAT)
        com(look)
        if com(lambda: ws.AutoFilterMode):                 # AutoFilter() without arguments toggles: off, then on
            com(lambda: setattr(ws, "AutoFilterMode", False))
        log = ws.Range(ws.Cells(1, 1), ws.Cells(end, width))
        com(lambda: log.AutoFilter())
        com(lambda: log.Columns.AutoFit())
        q_col = ws.Columns(QUESTION_COL)
        if (com(lambda: q_col.ColumnWidth) or 0) > QUESTION_MAX:
            com(lambda: setattr(q_col, "ColumnWidth", QUESTION_MAX))
        k_col = ws.Columns(LIMITS_COL)
        com(lambda: setattr(k_col, "ColumnWidth", LIMITS_WIDTH))
    try:
        styling._freeze(ws, 1, 0)
    except Exception:   # noqa: BLE001 - freezing is cosmetic
        pass
    return out
