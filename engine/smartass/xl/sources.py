"""The v2 Sources sheet (sub-spec 5 §4): a Known limits block on top, then an append-only log with one row per query
pull, linked both ways with the Q<n> Data header.

Layout:
  1        'Known limits'
  2..K     one line per distinct warn/block caveat text of the workbook's answers (what is already there, merged
           with this answer's caveats; NO_LIMITS while there is none). Rewritten in place on every call.
  K+1      blank
  K+2      the log header (LOG_COLUMNS), frozen below, AutoFilter on
  K+3..    the log: append only. A row for a (Q, query id) that is already logged says 'row <r>' (the latest earlier
           row) in Supersedes; earlier rows are never edited.
When K grows, whole rows are inserted above the log header, so the log moves down as one piece. Hyperlink
SubAddresses and 'row <r>' texts written before such a move keep their old row number (Excel does not rewrite them).
Sheet name: 'Sources', or 'Sources SV' when the user's own sheet holds that name (sheets._sources_name)."""
from __future__ import annotations

LIMITS_TITLE = "Known limits"
NO_LIMITS = "None recorded for the answers in this workbook."
OURS_A1 = ("Fetched", LIMITS_TITLE)      # A1 of a Sources sheet smartass wrote (v1 log: 'Fetched'; v2: Known limits)
SEVERITIES = ("warn", "block")
LOG_COLUMNS = ("Fetched", "Q", "Query id", "Question", "Cube", "Scenario/Version", "POV", "Periods", "Route", "Rows",
               "Control total", "Status", "Data sheet", "Supersedes")
DATA_COL, SUPERSEDES_COL, QUESTION_COL = 13, 14, 4
QUESTION_MAX = 60
SCAN = 500                                 # Known limits lines read back at most
XL_UP = -4162


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


def _sheet(t):
    from . import jobs, sheets
    name = sheets._sources_name(t)
    if not sheets._sheet_exists(t, name):
        ws = jobs.com(lambda: t.wb.Worksheets.Add(None, t.wb.Worksheets(t.wb.Worksheets.Count)))   # positional
        jobs.com(lambda: setattr(ws, "Name", name))
        added = getattr(t, "sheets_added", None)
        if added is not None:
            added.append(name)
    return name, jobs.com(lambda: t.wb.Worksheets(name))


def _read_limits(ws) -> tuple[list[str], int | None]:
    """(the Known limits lines, the log header row) of an existing sheet; a v1 log (A1 'Fetched') has no limits and
    its header in row 1; an empty sheet: ([], None)."""
    from .jobs import com
    a1 = com(lambda: ws.Range("A1").Value)
    if a1 == "Fetched":
        return [], 1
    if a1 != LIMITS_TITLE:
        return [], None
    col = com(lambda: ws.Range(f"A2:A{SCAN + 1}").Value)
    lines = []
    for (v,) in col:
        if v is None or v == "":
            break
        lines.append(str(v))
    return lines, len(lines) + 3


def _logged(ws, hdr: int) -> tuple[int, list[tuple]]:
    """(last used row of the log, [(row, Q, query id)] of the rows already logged)."""
    from .jobs import com
    last = max(com(lambda: ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row), hdr)
    if last <= hdr:
        return last, []
    v = com(lambda: ws.Range(ws.Cells(hdr + 1, 1), ws.Cells(last, 3)).Value)
    v = v if isinstance(v, tuple) and isinstance(v[0], tuple) else ((v,),)
    return last, [(hdr + 1 + i, r[1], r[2]) for i, r in enumerate(v) if len(r) >= 3]


def write_sources(t, q: int, pre, ws_data, header, fetched: str, status: dict | None = None) -> dict[str, int]:
    """Known limits rewritten, one log row appended per query of the answer, the links both ways (the Data
    header's Sources cell -> the row, the row's Data sheet -> 'Q<q> Data'!A1). header: the Data header
    (datasheet.Header; .sources holds each query's Sources cell). status: query id -> 'OK'/'DIFF' as read from
    the Data header after calculating. Returns {query id: Sources row}."""
    from ..render.datasheet import query_info
    from ..render.excel import VALUE_FORMAT
    from . import styling
    from .jobs import com, fast_excel
    status = status or {}
    kname, ws = _sheet(t)
    data = com(lambda: ws_data.Name)
    with fast_excel(t.app):
        old, old_hdr = _read_limits(ws)
        lines = limits(pre.caveats, old) or [NO_LIMITS]
        hdr = len(lines) + 3
        if old_hdr is not None and hdr > old_hdr:          # the log moves down as one piece
            at = max(old_hdr - 1, 1)
            rows = ws.Rows(f"{at}:{at + hdr - old_hdr - 1}")
            com(lambda: rows.Insert())
        block = ws.Range(f"A1:A{hdr - 1}")
        com(lambda: block.ClearContents())
        top = ws.Range(f"A1:A{len(lines) + 1}")
        com(lambda: setattr(top, "Value", tuple((_text(x),) for x in [LIMITS_TITLE, *lines])))
        head = ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, len(LOG_COLUMNS)))
        com(lambda: setattr(head, "Value", (tuple(LOG_COLUMNS),)))
        last, logged = _logged(ws, hdr)
        out: dict[str, int] = {}
        for query in pre.spec.queries:
            r = last + 1 + len(out)
            info = query_info(query, pre)
            prev = [row for row, qq, qid in logged if qq == f"Q{q}" and qid == query.id]
            row = (fetched, f"Q{q}", query.id, pre.spec.question, query.cube, info["scenario"], info["pov"],
                   info["periods"], info["route"], info["rows"], info["total"], status.get(query.id, ""), data,
                   f"row {prev[-1]}" if prev else "")
            rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, len(LOG_COLUMNS)))
            com(lambda rng=rng, row=row: setattr(rng, "Value", (tuple(_text(x) for x in row),)))
            # positional (Anchor, Address, SubAddress, ScreenTip, TextToDisplay): keyword TextToDisplay= was ignored
            com(lambda r=r: ws.Hyperlinks.Add(ws.Cells(r, DATA_COL), "", f"'{data}'!A1", data, data))
            cell = header.sources.get(query.id)
            if cell:
                com(lambda r=r, cell=cell: ws_data.Hyperlinks.Add(ws_data.Range(cell), "", f"'{kname}'!A{r}",
                                                                  f"This pull on the {kname} sheet", f"{kname} row {r}"))
            out[query.id] = r
        end = max(last + len(out), hdr + 1)

        def look():
            ws.Range("A1").Font.Bold = True
            ws.Range("A1").Font.Color = styling._rgb(styling.tn.NAVY)
            head.Font.Bold = True
            head.Font.Color = styling._rgb(styling.tn.NAVY)
            ws.Range(ws.Cells(hdr + 1, 11), ws.Cells(end, 11)).NumberFormat = VALUE_FORMAT
        com(look)
        if com(lambda: ws.AutoFilterMode):                 # AutoFilter() without arguments toggles: off, then on
            com(lambda: setattr(ws, "AutoFilterMode", False))
        log = ws.Range(ws.Cells(hdr, 1), ws.Cells(end, len(LOG_COLUMNS)))
        com(lambda: log.AutoFilter())
        cols = ws.Range(ws.Cells(hdr, 1), ws.Cells(end, len(LOG_COLUMNS)))
        com(lambda: cols.Columns.AutoFit())
        q_col = ws.Columns(QUESTION_COL)
        if (com(lambda: q_col.ColumnWidth) or 0) > QUESTION_MAX:
            com(lambda: setattr(q_col, "ColumnWidth", QUESTION_MAX))
    try:
        styling._freeze(ws, hdr, 0)
    except Exception:   # noqa: BLE001 - freezing is cosmetic
        pass
    return out
