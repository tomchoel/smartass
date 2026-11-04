"""sv_drill in the user's workbook: the lines table ('Q<n> Lines', Excel table Lines_Q<n>) and the formulas
answer cells use over it. Formula rules (tn-formelstil): SUMIFS by default; an exact-match FILTER form where SUMIFS
would not match the key literally ('Jul-26' -> a date, digit strings over 15 characters -> a rounded number,
'0123' -> 123, * ? ~ -> wildcards, a leading < > = -> a comparison, '12:30' -> a time, '50%' -> 0.5, 'TRUE' -> a
logical, '1,5' -> a number in a Norwegian locale, over 255 characters -> SUMIFS fails)."""
from __future__ import annotations

import datetime as dt
import re

from openpyxl.utils import get_column_letter

from smartass.core import config
from smartass.epm.drill_lines import KEY_LABELS
from .jobs import com, fast_excel
from .sheets import add_sheet, unique_sheet_name
from .target import Target
from .host import InputError

_DATE_LIKE = re.compile(r"^\s*(\d{1,4}[-./]\d{1,2}([-./]\d{1,4})?|[A-Za-z]{3,9}[-./ ]\d{1,4}"
                        r"|\d{1,2}[-./ ][A-Za-z]{3,9}([-./ ]\d{1,4})?)\s*$")
_PATTERN_CHARS = set("*?~")
_COERCED = re.compile(r"^\s*(\d{1,2}:\d{2}.*|.*%\s*|true|false|-?\d+,\d+)\s*$", re.IGNORECASE)
AGGS = ("sum", "count", "debit", "credit")


def table_name(q: int, alias: str = "d") -> str:
    return f"Lines_Q{q}" if alias == "d" else f"Lines_Q{q}_{alias}"


def _lit(v) -> str:
    return '"' + str(v).replace('"', '""') + '"'


def safe_for_sumifs(v) -> bool:
    """True when a SUMIFS/COUNTIFS criterion "<v>" matches exactly the cells holding the text v."""
    s = str(v)
    if len(s) > 255 or _DATE_LIKE.match(s) or _COERCED.match(s) or _PATTERN_CHARS & set(s) \
            or s[:1] in ("<", ">", "="):
        return False
    return not (s.isdigit() and (len(s) > 15 or (len(s) > 1 and s[0] == "0")))


def drill_formula(table: str, where: dict[str, str | list[str]], agg: str = "sum") -> str:
    """Answer-cell formula over a lines table. where: column label -> exact value ('' = blank), or a list of values
    (any of them: an OR via ISNUMBER(MATCH(...)) in the FILTER form).
    agg: sum | count | debit (sum of negatives) | credit (sum of positives)."""
    if agg not in AGGS:
        raise ValueError(f"unknown drill agg '{agg}' (use one of {', '.join(AGGS)})")
    where = {k: "" if v is None else v for k, v in where.items()}     # None = blank, like specs._drill_value
    empty = [k for k, v in where.items() if isinstance(v, list) and not v]
    if empty:
        raise InputError(f"drill where on {', '.join(empty)}: an empty list of values matches nothing")
    amt = f"{table}[AMOUNT]"
    sign = {"debit": '"<0"', "credit": '">0"'}.get(agg)
    if all(not isinstance(v, list) and safe_for_sumifs(v) for v in where.values()):
        crit = "".join(f",{table}[{k}],{_lit(v)}" for k, v in where.items())
        if agg == "count":
            # ROWS = every line, like COUNTIFS and Python's count (a lines table is never written with 0 rows)
            return f"=COUNTIFS({crit[1:]})" if where else f"=ROWS({amt})"
        if sign:
            return f"=SUMIFS({amt},{amt},{sign}{crit})"
        return f"=SUMIFS({amt}{crit})" if where else f"=SUM({amt})"
    conds = [f"ISNUMBER(MATCH({table}[{k}],{{{','.join(_lit(x) for x in v)}}},0))" if isinstance(v, list)
             else f"{table}[{k}]={_lit(v)}" for k, v in where.items()]
    if sign:
        conds.append(f"{amt}{sign.strip(chr(34))}")
    test = conds[0] if len(conds) == 1 else "*".join(f"({c})" for c in conds)
    if agg == "count":
        return f"=SUM(--({test}))" if len(conds) == 1 else f"=SUM({test})"
    return f"=SUM(FILTER({amt},{test},0))"


# ------------------------------------------------------------------ the lines sheet
AMT_FMT = '#,##0_);(#,##0);"–"_)'
AMOUNT_LABELS = ("AMOUNT", "DEBET_AMOUNT", "CREDIT_AMOUNT", "TRANS_AMOUNT", "Orig_Curr_AMOUNT")
DATE_LABELS = ("GL_POSTED_DATE", "BOOKING_DATE")
TEXT_LABELS = {"ACCOUNT", "AP_INVOICE_NUMBER", "AP_VOUCHER_NUMBER", "GL_VOUCHER_NUMBER", "PROJECT", "LINE",
               "Period", "SYSTEM", "RESERVE", "FUTURE1"}
WIDTHS = {"GL_DESCRIPTION_LINE": 45, "AP_CUSTOMER_NAME": 30, "ACCOUNT_NAME": 34, "BATCH_NAME": 34,
          "JOURNAL_NAME": 34, "PRODUCT_DESC": 30, "CC_DESC": 28, "PROJECT_DESC": 28}
CHUNK = 5000
_EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def _cell(label, v):
    """Value as written to Excel. Every text value gets a leading apostrophe, so Excel stores it literally: no
    number/date/time/percent/logical coercion ('43200', 'Jul-26', 'TRUE', '1,5') and no formula injection ('=...').
    Key columns (TEXT_LABELS) are text even if they arrive as numbers. Excel hides the apostrophe.
    Dates go as Excel serial numbers: pywin32 shifts a naive datetime by the UTC offset (10:25 arrived as 08:25,
    15.10.2026), which can move a posting near midnight to the previous day."""
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return (v.replace(tzinfo=None) - _EXCEL_EPOCH) / dt.timedelta(days=1)
    if isinstance(v, str) or label in TEXT_LABELS:
        return "'" + str(v)
    return v


def _local_formats(app) -> tuple[str, str]:
    """(amount, date) number formats in the user's Excel locale. Formats are set through NumberFormatLocal like
    styling.py does: through late-bound COM, NumberFormat is read in the user's locale too, so '#,##0' came out as
    '9422,980' and 'yyyy' as a literal in nb-NO (measured 15.10.2026)."""
    from .styling import tn
    intl = com(lambda: app.International)
    dec, thou, year, month, day = intl[2], intl[3], intl[18], intl[19], intl[20]   # xlDecimalSeparator ... xlDayCode
    return tn.to_local(AMT_FMT, dec, thou), f"{day * 2}.{month * 2}.{year * 4}"


def missing_columns(table_cols: list[str] | None, where: dict) -> list[str]:
    """where-columns that the lines table does not have (e.g. a detail column left out above DRILL_XL_FULL lines).
    table_cols None = table not found: nothing to check here."""
    if table_cols is None:
        return []
    have = {str(c).lower() for c in table_cols}
    return [k for k in where if str(k).lower() not in have]


def find_table(wb, tname: str):
    """(sheet, ListObject) of the table named tname anywhere in the workbook, else (None, None)."""
    for ws in wb.Worksheets:
        for lo in ws.ListObjects:
            if str(lo.Name).lower() == tname.lower():
                return ws, lo
    return None, None


def table_columns(wb, tname: str) -> list[str] | None:
    """The table's header; None only when no table has that name (a COM error propagates: 'cannot check')."""
    _, lo = com(lambda: find_table(wb, tname))
    if lo is None:
        return None
    return [str(v) for v in com(lambda: lo.HeaderRowRange.Value)[0]]


def lines_tables(wb) -> list[str]:
    """Names of the drill tables (Lines_*) in the workbook."""
    return com(lambda: [str(lo.Name) for ws in wb.Worksheets for lo in ws.ListObjects
                        if str(lo.Name).lower().startswith("lines_")])


def _stale_ranges(hdr: int, old_end: int, old_ncol: int, new_end: int, new_ncol: int) -> list[str]:
    """Cells of the old table outside the new one (a re-drill with fewer rows or columns): cleared after Resize."""
    out = []
    if old_end > new_end:
        out.append(f"A{new_end + 1}:{get_column_letter(max(old_ncol, new_ncol))}{old_end}")
    if old_ncol > new_ncol:
        out.append(f"{get_column_letter(new_ncol + 1)}{hdr}:{get_column_letter(old_ncol)}{new_end}")
    return out


def _source_text(source: str) -> str:
    return (f"{source} Extract {dt.datetime.now():%d.%m.%Y %H:%M}. AMOUNT has the FinMax sign (cost negative). "
            "Snapshot: run sv_drill again to update.")


def same_header(old, new) -> bool:
    """A table is reused in place only when its header is exactly the new column list (else formulas would point
    at the wrong or missing columns)."""
    return [str(x) for x in (old or [])] == [str(x) for x in new]


def _col_runs(values: list) -> list[tuple]:
    """[v per column] -> [(v, first column, last column)] (1-based), adjacent equal values joined."""
    out: list[list] = []
    for j, v in enumerate(values, 1):
        if out and out[-1][0] == v:
            out[-1][2] = j
        else:
            out.append([v, j, j])
    return [tuple(x) for x in out]


def _clear_body(lo) -> None:
    body = lo.DataBodyRange
    if body is not None:
        body.ClearContents()


def _title_fonts(ws) -> None:
    ws.Range("A2").Font.Color = 0
    ws.Range("A1").Font.Bold = True


def _name_table(lo, tname: str) -> None:
    lo.Name = tname
    lo.TableStyle = "TableStyleLight1"


def _freeze_below(t: Target, ws, hdr: int) -> None:
    """Freeze the panes below the header row, in one COM call (rerun as a whole when Excel is busy)."""
    t.wb.Activate()                          # ActiveWindow must be OUR workbook's window
    ws.Activate()
    t.app.ActiveWindow.FreezePanes = False
    ws.Range(f"A{hdr + 1}").Select()
    t.app.ActiveWindow.FreezePanes = True


def write_lines(t: Target, q: int, rows: list[dict], columns: list[str], title: str, source: str,
                alias: str = "d") -> dict:
    """'Q<n> Lines' (+ ' <alias>'): title, source, then the Excel table Lines_Q<n>. Above DRILL_XL_FULL lines only
    the key columns go to Excel (full detail stays in drill.db). A re-drill of the same question/alias with the SAME
    columns reuses the table in place (cleared, rewritten, ListObject.Resize), so the answer's formulas stay valid;
    with other columns the old table is removed and a new one with the same name is built (columns_changed).
    Returns {sheet, table, columns, rows, reused, columns_changed}."""
    if not rows:
        raise ValueError("write_lines needs at least one line (0 lines: use clear_lines)")
    cols = columns if len(rows) <= config.DRILL_XL_FULL else [c for c in columns if c in KEY_LABELS]
    tname = table_name(q, alias)
    last_col = get_column_letter(len(cols))
    amt_fmt, date_fmt = _local_formats(t.app)
    ws, lo = com(lambda: find_table(t.wb, tname))
    reused = changed = False
    with fast_excel(t.app):
        if lo is not None:
            name, hdr, nrow, old_ncol, header = com(lambda: (ws.Name, lo.HeaderRowRange.Row, lo.Range.Rows.Count,
                                                             lo.Range.Columns.Count, lo.HeaderRowRange.Value))
            old_end = hdr + nrow - 1
            if same_header(header[0], cols):
                reused = True
                com(lambda: _clear_body(lo))
            else:                                   # other columns: drop the old table, build anew (same name)
                changed = True
                com(lo.Delete)
                com(lambda: ws.Range(f"A{hdr}:{get_column_letter(old_ncol)}{old_end}").Clear())
                lo = None
        else:
            name = unique_sheet_name(t, f"Q{q} Lines" + ("" if alias == "d" else f" {alias}"))
            ws = add_sheet(t, name, after_q=q, color=0xA6A6A6)
            hdr = 4
        end = hdr + len(rows)
        com(lambda: setattr(ws.Range("A1:A2"), "Value", ((title,), (_source_text(source),))))
        com(lambda: _title_fonts(ws))
        com(lambda: setattr(ws.Range(f"A{hdr}:{last_col}{hdr}"), "Value", (tuple(cols),)))
        if reused:
            com(lambda: lo.Resize(ws.Range(f"A{hdr}:{last_col}{end}")))
            for addr in _stale_ranges(hdr, old_end, old_ncol, end, len(cols)):
                com(lambda a=addr: ws.Range(a).Clear())
        for i in range(0, len(rows), CHUNK):
            block = tuple(tuple(_cell(c, r.get(c)) for c in cols) for r in rows[i:i + CHUNK])
            top = hdr + 1 + i
            com(lambda b=block, tp=top: setattr(ws.Range(f"A{tp}:{last_col}{tp + len(b) - 1}"), "Value", b),
                limit=180)
        if not reused:
            lo = com(lambda: ws.ListObjects.Add(1, ws.Range(f"A{hdr}:{last_col}{end}"), None, 1))
            com(lambda: _name_table(lo, tname))
        fmts = [amt_fmt if c in AMOUNT_LABELS else date_fmt if c in DATE_LABELS else None for c in cols]
        for fmt, j1, j2 in _col_runs(fmts):     # adjacent columns with the same format / width: one range call
            if fmt is not None:
                rng = f"{get_column_letter(j1)}{hdr + 1}:{get_column_letter(j2)}{end}"
                com(lambda a=rng, f=fmt: setattr(ws.Range(a), "NumberFormatLocal", f))
        for w, j1, j2 in _col_runs([WIDTHS.get(c, 14) for c in cols]):
            rng = f"{get_column_letter(j1)}1:{get_column_letter(j2)}1"
            com(lambda a=rng, w=w: setattr(ws.Range(a).EntireColumn, "ColumnWidth", w))
        com(lambda: _freeze_below(t, ws, hdr))
    return {"sheet": name, "table": tname, "columns": cols, "rows": len(rows), "reused": reused,
            "columns_changed": changed}


def clear_lines(t: Target, q: int, title: str, source: str, alias: str = "d") -> dict | None:
    """A re-drill that returned 0 lines: empty the existing table's body (the table and its columns stay, so the
    answer's formulas give 0 and Excel matches drill.db). None when there is no table."""
    tname = table_name(q, alias)
    ws, lo = com(lambda: find_table(t.wb, tname))
    if lo is None:
        return None
    with fast_excel(t.app):
        body = com(lambda: lo.DataBodyRange)
        if body is not None:
            com(body.Delete)
        com(lambda: setattr(ws.Range("A1"), "Value", title))
        com(lambda: setattr(ws.Range("A2"), "Value", _source_text(source) + " 0 lines."))
    return {"sheet": com(lambda: ws.Name), "table": tname, "columns": [], "rows": 0, "reused": True,
            "columns_changed": False}


def mark_incomplete(t: Target, q: int, alias: str = "d") -> None:
    """Last resort after a failed re-drill whose old lines could not be restored: a red warning in A2."""
    ws, lo = com(lambda: find_table(t.wb, table_name(q, alias)))
    if ws is None:
        return
    com(lambda: setattr(ws.Range("A2"), "Value", "INCOMPLETE – run sv_drill again"))
    com(lambda: setattr(ws.Range("A2").Font, "Color", 0x0000C0))     # BGR: red
