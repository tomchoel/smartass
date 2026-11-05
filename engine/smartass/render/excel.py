# engine/smartass/render/excel.py
"""The COM writer: Q<n> Data as an Excel table (label columns as text first, so 'Sep 26' never becomes a date),
Q<n> Answer as one rectangle of values and formulas (Formula2), then v1's TN styling. Runs inside a user-Excel job
(v1 host().call(..., user_excel=True)); every COM call goes through v1 jobs.com (busy retry)."""
from __future__ import annotations

from datetime import datetime

from ..xl import styling
from ..xl.jobs import com

from ..model.errors import GuardError
from ..model.frame import Frame
from .datasheet import Header
from .layout import (CODE_COL, FIRST_COL, LABEL_COL, W_A, W_CODE, W_NUM, SheetPlan, col_letter,  # noqa: F401
                     col_number, label_width)  # col_number, label_width: re-exported (charts, tests)

XL_SRC_RANGE = 1
XL_YES = 1


def data_grid(frame: Frame) -> list[tuple]:
    return [frame.columns(), *frame.rows()]


def _put(ws, g: list[tuple], r0: int = 1):
    """Label columns as text first, then the values; returns the range. Only for a range that is not a table yet."""
    n, w = len(g), len(g[0])
    labels = ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + n - 1, w - 1))
    com(lambda: setattr(labels, "NumberFormat", "@"))
    rng = ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + n - 1, w))
    com(lambda: setattr(rng, "Value2", tuple(g)))
    return rng


def _create_table(ws, g: list[tuple], q: int, r0: int = 1):
    rng = _put(ws, g, r0)
    t = com(lambda: ws.ListObjects.Add(XL_SRC_RANGE, rng, None, XL_YES))
    com(lambda: setattr(t, "Name", f"Data_Q{q}"))
    return t


def write_data(ws, frame: Frame, q: int, header: Header | None = None) -> dict:
    """A new Q<n> Data sheet: the table at A<len(header)+2> (A1 without a header), then the header (its formulas
    name the table, so it must exist first), then the calm style. Returns where the cells plan 5 Task 11 and the
    layout check need sit: {table_row, sources: {query id: cell}, status: {query id: cell}}."""
    g = data_grid(frame)
    r0 = header.table_row if header else 1
    lo = _create_table(ws, g, q, r0)
    if header:
        _write_header(ws, header)
    _style(ws, lo, len(g), len(g[0]))
    return _cells(header, r0)


def _cells(header: Header | None, r0: int) -> dict:
    return {"table_row": r0, "sources": dict(header.sources) if header else {},
            "status": dict(header.status) if header else {}}


def _table(ws, q: int):
    try:
        return com(lambda: ws.ListObjects(f"Data_Q{q}"))
    except Exception:   # noqa: BLE001 - no such table (an older answer)
        return None


def _need_table(ws, q: int):
    lo = _table(ws, q)
    if lo is None:
        raise GuardError(f"table Data_Q{q} is missing on sheet 'Q{q} Data', so the answer cannot be redrawn in place",
                         "ask the question again with fin_analyze(spec); the old sheets were left as they were")
    return lo


def _origin(lo) -> tuple[int, int]:
    """The table's own top-left cell (its header row): the sheet may hold a provenance header above it."""
    head = com(lambda: lo.HeaderRowRange)
    return com(lambda: head.Row), com(lambda: head.Column)


def read_data(ws, q: int) -> list[tuple] | None:
    """Header + body of table Data_Q<q> as a grid (wherever the table sits), or None without such a table."""
    lo = _table(ws, q)
    if lo is None:
        return None
    v = com(lambda: lo.Range.Value2)
    return [tuple(r) for r in v] if isinstance(v, tuple) else None


def rewrite_grid(ws, g: list[tuple], q: int) -> None:
    """Put grid g into the EXISTING table Data_Q<q>, from the table's own origin (never A1): resize it, write the
    values, clear rows left below when it shrank. The table is never deleted, recreated or renamed, so formulas that
    point at it keep working. Header and body are written as two rectangles: Value2 on one rectangle covering a
    table's header AND body deletes the table in Excel (measured 04.11.2026, even with its own values) and every
    formula on it becomes #REF!. No table: refused (GuardError) before anything is touched - clearing the sheet and
    creating a new table would leave the answer's formulas, and its backup, pointing at a deleted table. Touches
    nothing above the table (the restore after a failed redraw calls it directly)."""
    lo = _need_table(ws, q)
    r0, c0 = _origin(lo)
    old_rows = com(lambda: lo.Range.Rows.Count)
    n, w = len(g), len(g[0])

    def cells(r1, c1, r2, c2):
        return ws.Range(ws.Cells(r0 + r1, c0 + c1), ws.Cells(r0 + r2, c0 + c2))
    labels = cells(0, 0, n - 1, w - 2)
    com(lambda: setattr(labels, "NumberFormat", "@"))              # text first: 'Sep 26' never becomes a date
    rng = cells(0, 0, max(n, 2) - 1, w - 1)                         # a table keeps at least one body row
    com(lambda: lo.Resize(rng))
    head = cells(0, 0, 0, w - 1)
    com(lambda: setattr(head, "Value2", (tuple(g[0]),)))
    if n > 1:
        body = cells(1, 0, n - 1, w - 1)
        com(lambda: setattr(body, "Value2", tuple(tuple(r) for r in g[1:])))
    if old_rows > n:
        left = cells(n, 0, old_rows - 1, w - 1)
        com(lambda: left.Clear())


def rewrite_data(ws, frame: Frame, q: int, header: Header | None = None) -> dict:
    """A redraw: the existing table moves (whole rows inserted or deleted above it) to A<len(header)+2>, its values
    are rewritten in place, then the header is written again (fetched time, control total). An old answer's table at
    A1 gets the header rows inserted above it: inserting whole rows keeps the table, its name and the answer's
    structured references."""
    lo = _need_table(ws, q)
    g = data_grid(frame)
    if header:
        _place(ws, lo, header.table_row)
    rewrite_grid(ws, g, q)
    if header:
        _write_header(ws, header)
    _style(ws, lo, len(g), len(g[0]))
    return _cells(header, _origin(lo)[0])


def _place(ws, lo, want: int) -> None:
    """Whole rows inserted (or, for a smaller header, deleted) above the table until its header row is `want`,
    then the rows above it cleared for the new header."""
    row, _ = _origin(lo)
    if row < want:
        rows = ws.Rows(f"1:{want - row}")
        com(lambda: rows.Insert())
    elif row > want:
        rows = ws.Rows(f"1:{row - want}")
        com(lambda: rows.Delete())
    above = ws.Rows(f"1:{want - 1}")
    com(lambda: above.Clear())


def _text(v):
    if isinstance(v, str) and v and not v.startswith("="):
        return "'" + v                   # text stays text: 'Sep 26' never becomes a date
    return "" if v is None else v


VALUE_FORMAT = '#,##0;(#,##0);"–"'
VALUE_BLUE = 0xC00000                    # BGR: blue, the FAST colour of imported numbers
LABEL_MAX, HEADER_MIN = 40, 16           # table columns fitted, at most 40; A:D at least 16 (header labels, totals)


def _write_header(ws, header: Header) -> None:
    """The label/value block in rows 1..len(header), columns A:D, as one rectangle; the title text in A1. The block
    gets the General format first (rows inserted above an old table may carry its '@', and a formula in a text cell
    stays text). The look: title bold navy, labels grey, values ink, the control totals as numbers."""
    h = len(header)
    g = [["", "", "", ""] for _ in range(h)]
    for i, (lab, value) in enumerate(header):
        if i == 0:
            g[0][0] = _text(value)
            continue
        g[i][0] = _text(lab)
        for j, x in enumerate(value if isinstance(value, tuple) else (value,)):
            g[i][1 + j] = _text(x)
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(h, 4))
    com(lambda: setattr(rng, "NumberFormat", "General"))
    com(lambda: setattr(rng, "Formula", tuple(tuple(r) for r in g)))
    numeric = [i + 1 for i, (lab, _) in enumerate(header) if lab in ("Control total", "Tolerance (NOK)")]

    def look():
        f = rng.Font
        f.Name = styling.tn.FONT
        f.Size = 9
        f.Color = styling._rgb(styling.tn.INK)
        title = ws.Cells(1, 1).Font
        title.Bold = True
        title.Size = 14
        title.Color = styling._rgb(styling.tn.NAVY)
        if h > 1:
            ws.Range(ws.Cells(2, 1), ws.Cells(h, 1)).Font.Color = styling._rgb(styling.tn.GREY)
        for r in numeric:
            rr = ws.Range(ws.Cells(r, 2), ws.Cells(r, 3))
            rr.NumberFormat = styling.local_format(rr, VALUE_FORMAT)
    com(look)


def _style(ws, lo, n: int, w: int) -> None:
    """The calm table look: TableStyleLight1 without banding and a bold header; Value in the blue import format,
    right-aligned; columns fitted to the table (not the header above it), at most 40, A:D at least 16; panes frozen
    below the table header (cosmetic: a failure is ignored)."""
    r0, c0 = _origin(lo)
    com(lambda: setattr(lo, "TableStyle", "TableStyleLight1"))
    com(lambda: setattr(lo, "ShowTableStyleRowStripes", False))
    com(lambda: setattr(lo.HeaderRowRange.Font, "Bold", True))
    if n > 1:
        def value():
            v = ws.Range(ws.Cells(r0 + 1, c0 + w - 1), ws.Cells(r0 + n - 1, c0 + w - 1))
            v.NumberFormat = styling.local_format(v, VALUE_FORMAT)
            v.Font.Color = VALUE_BLUE
            v.HorizontalAlignment = styling.tn.H_RIGHT
        com(value)
    for c in range(c0, c0 + w):
        col = ws.Range(ws.Cells(r0, c), ws.Cells(r0 + max(n, 2) - 1, c))
        com(lambda: col.Columns.AutoFit())
        width = com(lambda: col.ColumnWidth) or 0
        fit = min(width, LABEL_MAX)
        if c <= 4:
            fit = max(fit, HEADER_MIN)
        if fit != width:
            com(lambda: setattr(col.EntireColumn, "ColumnWidth", fit))
    try:
        styling._freeze(ws, r0, 0)
    except Exception:   # noqa: BLE001 - freezing is cosmetic
        pass


def answer_grid(p: SheetPlan) -> tuple[int, int, list[list]]:
    rows = [r for r, _ in p.cells]
    cols = [c for _, c in p.cells]
    r0, c0 = min(rows), min(cols)
    g = [[""] * (max(cols) - c0 + 1) for _ in range(max(rows) - r0 + 1)]
    for (r, c), v in p.cells.items():
        if isinstance(v, str) and v and not v.startswith("="):
            v = "'" + v                  # text stays text: Excel would turn 'Sep 26' into a date (as v1 answer_write)
        g[r - r0][c - c0] = "" if v is None else v
    return r0, c0, g


def write_answer(ws, p: SheetPlan, warnings: list[str]) -> list[str]:
    """Write every cell in one Formula2 rectangle, then style. Returns the styling lint findings."""
    r0, c0, g = answer_grid(p)
    rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(g) - 1, c0 + len(g[0]) - 1))
    com(lambda: setattr(rng, "Formula2", tuple(tuple(x) for x in g)))
    return styling.apply_report(ws, answer_roles(p))


def answer_roles(p: SheetPlan) -> dict:
    """The apply_report roles of an answer: widths (A 2, B code, C label, numbers 16, the layout's gap and marker
    widths), codes in B and labels in C."""
    widths = {"A": W_A, "B": W_CODE, "C": label_width(p),
              **{col_letter(c): W_NUM for c in range(FIRST_COL, col_number(p.last_col) + 1)}, **p.widths}
    code, label = col_letter(CODE_COL), col_letter(LABEL_COL)
    roles = {"last_col": p.last_col, "label_cols": [code, label], "code_cols": [code], "last_row": p.last_row,
             "col_widths": widths, "rows": p.rows}
    if p.helper_cols:                    # chart helpers: grey, inside the owned block, outside the print area
        roles |= {"helper_cols": list(p.helper_cols), "print_last_col": p.print_last_col}
    if p.footer:                         # the v2 print setup (styling.print_setup) up to the tables' last column
        roles |= {"print_footer": p.footer, "print_last_col": p.print_last_col or p.last_col}
    return roles


def answer_footer(fetched: datetime) -> str:
    """The answer's print footer: the source, the fetch date, the page number."""
    return f"Source: Oracle EPM Cloud · fetched {fetched:%Y-%m-%d} · page &P of &N"


def sheet_status(mismatches: list[str]) -> str:
    """The 'Sheet = model' verdict, as the status cell and the reply's checks show it."""
    return "OK" if not mismatches else f"DIFF in {len(mismatches)} cells"


def write_status(ws, p: SheetPlan, mismatches: list[str]) -> str:
    """After the readback: the verdict into the layout's empty 'Sheet = model' cell (text, never a formula)."""
    status = sheet_status(mismatches)
    if p.status_cell:
        rng = ws.Range(p.status_cell)
        com(lambda: setattr(rng, "Value2", status))
    return status


def read_block(ws, p: SheetPlan) -> tuple:
    rng = ws.Range(f"A1:{p.last_col}{p.last_row}")
    v = com(lambda: rng.Value2)
    return v if isinstance(v, tuple) else ((v,),)
