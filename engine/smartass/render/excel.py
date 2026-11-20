# engine/smartass/render/excel.py
"""What the answer's sheets share with Excel: the Data grid, the Answer's cells as one Formula2-ready rectangle (text
with a leading apostrophe, so 'Sep 26' never becomes a date) and its apply_report roles, the shared formats and
footers, a data sheet's page setup (styling.print_setup), and after the move-in the read-back block and the 'Sheet =
model' status cell. The sheets themselves are built as a file (render/xlsx.py) and moved in (xl/movein.py); the
read-back and the status are recorded into the job's batches (xl/inproc)."""
from __future__ import annotations

from datetime import datetime

from ..xl import styling

from ..model.frame import Frame
from .layout import (CODE_COL, FIRST_COL, LABEL_COL, W_A, W_CODE, W_NUM, SheetPlan, col_letter,  # noqa: F401
                     col_number, label_width)  # col_number, label_width: re-exported (charts, tests)


def data_grid(frame: Frame) -> list[tuple]:
    return [frame.columns(), *frame.rows()]


def _text(v):
    if isinstance(v, str) and v and not v.startswith("="):
        return "'" + v                   # text stays text: 'Sep 26' never becomes a date
    return "" if v is None else v


VALUE_FORMAT = '#,##0;(#,##0);"–"'
LABEL_MAX, HEADER_MIN = 40, 16           # table columns fitted, at most 40; A:D at least 16 (header labels, totals)
HEADER_COLS = 4                          # the Data header block is A:D
SHEET_FOOTER = "Source: Oracle EPM Cloud · &A · page &P of &N"     # Data, Lines and Sources sheets (&A: sheet name)


def sheet_print(ws, last_col: str, last_row: int, header_row: int | None = None) -> None:
    """A data sheet's page setup (styling.print_setup): A1:<last_col><last_row>, landscape, one page wide, the sheet
    name and page number in the footer. header_row (a Lines sheet: 40+ columns are unreadable fitted to one page):
    100 % zoom, as many pages wide as needed, that table header row repeated on every page. Cosmetic: a failure is
    ignored."""
    try:
        if header_row:
            styling.print_setup(ws, last_col, last_row, SHEET_FOOTER, fit_wide=False,
                                title_rows=f"${header_row}:${header_row}")
            return
        styling.print_setup(ws, last_col, last_row, SHEET_FOOTER)
    except Exception:   # noqa: BLE001 - no printer, a busy Excel: the data is written
        pass


def answer_grid(p: SheetPlan) -> tuple[int, int, list[list]]:
    rows = [r for r, _ in p.cells]
    cols = [c for _, c in p.cells]
    r0, c0 = min(rows), min(cols)
    g = [[""] * (max(cols) - c0 + 1) for _ in range(max(rows) - r0 + 1)]
    for (r, c), v in p.cells.items():
        if isinstance(v, str) and v and ((r, c) in p.literal or not v.startswith("=")):
            v = "'" + v                  # text stays text: Excel would turn 'Sep 26' into a date (as v1 answer_write)
        g[r - r0][c - c0] = "" if v is None else v
    return r0, c0, g


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
        roles |= {"print_footer": p.footer, "print_last_col": p.print_last_col or p.last_col,
                  "print_area_col": p.print_area_col or p.print_last_col or p.last_col,   # charts included
                  "keep": keep_spans(p)}
    return roles


def keep_spans(p: SheetPlan) -> list[tuple[int, int]]:
    """The row spans a page break must not cut: each section from its heading to its last non-blank row, and the
    charts zone (a chart cut in two prints as two halves)."""
    kinds = {r["row"]: r.get("kind") for r in p.rows}
    zones = dict(p.zones)
    heads = sorted(r for r, k in kinds.items() if k == "section")
    marks = sorted(set(heads) | set(zones.values()) | {p.last_row + 1})
    filled = {r for (r, _c) in p.cells} | {r for r, k in kinds.items() if k not in (None, "blank")}
    out = []
    for s in heads:
        nxt = next((m for m in marks if m > s), p.last_row + 1)
        e = nxt - 1
        while e > s and e not in filled:
            e -= 1
        out.append((s, e))
    from .layout import CHART_ROWS             # each row of (two) charts: a page may fall between rows, never through
    for r in sorted({c.anchor_row for c in p.charts if getattr(c, "anchor_row", None)}):
        out.append((r, min(r + CHART_ROWS - 1, p.last_row)))
    return sorted(out)


def answer_footer(fetched: datetime) -> str:
    """The answer's print footer: the source, the fetch date, the page number."""
    return f"Source: Oracle EPM Cloud · fetched {fetched:%Y-%m-%d} · page &P of &N"


def sheet_status(mismatches: list[str]) -> str:
    """The 'Sheet = model' verdict, as the status cell and the reply's checks show it."""
    return "OK" if not mismatches else f"DIFF in {len(mismatches)} cells"


def write_status(ws, p: SheetPlan, mismatches: list[str]) -> str:
    """After the readback: the verdict into the layout's empty 'Sheet = model' cell (text, never a formula); ws: the
    Answer sheet's node in the close-out batch (the set is recorded)."""
    status = sheet_status(mismatches)
    if p.status_cell:
        ws.Range(p.status_cell).Value2 = status
    return status


def read_block(ws, p: SheetPlan, b, key: str = "block"):
    """The read of A1:<last col><last row> (Value2) recorded into b (ws: the Answer sheet's node there); the function
    returned gives the block from the batch's result, always a grid."""
    b.ret(ws.Range(f"A1:{p.last_col}{p.last_row}").Value2, key)

    def block(res) -> tuple:
        v = res.values[key]
        return v if isinstance(v, tuple) else ((v,),)
    return block
