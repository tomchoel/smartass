# engine/smartass/render/excel.py
"""The COM writer: Q<n> Data as an Excel table (label columns as text first, so 'Sep 26' never becomes a date),
Q<n> Answer as one rectangle of values and formulas (Formula2), then v1's TN styling. Runs inside a user-Excel job
(v1 host().call(..., user_excel=True)); every COM call goes through v1 jobs.com (busy retry)."""
from __future__ import annotations

from smartview_mcp import styling
from smartview_mcp.excel.jobs import com

from ..model.frame import COLUMNS, Frame
from .layout import SheetPlan, col_letter

XL_SRC_RANGE = 1
XL_YES = 1


def data_grid(frame: Frame) -> list[tuple]:
    return [COLUMNS, *frame.rows()]


def write_data(ws, frame: Frame, q: int) -> None:
    g = data_grid(frame)
    n = len(g)
    labels = ws.Range(ws.Cells(1, 1), ws.Cells(n, len(COLUMNS) - 1))
    com(lambda: setattr(labels, "NumberFormat", "@"))
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(n, len(COLUMNS)))
    com(lambda: setattr(rng, "Value2", tuple(g)))
    t = com(lambda: ws.ListObjects.Add(XL_SRC_RANGE, rng, None, XL_YES))
    com(lambda: setattr(t, "Name", f"Data_Q{q}"))


def answer_grid(p: SheetPlan) -> tuple[int, int, list[list]]:
    rows = [r for r, _ in p.cells]
    cols = [c for _, c in p.cells]
    r0, c0 = min(rows), min(cols)
    g = [[""] * (max(cols) - c0 + 1) for _ in range(max(rows) - r0 + 1)]
    for (r, c), v in p.cells.items():
        g[r - r0][c - c0] = "" if v is None else v
    return r0, c0, g


def write_answer(ws, p: SheetPlan, warnings: list[str]) -> list[str]:
    """Write every cell in one Formula2 rectangle, then style. Returns the styling lint findings."""
    r0, c0, g = answer_grid(p)
    rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(g) - 1, c0 + len(g[0]) - 1))
    com(lambda: setattr(rng, "Formula2", tuple(tuple(x) for x in g)))
    widths = {"B": 46, **{col_letter(c): 16 for c in range(3, col_number(p.last_col) + 1)}}
    roles = {"last_col": p.last_col, "label_cols": ["B"], "last_row": p.last_row, "col_widths": widths,
             "rows": p.rows}
    return styling.apply_report(ws, roles)


def read_block(ws, p: SheetPlan) -> tuple:
    rng = ws.Range(f"A1:{p.last_col}{p.last_row}")
    v = com(lambda: rng.Value2)
    return v if isinstance(v, tuple) else ((v,),)


def col_number(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n
