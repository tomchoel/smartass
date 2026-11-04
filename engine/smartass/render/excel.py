# engine/smartass/render/excel.py
"""The COM writer: Q<n> Data as an Excel table (label columns as text first, so 'Sep 26' never becomes a date),
Q<n> Answer as one rectangle of values and formulas (Formula2), then v1's TN styling. Runs inside a user-Excel job
(v1 host().call(..., user_excel=True)); every COM call goes through v1 jobs.com (busy retry)."""
from __future__ import annotations

from ..xl import styling
from ..xl.jobs import com

from ..model.errors import GuardError
from ..model.frame import COLUMNS, Frame
from .layout import SheetPlan, col_letter

XL_SRC_RANGE = 1
XL_YES = 1


def data_grid(frame: Frame) -> list[tuple]:
    return [COLUMNS, *frame.rows()]


def _put(ws, g: list[tuple]):
    """Label columns as text first, then the values; returns the range. Only for a range that is not a table yet."""
    n = len(g)
    labels = ws.Range(ws.Cells(1, 1), ws.Cells(n, len(COLUMNS) - 1))
    com(lambda: setattr(labels, "NumberFormat", "@"))
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(n, len(COLUMNS)))
    com(lambda: setattr(rng, "Value2", tuple(g)))
    return rng


def _create_table(ws, g: list[tuple], q: int) -> None:
    rng = _put(ws, g)
    t = com(lambda: ws.ListObjects.Add(XL_SRC_RANGE, rng, None, XL_YES))
    com(lambda: setattr(t, "Name", f"Data_Q{q}"))


def write_data(ws, frame: Frame, q: int) -> None:
    _create_table(ws, data_grid(frame), q)


def _table(ws, q: int):
    try:
        return com(lambda: ws.ListObjects(f"Data_Q{q}"))
    except Exception:   # noqa: BLE001 - no such table (an older answer)
        return None


def read_data(ws, q: int) -> list[tuple] | None:
    """Header + body of table Data_Q<q> as a grid, or None when the sheet has no such table."""
    lo = _table(ws, q)
    if lo is None:
        return None
    v = com(lambda: lo.Range.Value2)
    return [tuple(r) for r in v] if isinstance(v, tuple) else None


def rewrite_grid(ws, g: list[tuple], q: int) -> None:
    """Put grid g into the EXISTING table Data_Q<q>: resize it, write the values, clear rows left below when it
    shrank. The table is never deleted, recreated or renamed, so formulas that point at it keep working.
    Header and body are written as two rectangles: Value2 on one rectangle covering a table's header AND body deletes
    the table in Excel (measured 04.11.2026, even with its own values) and every formula on it becomes #REF!.
    No table: refused (GuardError) before anything is touched - clearing the sheet and creating a new table would
    leave the answer's formulas, and its backup, pointing at a deleted table."""
    lo = _table(ws, q)
    if lo is None:
        raise GuardError(f"table Data_Q{q} is missing on sheet 'Q{q} Data', so the answer cannot be redrawn in place",
                         "ask the question again with fin_analyze(spec); the old sheets were left as they were")
    old_rows = com(lambda: lo.Range.Rows.Count)
    n, w = len(g), len(COLUMNS)
    labels = ws.Range(ws.Cells(1, 1), ws.Cells(n, w - 1))
    com(lambda: setattr(labels, "NumberFormat", "@"))              # text first: 'Sep 26' never becomes a date
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(max(n, 2), w))          # a table keeps at least one body row
    com(lambda: lo.Resize(rng))
    head = ws.Range(ws.Cells(1, 1), ws.Cells(1, w))
    com(lambda: setattr(head, "Value2", (tuple(g[0]),)))
    if n > 1:
        body = ws.Range(ws.Cells(2, 1), ws.Cells(n, w))
        com(lambda: setattr(body, "Value2", tuple(tuple(r) for r in g[1:])))
    if old_rows > n:
        left = ws.Range(ws.Cells(n + 1, 1), ws.Cells(old_rows, w))
        com(lambda: left.Clear())


def rewrite_data(ws, frame: Frame, q: int) -> None:
    rewrite_grid(ws, data_grid(frame), q)


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
