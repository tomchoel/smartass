"""Changes to the user's workbook: write cells, insert/delete rows and columns, add a sheet."""
from __future__ import annotations

import re

from openpyxl.utils import get_column_letter as L
from openpyxl.utils.cell import range_boundaries

from . import books, fmt, history
from .. import formulas as fx
from ..excel.jobs import com
from ..excel.target import _wait_ready
from ..host import InputError

def _check_written(ws, rng) -> dict:
    """tn-formelstil on the formulas in the changed rows (whole used width, so row consistency is judged too)."""
    ur = com(lambda: ws.UsedRange)
    r1, r2 = com(lambda: rng.Row), com(lambda: rng.Row) + int(com(lambda: rng.Rows.Count)) - 1
    c1, c2 = com(lambda: ur.Column), com(lambda: ur.Column) + int(com(lambda: ur.Columns.Count)) - 1
    if (r2 - r1 + 1) * (c2 - c1 + 1) > 60_000:
        c1, c2 = com(lambda: rng.Column), com(lambda: rng.Column) + int(com(lambda: rng.Columns.Count)) - 1
    block = com(lambda: ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)))
    name = com(lambda: ws.Name)
    cells = [(name, f"{L(c1 + j)}{r1 + i}", v) for i, row in enumerate(books._grid(com(lambda: block.Formula2)))
             for j, v in enumerate(row) if isinstance(v, str) and v.startswith("=")]
    return fx.check(cells) if cells else {"ok": True, "checked": 0, "block": [], "warn": []}


def _text_safe(v):
    """Values are typed like the user would: numbers stay numbers, '=...' is a formula, other text is stored AS TEXT
    (Excel would otherwise turn 'Sep 26' into a date and '1/2' into a fraction)."""
    if isinstance(v, str) and not v.startswith("=") and v != "":
        return "'" + v
    return v


# ------------------------------------------------------------------ write
def write(workbook, sheet, address, values=None, formulas=None, formula=None, number_format=None,
          allow_archive=False, label=None) -> dict:
    given = [x is not None for x in (values, formulas, formula)]
    if sum(given) > 1:
        raise InputError("Give ONE of values (2D), formulas (2D) or formula (one formula filled over the range).")
    if not any(given) and number_format is None:
        raise InputError("Nothing to write: give values, formulas, formula or number_format.")
    b = books.find_book(workbook, write=True, allow_archive=allow_archive)
    ws = b.sheet(sheet)
    _wait_ready(b.app)
    grid = values if values is not None else formulas
    if grid is not None:
        if not grid or not isinstance(grid, list) or not isinstance(grid[0], list):
            raise InputError("values/formulas must be a 2D list: [[row1...], [row2...]].")
        width = max(len(r) for r in grid)
        grid = [list(r) + [None] * (width - len(r)) for r in grid]
        anchor = books._rng(ws, address)
        rng = com(lambda: anchor.Cells(1, 1).GetResize(len(grid), width))   # late binding: Get<Prop>(args)
    else:
        rng = books._rng(ws, address)
    addr = books._addr(rng)
    ch = history._new_change(b, ws, "write", addr, history._capture(ws, rng), label=label)

    def mutate():
        if values is not None:
            com(lambda: setattr(rng, "Formula2", tuple(tuple(_text_safe(v) for v in r) for r in grid)))
        elif formulas is not None:
            com(lambda: setattr(rng, "Formula2", tuple(tuple(r) for r in grid)))
        elif formula is not None:
            com(lambda: setattr(rng, "Formula2", formula))     # relative refs adjust per cell, like a fill
        if number_format is not None:
            com(lambda: setattr(rng, "NumberFormat", number_format))
    history._apply(b, ws, ch, mutate, undo=lambda: history._restore(ws, addr, ch["before"]))
    ch["after"] = history._state_hash(rng)
    history._save_change(ch)
    out = {"change_id": ch["id"], "workbook": b.name, "sheet": ch["sheet"], "address": addr,
           "note": "Not saved. xl_revert(change_id) undoes this change (Ctrl+Z cannot).", **_notes(b)}
    try:                                   # the change is applied and recorded: reporting must not hide that
        com(lambda: ws.Calculate())
        out["formula_style"] = _check_written(ws, rng)
        out["result_preview"] = [[books._clean(v) for v in row] for row in books._grid(com(lambda: rng.Value2))][:20]
        history._jump(b, ws, addr)
    except Exception as e:   # noqa: BLE001
        out["warning"] = f"written, but the check/preview failed: {e}"
    return out


def _notes(b: books.Book) -> dict:
    return {"notes": b.notes} if b.notes else {}


def rows_cols(workbook, sheet, action: str, at: str, allow_archive=False, force=False, label=None) -> dict:
    """insert_rows / delete_rows ('5:7') or insert_cols / delete_cols ('C:D')."""
    b = books.find_book(workbook, write=True, allow_archive=allow_archive)
    ws = b.sheet(sheet)
    _wait_ready(b.app)
    is_row = action.endswith("rows")
    spec = at.replace("$", "").strip()
    if ":" not in spec:
        spec = f"{spec}:{spec}"
    target = com(lambda: (ws.Rows if is_row else ws.Columns)(spec))
    ur = com(lambda: ws.UsedRange)
    if is_row:
        r1 = com(lambda: target.Row); n = int(com(lambda: target.Rows.Count))
        c2 = com(lambda: ur.Column) + int(com(lambda: ur.Columns.Count)) - 1
        area = f"A{r1}:{L(max(c2, 1))}{r1 + n - 1}"
    else:
        c1 = com(lambda: target.Column); n = int(com(lambda: target.Columns.Count))
        r2 = com(lambda: ur.Row) + int(com(lambda: ur.Rows.Count)) - 1
        area = f"{L(c1)}1:{L(c1 + n - 1)}{max(r2, 1)}"
    first = com(lambda: target.Row) if is_row else com(lambda: target.Column)
    if action.startswith("delete"):
        deps = _dependents_outside(b, ws, area, is_row)
        if deps and not force:
            raise InputError(f"{len(deps)} formula cell(s) depend on {spec} (e.g. {deps[:5]}); deleting turns them "
                             "into #REF! and a revert cannot repair that. Pass force=true if that is intended.")
        before = history._capture(ws, com(lambda: ws.Range(area)))
        follow = history._fingerprint(ws, is_row, first + n)       # the line that moves up/left into the gap
    else:
        before = None
        follow = history._fingerprint(ws, is_row, first)           # the line that moves down/right
    ch = history._new_change(b, ws, action, spec, before, extra={"area": area, "count": n, "first": first,
                                                         "is_row": is_row, "follow": follow}, label=label)
    history._apply(b, ws, ch, lambda: com(lambda: target.Insert() if action.startswith("insert") else target.Delete()))
    ch["after"] = None
    history._save_change(ch)
    history._jump(b, ws, spec)
    return {"change_id": ch["id"], "workbook": b.name, "sheet": ch["sheet"], "action": action, "at": spec,
            "note": "Not saved. xl_revert(change_id) undoes this.", **_notes(b)}


def _inside(area: str, sheet_cell: tuple[int, int]) -> bool:
    c1, r1, c2, r2 = range_boundaries(area.replace("$", ""))
    c, r = sheet_cell
    return (c1 or 1) <= c <= (c2 or 16384) and (r1 or 1) <= r <= (r2 or 1048576)


def _dependents_outside(b: books.Book, ws, area: str, rows_mode: bool) -> list[str]:
    """Formula cells outside the area that would become #REF! if it is deleted. Excel only breaks a reference that
    lies ENTIRELY inside the deleted rows/columns; whole columns (for a row delete) and partly overlapping ranges
    (SUM(C3:C5) when row 3 goes) just shrink. Scans every sheet's formulas, own sheet included."""
    me = com(lambda: ws.Name)
    ac1, ar1, ac2, ar2 = range_boundaries(area.replace("$", ""))
    out = []
    for name, a, _fo, (c1, r1, c2, r2) in fmt.refs_into(b, me, own=True):
        if f"{name}!{a}" in out or (name == me and _inside(area, range_boundaries(a)[:2])):
            continue
        if rows_mode:
            broken = r1 is not None and r2 is not None and ar1 <= r1 and r2 <= ar2
        else:
            broken = c1 is not None and c2 is not None and ac1 <= c1 and c2 <= ac2
        if broken:
            out.append(f"{name}!{a}")
    return out


def add_sheet(workbook, name: str, after: str | None = None, allow_archive=False) -> dict:
    name = (name or "").strip()
    if not name or len(name) > 31 or re.search(r"[:\\/?*\[\]]", name) or name.startswith("'") or name.endswith("'"):
        raise InputError(f"'{name}' is not a valid sheet name (1-31 characters, none of : \\ / ? * [ ], no ' at the "
                         "ends).")
    b = books.find_book(workbook, write=True, allow_archive=allow_archive)
    _wait_ready(b.app)
    if any(str(s.Name).lower() == name.lower() for s in b.wb.Worksheets):
        raise InputError(f"A sheet named '{name}' already exists.")
    ref = b.sheet(after) if after else com(lambda: b.wb.Worksheets(b.wb.Worksheets.Count))
    ws = com(lambda: b.wb.Worksheets.Add(None, ref))   # positional: keyword After= is ignored via late-bound COM
    try:
        com(lambda: setattr(ws, "Name", name))
    except Exception:
        history.delete_sheet(b.app, ws)
        raise
    ch = history._new_change(b, ws, "add_sheet", "A1", None, extra={"name": com(lambda: ws.Name)})
    ch["state"] = "applied"
    history._save_change(ch)
    history._jump(b, ws, "A1")
    return {"change_id": ch["id"], "workbook": b.name, "sheet": ch["sheet"], "note": "Not saved.", **_notes(b)}
