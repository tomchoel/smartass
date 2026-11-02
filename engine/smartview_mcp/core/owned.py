"""The tool-owned range of an answer sheet (design §3.3, phase 5): what smartass drew - A1 to the last column and
row of the table, checks, notes and findings - is stored with the spec. A redraw clears only that block and only
charts named CHART_NAME, and makes room (inserts rows/columns) when a larger answer would overwrite user content."""
from __future__ import annotations

from ..excel import sheets
from ..excel.jobs import com as _com

CHART_NAME = "smartass chart"


def growth(old: dict | None, new: dict, used) -> dict:
    """Rows/columns to insert before writing `new` over `old` (both {"rows", "cols"}); used(r1, c1, r2, c2) -> bool
    says whether that band holds anything. No old range (first draw / pre-phase-5 answer): nothing to insert."""
    out = {"rows": None, "cols": None}
    if not old:
        return out
    if new["rows"] > old["rows"] and used(old["rows"] + 1, 1, new["rows"], max(old["cols"], new["cols"])):
        out["rows"] = (old["rows"] + 1, new["rows"])
    if new["cols"] > old["cols"] and used(1, old["cols"] + 1, max(old["rows"], new["rows"]), new["cols"]):
        out["cols"] = (old["cols"] + 1, new["cols"])
    return out


def make_room(ws, old: dict | None, new: dict, com=_com) -> dict:
    """Insert rows below / columns right of the old block when the larger new block would cover user content."""
    def used(r1, c1, r2, c2):
        rng = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
        return com(lambda: ws.Application.WorksheetFunction.CountA(rng)) > 0
    g = growth(old, new, used)
    if g["rows"]:
        a, b = g["rows"]
        com(lambda: ws.Rows(f"{a}:{b}").Insert())
    if g["cols"]:
        a, b = g["cols"]
        com(lambda: ws.Range(ws.Cells(1, a), ws.Cells(1, b)).EntireColumn.Insert())
    return g


def clear(ws, old: dict | None, com=_com, cell=None) -> None:
    """Clear what the tool drew last time (everything for an answer without a stored range)."""
    cell = cell or ws.Cells
    if not old:
        com(lambda: ws.Cells.Clear())
    else:
        com(lambda: ws.Range(cell(1, 1), cell(old["rows"], old["cols"])).Clear())
    drop_charts(ws, com=com, everything=not old)


def drop_charts(ws, com=_com, everything: bool = False) -> None:
    for co in list(com(lambda: ws.ChartObjects())):
        if everything or str(co.Name).startswith(CHART_NAME):
            com(co.Delete)


def backup(t, name: str, com=_com) -> str:
    """Hidden copy of the answer sheet before a follow-up redraw; returns its name."""
    ws = com(lambda: t.wb.Worksheets(name))
    com(lambda: ws.Copy(None, ws))                                  # positional After (see sheets.add_sheet)
    copy_ = com(lambda: t.wb.Worksheets(ws.Index + 1))
    bname = sheets.unique_sheet_name(t, ("~" + name)[:31])
    com(lambda: setattr(copy_, "Name", bname))
    com(lambda: setattr(copy_, "Visible", 0))                      # xlSheetHidden
    return bname


def restore(t, name: str, bname: str, com=_com) -> None:
    """The redraw failed: the backup becomes the answer sheet again."""
    sheets.drop_parked(t, {name: name})                          # deletes the broken sheet (alerts off)
    b = com(lambda: t.wb.Worksheets(bname))
    com(lambda: setattr(b, "Visible", -1))                          # xlSheetVisible
    com(lambda: setattr(b, "Name", name))
