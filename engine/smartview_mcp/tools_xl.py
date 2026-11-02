"""xl_* tools: live work in the user's open workbooks. Registered on the shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

from typing import Literal

from mcp.server.mcpserver.exceptions import ToolError

from .host import HostError
from .server import LOCAL, WRITE, _call_user, _doc, _err, tool

# ================================================================== xl_*: live work in the user's open workbooks
# Any open workbook, no Smart View involved. Active workbook by default (or name one). Every change is snapshotted
# (xl_revert undoes it - Ctrl+Z cannot undo automation), the screen stays frozen while writing and then jumps to the
# change, and the workbook is NEVER saved: the user reviews and saves. Formulas must pass tn-formelstil.
from .workbench import books, changes, fmt, history  # noqa: E402

XL_RULES = ("Formulas follow skill tn-formelstil: SUM over ranges (never a + chain of 3+ cells), label-keyed XLOOKUP "
            "(no INDEX/MATCH(1,1*...), no guessed ranges like $A$5:$Z$400), variance % "
            "=IF(D=0,\"\",(C-D)/ABS(D)), no hard-coded numbers. A result with formula_style.ok=false has BLOCK "
            "findings: fix them (write again) before calling the work done.")


async def _xl(fn, what: str):
    try:
        return await _call_user(lambda h: fn(), what=what)
    except (HostError, ToolError) as e:
        raise _err(e)


@tool(annotations=LOCAL)
async def xl_workbooks() -> list[dict]:
    """All workbooks open in the user's Excel window(s): name, path, sheets, active sheet, selection, unsaved
    changes, read-only, archive flag. Start here when the user talks about 'my workbook' / 'this sheet'."""
    return await _xl(books.workbooks, "xl_workbooks")


@tool(annotations=LOCAL)
async def xl_overview(workbook: str | None = None, sheet: str | None = None, preview_rows: int = 12) -> dict:
    """Map of an open workbook (default: the active one): per sheet the used range, formula/constant counts,
    tables, a value preview; plus defined names. Give sheet for a longer preview of one sheet."""
    return await _xl(lambda: books.overview(workbook, sheet, preview_rows), "xl_overview")


@tool(annotations=LOCAL)
async def xl_read(range: str | None = None, sheet: str | None = None, workbook: str | None = None,
                  what: Literal["values", "formulas", "both", "formats"] = "values",
                  max_cells: int = books.MAX_READ_CELLS) -> dict:
    """Read a range of an open workbook live (unsaved edits included). range: 'B5:F30', a name, or omitted/
    'selection' = what the user has selected. what='formulas' returns {cell: formula} for formula cells only."""
    if what == "formats":
        max_cells = min(max_cells, 600)
    return await _xl(lambda: books.read(workbook, sheet, range, what, max_cells), "xl_read")


@tool(annotations=WRITE)
@_doc(XL_RULES)
async def xl_write(range: str, sheet: str | None = None, workbook: str | None = None,
                   values: list[list] | None = None, formulas: list[list] | None = None, formula: str | None = None,
                   number_format: str | None = None, allow_archive: bool = False, label: str | None = None) -> dict:
    """Write into an open workbook live. ONE of: values (2D; text stays text, '=..' is a formula), formulas (2D),
    or formula (ONE formula written over the whole range - relative references adjust per cell, like a fill
    right/down). range: the target, or its top-left cell for values/formulas. number_format: US code, e.g.
    '#,##0;(#,##0);"–"'. Snapshot first -> change_id for xl_revert. Never saves. Archive files refused unless
    allow_archive (only when the user explicitly asked)."""
    return await _xl(lambda: changes.write(workbook, sheet, range, values, formulas, formula, number_format,
                                          allow_archive, label), "xl_write")


@tool(annotations=WRITE)
async def xl_rows(action: Literal["insert_rows", "delete_rows", "insert_cols", "delete_cols"], at: str,
                  sheet: str | None = None, workbook: str | None = None, force: bool = False,
                  allow_archive: bool = False, label: str | None = None) -> dict:
    """Insert or delete whole rows ('12' or '12:14') or columns ('D' or 'D:E') live. A delete that would break
    formulas (#REF!) is refused unless force=true. Snapshot -> xl_revert. Never saves."""
    return await _xl(lambda: changes.rows_cols(workbook, sheet, action, at, allow_archive, force, label), "xl_rows")


@tool(annotations=WRITE)
async def xl_add_sheet(name: str, after: str | None = None, workbook: str | None = None,
                       allow_archive: bool = False) -> dict:
    """Add an empty sheet to an open workbook (after the given sheet, default last). xl_revert deletes it."""
    return await _xl(lambda: changes.add_sheet(workbook, name, after, allow_archive), "xl_add_sheet")


@tool(annotations=WRITE)
async def xl_format(range: str, style: str | None = None, sheet: str | None = None, workbook: str | None = None,
                    number_format: str | None = None, column_width: float | None = None,
                    allow_archive: bool = False, label: str | None = None) -> dict:
    """TN report style on a range: style = a TN cell style ('TN Tittel', 'TN Seksjon', 'TN Kolonne tall', 'TN Note'
    ...) or a number kind with optional emphasis: 'heltall', 'heltall.sum', 'prosent1.hovedsum', 'etikett.sum',
    'tall1', 'faktor', 'belop.input' ... Styles are installed in the workbook if missing. Snapshot -> xl_revert."""
    return await _xl(lambda: fmt.format_range(workbook, sheet, range, style, number_format, column_width,
                                                 allow_archive, label), "xl_format")


@tool(annotations=WRITE)
async def xl_format_sheet(rows: dict[str, str], last_col: str, sheet: str | None = None, workbook: str | None = None,
                          label_cols: list[str] | None = None, num: dict[str, str] | None = None,
                          col_widths: dict[str, float] | None = None, freeze_rows: int | None = None,
                          allow_archive: bool = False) -> dict:
    """Whole sheet in TN report style (skill tn-rapportstil; the same engine as the smartview answers), contents
    untouched. rows: {"1": "title", "2": "subtitle", "4": "header", "5-12": "line", "13": "sum", "15": "grand",
    "17": "note", "3": "section", "6": "text"}; unlisted rows with content = line, empty = blank. last_col: right
    edge of the table ('F'). label_cols default ['B']. num: number kind per value column {"C": "heltall",
    "F": "prosent1"} (default heltall). col_widths {"A": 2.5, "B": 46, "C:F": 17}. Returns style_lint. Snapshot."""
    return await _xl(lambda: fmt.format_sheet(workbook, sheet, last_col, rows, label_cols, num, col_widths,
                                                 freeze_rows, allow_archive), "xl_format_sheet")


@tool(annotations=LOCAL)
async def xl_trace(cell: str, direction: Literal["precedents", "dependents"] = "precedents",
                   sheet: str | None = None, workbook: str | None = None) -> dict:
    """Explain a cell: its formula and value, with row label / column header; precedents = every reference in the
    formula with its value (or sum), formula and labels - enough to explain the number in words; dependents = cells
    on this and other sheets that use it."""
    return await _xl(lambda: fmt.trace(workbook, sheet, cell, direction), "xl_trace")


@tool(annotations=LOCAL)
async def xl_check(sheet: str | None = None, workbook: str | None = None) -> dict:
    """tn-formelstil check of the live formulas in one sheet or the whole open workbook: BLOCK findings (+ chains,
    guessed fixed ranges, legacy INDEX/MATCH lookups, 0 on a zero base, hard-coded numbers) and warnings."""
    return await _xl(lambda: fmt.check(workbook, sheet), "xl_check")


@tool(annotations=LOCAL)
async def xl_changes(workbook: str | None = None, limit: int = 20) -> list[dict]:
    """Recent xl_* changes (newest first, kept 7 days): id, time, workbook, sheet, operation, range, reverted."""
    return history.changes(workbook, limit)


@tool(annotations=WRITE)
async def xl_revert(change_id: str | None = None, workbook: str | None = None, force: bool = False) -> dict:
    """Undo an xl_* change from its snapshot (values, formulas, formats, column widths; inserted rows/sheets are
    removed, deleted rows restored). Default: the newest unreverted change in the workbook. Refuses when the cells
    were edited afterwards or later changes on the sheet are not reverted yet, unless force=true."""
    return await _xl(lambda: history.revert(change_id, workbook, force), "xl_revert")
