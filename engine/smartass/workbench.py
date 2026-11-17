# engine/smartass/workbench.py
"""v2's safety net for edits to an open workbook: xl_snapshot records a range's state in v1's snapshot store
(workbench/history, kept 7 days) without changing anything; xl_revert puts it back. Use before any Excel MCP edit.
Runs inside a user-Excel job (server wraps it in _user)."""
from __future__ import annotations


def _find_book(workbook, allow_archive: bool = False):
    from .xl.workbench import books
    return books.find_book(workbook, write=True, allow_archive=allow_archive)            # also pauses AutoSave, as v1 does for edits


def _wait_ready(app):
    from .xl.target import _wait_ready as wait
    wait(app)


def _range(ws, address):
    from .xl.jobs import com
    from .xl.workbench import books
    rng = books._rng(ws, address) if address else com(lambda: ws.UsedRange)
    return rng, books._addr(rng)


def _sheet_name(ws) -> str:
    from .xl.jobs import com
    return str(com(lambda: ws.Name))


def _capture(ws, rng):
    from .xl.workbench import history
    return history._capture(ws, rng)


def _record(b, ws, address, snap, label):
    from .xl.workbench import history
    ch = history._new_change(b, ws, "write", address, snap, label=label)
    ch["state"] = "applied"                                 # nothing is mutated; no "after" hash, so revert never
    history._save_change(ch)                                # refuses on "edited after" - the edit follows this
    return ch


def _history_revert(change_id, workbook, force):
    from .xl.workbench import history
    return history.revert(change_id, workbook, force=force)


def snapshot(workbook: str | None, sheet: str | None, address: str | None, label: str = "xl_snapshot",
             allow_archive: bool = False) -> dict:
    b = _find_book(workbook, allow_archive)
    ws = b.sheet(sheet)
    _wait_ready(b.app)
    rng, addr = _range(ws, address)
    ch = _record(b, ws, addr, _capture(ws, rng), label)
    out = {"change_id": ch["id"], "workbook": b.name, "sheet": sheet or _sheet_name(ws), "address": addr}
    notes = list(getattr(b, "notes", None) or [])          # e.g. "AutoSave was switched off" (v1 find_book)
    if notes:
        out["notes"] = notes
    return out


def revert(change_id: str | None, workbook: str | None, force: bool = False) -> dict:
    return _history_revert(change_id, workbook, force)


def _format_sheet(workbook, sheet, last_col, rows, allow_archive: bool = False):
    from .xl.workbench import fmt
    return fmt.format_sheet(workbook, sheet, last_col, rows, allow_archive=allow_archive)


def _check(workbook, sheet):
    from .xl.workbench import fmt
    return fmt.check(workbook, sheet)


def tidy(workbook, sheet, rows: dict | None = None, last_col: str | None = None,
         allow_archive: bool = False) -> dict:
    """TN style first when rows + last_col are given (v1 snapshots before styling, so it can be reverted), then
    always the tn-formelstil check ({ok, block, warn, counts, ...} - ok None means the checker itself failed)."""
    out: dict = {}
    if rows and last_col:
        out["styled"] = _format_sheet(workbook, sheet, last_col, rows, allow_archive)
    out["check"] = _check(workbook, sheet)
    return out
