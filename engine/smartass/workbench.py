# engine/smartass/workbench.py
"""v2's safety net for edits to an open workbook: xl_snapshot records a range's state in v1's snapshot store
(workbench/history, kept 7 days) without changing anything; xl_revert puts it back. Use before any Excel MCP edit.
Runs inside a user-Excel job (server wraps it in _user)."""
from __future__ import annotations


def _find_book(workbook):
    from smartview_mcp.workbench import books
    return books.find_book(workbook, write=True)            # also pauses AutoSave, as v1 does for edits


def _wait_ready(app):
    from smartview_mcp.excel.target import _wait_ready as wait
    wait(app)


def _range(ws, address):
    from smartview_mcp.excel.jobs import com
    from smartview_mcp.workbench import books
    rng = books._rng(ws, address) if address else com(lambda: ws.UsedRange)
    return rng, books._addr(rng)


def _sheet_name(ws) -> str:
    from smartview_mcp.excel.jobs import com
    return str(com(lambda: ws.Name))


def _capture(ws, rng):
    from smartview_mcp.workbench import history
    return history._capture(ws, rng)


def _record(b, ws, address, snap, label):
    from smartview_mcp.workbench import history
    ch = history._new_change(b, ws, "write", address, snap, label=label)
    ch["state"] = "applied"                                 # nothing is mutated; no "after" hash, so revert never
    history._save_change(ch)                                # refuses on "edited after" - the edit follows this
    return ch


def _history_revert(change_id, workbook, force):
    from smartview_mcp.workbench import history
    return history.revert(change_id, workbook, force=force)


def snapshot(workbook: str | None, sheet: str | None, address: str | None, label: str = "xl_snapshot") -> dict:
    b = _find_book(workbook)
    ws = b.sheet(sheet)
    _wait_ready(b.app)
    rng, addr = _range(ws, address)
    ch = _record(b, ws, addr, _capture(ws, rng), label)
    return {"change_id": ch["id"], "workbook": b.name, "sheet": sheet or _sheet_name(ws), "address": addr}


def revert(change_id: str | None, workbook: str | None, force: bool = False) -> dict:
    return _history_revert(change_id, workbook, force)
