"""Helpers for the live matrix: Excel helpers, the engine's drill Edge, table parsing and cleanup; the injections
are added by later tasks."""
from __future__ import annotations

import re
import shutil
from pathlib import Path

_IN_MATRIX = re.compile(r"[\\/]_matrix[\\/][^\\/]+$", re.IGNORECASE)


def _excel_apps() -> list:
    """Every running Excel instance reachable through the running object table (one per window handle)."""
    import pythoncom
    import win32com.client

    apps: dict[int, object] = {}
    rot = pythoncom.GetRunningObjectTable()
    for moniker in rot:
        try:
            obj = win32com.client.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
            app = obj.Application
            apps.setdefault(int(app.Hwnd), app)
        except Exception:
            continue
    return list(apps.values())


def close_scratch_workbooks() -> list[str]:
    """Close, unsaved, the open workbooks that live in a _matrix folder. Never touches any other workbook.
    FullName can be a OneDrive https URL, so the match is on the parent folder name, not the local path."""
    closed: list[str] = []
    try:
        apps = _excel_apps()
    except Exception:
        return closed
    for app in apps:
        try:
            books = [app.Workbooks(i) for i in range(1, app.Workbooks.Count + 1)]
        except Exception:
            continue
        for wb in books:
            try:
                name = str(wb.FullName)
                if _IN_MATRIX.search(name):
                    wb.Close(SaveChanges=False)
                    closed.append(name)
            except Exception:
                continue
    return closed


def close_scratch_book(path: str, save: bool = False) -> bool:
    """Close ONE workbook of ours (it must live in a _matrix folder); save first if asked and it has changes."""
    want = Path(path).name.lower()
    for app in _excel_apps():
        try:
            books = [app.Workbooks(i) for i in range(1, app.Workbooks.Count + 1)]
        except Exception:
            continue
        for wb in books:
            try:
                name = str(wb.FullName)
                if str(wb.Name).lower() != want or not _IN_MATRIX.search(name):
                    continue
                if save and not wb.Saved:
                    wb.Save()
                wb.Close(SaveChanges=False)
                return True
            except Exception:
                continue
    return False


def kill_engine_edge() -> list[int]:
    """Kill only the Edge processes the engine's drill started: their --user-data-dir is the drill profile.
    The user's own Edge (another profile) is never touched."""
    from smartview_mcp import config
    from smartview_mcp.drill_browser import kill_our_chrome
    return kill_our_chrome(config.DRILL_PROFILE)


def number(table: str, row: int = 0, col: int = -1) -> float | None:
    """A value from an sv_query markdown table ('1,858,190,823.29'): data row `row`, column `col`."""
    data = [ln for ln in table.splitlines()[2:] if ln.startswith("|")]
    if row >= len(data):
        return None
    cell = [c.strip() for c in data[row].strip("|").split("|")][col]
    try:
        return float(cell.replace(",", "").replace(" ", ""))
    except ValueError:
        return None


def values(client, wb: str, sheet: str, rng: str, what: str = "values") -> dict:
    """xl_read of one range of an open workbook (values, formulas, both or formats)."""
    body, _ = client.must("xl_read", {"workbook": wb, "sheet": sheet, "range": rng, "what": what})
    return body


def cleanup(ctx) -> None:
    """Close the scratch workbooks unsaved, delete the extra files the run made elsewhere (ctx.state['files'], e.g.
    sv_refresh_workbook's output) and the _matrix folder, unless --keep."""
    if ctx.keep:
        return
    close_scratch_workbooks()
    for f in ctx.state.get("files", []):
        try:
            Path(f).unlink(missing_ok=True)
        except OSError:
            pass
    shutil.rmtree(ctx.scratch, ignore_errors=True)
