"""Helpers for the live matrix: cleanup now; injections and Excel helpers are added by later tasks."""
from __future__ import annotations

import re
import shutil

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


def cleanup(ctx) -> None:
    """Close the scratch workbooks unsaved and delete the _matrix folder, unless --keep."""
    if ctx.keep:
        return
    close_scratch_workbooks()
    shutil.rmtree(ctx.scratch, ignore_errors=True)
