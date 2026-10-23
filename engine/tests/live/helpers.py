"""Helpers for the live matrix: Excel helpers, the engine's drill Edge, table parsing, cleanup, the dialog watchdog
and the S8 injections."""
from __future__ import annotations

import contextlib
import re
import shutil
import subprocess
import threading
import time
from pathlib import Path

from smartview_mcp.auditlib import is_engine_excel

_IN_MATRIX = re.compile(r"[\\/]_matrix[\\/][^\\/]+$", re.IGNORECASE)
BM_CLICK, WM_GETOBJECT, OBJID_NATIVEOM, SMTO_ABORTIFHUNG = 0x00F5, 0x003D, -16, 0x0002


# ================================================================== engine-owned Excel and the dialog watchdog
def engine_excel_pids(scratch) -> set[int]:
    """PIDs of the Excel processes the engine or the matrix owns (see auditlib.is_engine_excel)."""
    import psutil

    out = set()
    for p in psutil.process_iter(["name", "cmdline"]):
        try:
            if (p.info["name"] or "").lower() == "excel.exe" and \
                    is_engine_excel(subprocess.list2cmdline(p.info["cmdline"] or []), str(scratch)):
                out.add(p.pid)
        except (psutil.Error, TypeError):
            continue
    return out


def _windows(pids: set[int], visible: bool = True) -> list[tuple[int, int, str, str]]:
    """[(hwnd, pid, class, title)] of the top-level windows of `pids`."""
    import win32gui
    import win32process

    out = []

    def cb(h, _):
        try:
            if visible and not win32gui.IsWindowVisible(h):
                return True
            pid = win32process.GetWindowThreadProcessId(h)[1]
            if pid in pids:
                out.append((h, pid, win32gui.GetClassName(h), win32gui.GetWindowText(h)))
        except Exception:
            pass
        return True
    win32gui.EnumWindows(cb, None)
    return out


def _children(hwnd: int) -> list[tuple[int, str, str]]:
    import win32gui

    out = []

    def cb(h, _):
        try:
            out.append((h, win32gui.GetClassName(h), win32gui.GetWindowText(h)))
        except Exception:
            pass
        return True
    try:
        win32gui.EnumChildWindows(hwnd, cb, None)
    except Exception:
        pass
    return out


def dialog_text(hwnd: int) -> str:
    return " ".join(t.strip() for _h, cls, t in _children(hwnd) if cls == "Static" and t.strip())[:300]


def click_button(hwnd: int, names=("OK", "No", "Close")) -> str | None:
    """BM_CLICK the first button titled one of `names` (in that order of preference); returns its title."""
    import win32gui

    buttons = {t.replace("&", "").strip().lower(): h for h, cls, t in _children(hwnd) if cls == "Button"}
    for n in names:
        h = buttons.get(n.lower())
        if h:
            win32gui.PostMessage(h, BM_CLICK, 0, 0)
            return n
    return None


def _kill_engine_excel(pid: int, scratch) -> bool:
    """Kill `pid` only if it is (still) an engine-owned Excel."""
    import psutil

    if pid not in engine_excel_pids(scratch):
        return False
    try:
        psutil.Process(pid).kill()
        return True
    except psutil.Error:
        return False


class DialogWatch:
    """Context manager: every `interval` s a daemon thread looks at the visible top-level windows of ENGINE-OWNED
    Excel processes only (auditlib.is_engine_excel). A '#32770' dialog or a window titled 'Login' / 'Script Error' is
    recorded ({title, text, pid, seconds_into_scenario}) and dismissed (BM_CLICK OK / No / Close). A 'Login' window is
    never clicked (its buttons would cancel a real sign-in; the SSO watcher clicks 'Telenor SSO', which today took up
    to 18 s, with loop stalls of 51-76 s); one still there `login_grace` s after it was first seen gets its Excel
    killed. PIDs in `ignoring()` (a dialog a scenario injected on purpose) are left alone."""

    WATCH_TITLES = ("Login", "Script Error")

    def __init__(self, scratch, interval: float = 3.0, login_grace: float = 90.0):
        self.scratch, self.interval, self.login_grace = scratch, interval, login_grace
        self.dialogs: list[dict] = []
        self._seen: dict[int, tuple[float, dict]] = {}
        self._ignore: set[int] = set()
        self._stop = threading.Event()
        self._t0 = time.monotonic()
        self._thread: threading.Thread | None = None

    def __enter__(self):
        self._t0 = time.monotonic()
        self._thread = threading.Thread(target=self._loop, daemon=True, name="dialog-watch")
        self._thread.start()
        return self

    def __exit__(self, *exc):
        self._stop.set()
        if self._thread:
            self._thread.join(15)
        return False

    @contextlib.contextmanager
    def ignoring(self, pid: int):
        self._ignore.add(pid)
        try:
            yield
        finally:
            self._ignore.discard(pid)

    def _loop(self):
        while not self._stop.wait(self.interval):
            try:
                self.scan()
            except Exception:
                pass

    def scan(self) -> None:
        pids = engine_excel_pids(self.scratch) - self._ignore
        if not pids:
            return
        now = time.monotonic()
        for hwnd, pid, cls, title in _windows(pids):
            if cls != "#32770" and title not in self.WATCH_TITLES:
                continue
            if hwnd not in self._seen:
                rec = {"title": title, "text": dialog_text(hwnd), "pid": pid,
                       "seconds_into_scenario": round(now - self._t0, 1)}
                self.dialogs.append(rec)
                self._seen[hwnd] = (now, rec)
                if title != "Login":
                    rec["clicked"] = click_button(hwnd)
                continue
            first, rec = self._seen[hwnd]
            if title == "Login":
                if now - first >= self.login_grace and not rec.get("killed"):
                    rec["killed"] = _kill_engine_excel(pid, self.scratch)
            elif not rec.get("clicked"):
                rec["clicked"] = click_button(hwnd)


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
