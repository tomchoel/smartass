"""Helpers for the live matrix: Excel helpers, the engine's drill Edge, table parsing, cleanup, the dialog watchdog
and the S8 injections."""
from __future__ import annotations

import contextlib
import os
import re
import shutil
import subprocess
import threading
import time
from pathlib import Path

from auditlib import is_engine_excel, poll_kill, stale_suspended

_IN_MATRIX = re.compile(r"[\\/]_matrix[\\/][^\\/]+$", re.IGNORECASE)
BM_CLICK, WM_GETOBJECT, OBJID_NATIVEOM, SMTO_ABORTIFHUNG = 0x00F5, 0x003D, -16, 0x0002
# '#32770': Win32 message boxes (Smart View's). 'NUIDialog': Excel's own alerts (ALERT, MsgBox) - no Button children,
# so the watchdog can only record those.
DIALOG_CLASSES = ("#32770", "NUIDialog")


# ================================================================== engine-owned Excel and the dialog watchdog
_CLIENTS: list = []          # the matrix's own MCP clients (ctx.client, S8.parallel's second): see matrix_owners()


@contextlib.contextmanager
def tracked(client):
    """While inside, `client`'s child_server.py processes count as owners of engine Excels (matrix_owners)."""
    _CLIENTS.append(client)
    try:
        yield client
    finally:
        _CLIENTS.remove(client)


def matrix_owners() -> dict[int, float]:
    """{pid: create time} of the child_server.py processes of the matrix's own clients (read fresh: the relay may
    have restarted one)."""
    import psutil

    out = {}
    for c in list(_CLIENTS):
        try:
            pids = server_pids(c)
        except (psutil.Error, AttributeError):
            continue
        for pid in pids:
            try:
                out[pid] = psutil.Process(pid).create_time()
            except psutil.Error:
                continue
    return out


def pid_registry() -> list[dict]:
    """The engine's pids.json (written by hostproc._PidRegistry: [{pid, ctime, owner, owner_ctime}])."""
    import json

    from smartass.core import config

    try:
        items = json.loads(Path(config.PID_FILE).read_text())
    except (OSError, ValueError):
        return []
    return [i for i in items if isinstance(i, dict)] if isinstance(items, list) else []


def engine_excel_pids(scratch) -> set[int]:
    """PIDs of the Excel processes the matrix owns (see auditlib.is_engine_excel): the hidden Excels pids.json records
    for the matrix's own servers, and our '/x <scratch workbook>' Excels."""
    import psutil

    registry, owners = pid_registry(), matrix_owners()
    out = set()
    for p in psutil.process_iter(["name", "cmdline", "create_time"]):
        try:
            if (p.info["name"] or "").lower() == "excel.exe" and \
                    is_engine_excel(subprocess.list2cmdline(p.info["cmdline"] or []), str(scratch), p.pid,
                                    p.info["create_time"], registry, owners):
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
            if cls not in DIALOG_CLASSES and title not in self.WATCH_TITLES:
                continue
            if hwnd not in self._seen:
                rec = {"title": title, "text": dialog_text(hwnd), "pid": pid,
                       "seconds_into_scenario": round(now - self._t0, 1)}
                self.dialogs.append(rec)
                self._seen[hwnd] = (now, rec)
                if title != "Login":
                    rec["clicked"] = dismiss_dialog(hwnd) or click_button(hwnd)
                continue
            first, rec = self._seen[hwnd]
            if title == "Login":
                if now - first >= self.login_grace and not rec.get("killed"):
                    rec["killed"] = _kill_engine_excel(pid, self.scratch)
            elif not rec.get("clicked"):
                rec["clicked"] = dismiss_dialog(hwnd) or click_button(hwnd)


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


def excel_pids() -> set[int]:
    import psutil
    return {p.pid for p in psutil.process_iter(["name"]) if (p.info["name"] or "").lower() == "excel.exe"}


def leftover_excel(pid: int, before: set[int], n_books: int, parent: str | None) -> bool:
    """An Excel the run started (the engine opens one per parallel agent), left without a workbook once the scratch
    books are closed. Never one that was running before the run or that the user started (parent explorer.exe)."""
    return pid not in before and n_books == 0 and (parent or "").lower() != "explorer.exe"


def quit_leftover_excels(before: set[int]) -> list[int]:
    import psutil
    import win32process

    quit_: list[int] = []
    for app in _excel_apps():
        try:
            pid = win32process.GetWindowThreadProcessId(int(app.Hwnd))[1]
            try:
                parent = psutil.Process(pid).parent()
                pname = parent.name() if parent else None
            except psutil.Error:
                pname = None
            books = int(app.Workbooks.Count) + int(app.ProtectedViewWindows.Count)   # a Protected View file counts
            if leftover_excel(pid, before, books, pname):
                app.Quit()
                quit_.append(pid)
        except Exception:
            continue
    return quit_


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
    from smartass.core import config
    from smartview_mcp.drill_browser import kill_our_chrome
    return kill_our_chrome(config.DRILL_PROFILE)


def kill_engine_edge_when_seen(done, timeout: float = 60, interval: float = 0.5) -> tuple[list[int], bool]:
    """Poll for the drill Edge for up to `timeout` s and kill it the moment it appears; stops early once `done()`
    (the call under test finished). Returns (killed pids, the call was still running at the kill)."""
    return poll_kill(kill_engine_edge, done, timeout, interval, time.sleep, time.monotonic)


def oracle_cell(lay, values: list[list]) -> float | None:
    """The single data cell of a one-row, one-column provider grid (row 1 = POV, then the column headers, then the
    data rows; the data starts after the row-label columns). None for #Missing/#NoAccess or a blank."""
    v = values[1 + lay.nc][lay.nr]
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def oracle_value(connection: str, pov: dict, row: tuple[str, str], col: tuple[str, str]) -> float | None:
    """One cube cell over the Smart View HTTP provider, in THIS process (bench.oracle_pass runs it after the MCP
    server exited, so its hidden Edge is free to start on the drill profile)."""
    from smartview_mcp import svhttp
    from smartview_mcp.grid import Axis
    f = svhttp.shared().fetch(connection, pov, [Axis(row[0], [row[1]])], [Axis(col[0], [col[1]])])
    return oracle_cell(f.layout, f.values)


WHS_POV = {"All Entities": "un.SY.2WHS", "All Scenarios": "Actual", "All Currencies": "NOK",
           "All Balance Types": "Accumulated"}


def tiny_spec(question: str = "LIVE CHECK: EBITDA un.SY.2WHS YTD Sep 26", pov: dict | None = None,
              accounts: list[str] | None = None, months: list[str] | None = None) -> dict:
    """The smallest answer spec v1 (one pull, one line per account x month); `pov` defaults to Wholesale YTD."""
    pov = pov or WHS_POV
    accounts, months = accounts or ["ac.EBITDA"], months or ["Sep 26"]
    return {"version": 1, "question": question,
            "gate": {"outcome": "clear", "checklist": {
                "period": ", ".join(months), "scenario": "Actual", "entity": pov.get("All Entities", ""),
                "measure": ", ".join(accounts), "currency_sign": "NOK, EPM signs", "granularity": "one cell each"}},
            "pulls": [{"id": "p", "connection": "FinMax",
                       "rows": [{"dimension": "All Accounts", "members": accounts}],
                       "columns": [{"dimension": "All Time Periods", "members": months}], "pov": pov}],
            "layout": {"title": question, "columns": [{"header": m} for m in months],
                       "sections": [{"title": "Lines", "lines": [
                           {"key": f"a{i}", "label": a, "cells": [{"pull": "p", "row": [a], "col": [m]} for m in months]}
                           for i, a in enumerate(accounts)]}]}}


TINY_SPEC = tiny_spec()


def tiny_answer(client, target: str = "new", timeout: float = 600, spec: dict | None = None):
    """The smallest pipeline call (one pull, one line) - what the matrix used the old query tool for: (body, secs, err)."""
    return client.call("sv_answer", {"spec": spec or TINY_SPEC, "target": target}, timeout=timeout)


def values(client, wb: str, sheet: str, rng: str, what: str = "values") -> dict:
    """xl_read of one range of an open workbook (values, formulas, both or formats)."""
    body, _ = client.must("xl_read", {"workbook": wb, "sheet": sheet, "range": rng, "what": what})
    return body


def cleanup(ctx) -> None:
    """Stop the fresh install's processes (S9) always; then, unless --keep: close the scratch workbooks unsaved, the
    _matrix folder and the fresh profile."""
    import pythoncom

    # COM may so far only have been initialised on worker threads (bench's FirstVisible); the COM calls below run here
    pythoncom.CoInitialize()
    fresh = ctx.state.get("fresh") or {}
    if fresh.get("owners"):
        kill_fresh_excels(fresh["owners"])
    if FRESH_ROOT.exists():
        stop_fresh_processes()
    if ctx.keep:
        return
    remove_fresh()
    close_scratch_workbooks()
    if ctx.state.get("excels_before") is not None:
        left = quit_leftover_excels(ctx.state["excels_before"])
        if left:
            print(f"cleanup: quit the empty Excel(s) this run started: {left}")
    shutil.rmtree(ctx.scratch, ignore_errors=True)


# ================================================================== S8 injections (each undoes itself in finally)
BAK = ".matrix-bak"
COOKIE_FILES = ("Cookies", "Cookies-journal")


def _cookie_dirs() -> list[Path]:
    """Cookie folders of the engine's drill Edge profiles (drill_edge, drill_edge_<channel>)."""
    from smartass.core import config

    out = []
    for prof in Path(config.DRILL_PROFILE).parent.glob(Path(config.DRILL_PROFILE).name + "*"):
        if prof.is_dir():
            out += [d for d in (prof / "Default" / "Network", prof / "Default") if d.is_dir()]
    return out


def restore_matrix_bak() -> list[str]:
    """Put every *.matrix-bak cookie store back (a crashed S8.sso run left it). A store the engine's Edge made in the
    meantime is replaced, and a fresh journal without a backup is dropped so it cannot roll into the old database."""
    restored = []
    dirs = [d for d in _cookie_dirs() if any(d.glob("*" + BAK))]
    if not dirs:
        return restored
    kill_engine_edge()                        # it holds the cookie store open
    time.sleep(1)
    for d in dirs:
        for name in COOKIE_FILES:
            if not (d / (name + BAK)).exists() and (d / name).exists() and name.endswith("-journal"):
                (d / name).unlink()
        for bak in d.glob("*" + BAK):
            orig = bak.with_name(bak.name[:-len(BAK)])
            for i in range(5):
                try:
                    bak.replace(orig)
                    restored.append(str(orig))
                    break
                except OSError:
                    if i == 4:
                        raise
                    time.sleep(1)
    return restored


@contextlib.contextmanager
def expired_sso():
    """The engine's drill Edge loses its sign-in: its cookie store is renamed to *.matrix-bak for the duration.
    (The engine keeps no SSO token file of its own; Smart View's EPM session lives inside its Excel.)
    A leftover *.matrix-bak (an earlier run's backup of the REAL sign-in) is restored first; if one still remains,
    Limit, and nothing is renamed, so a backup is never overwritten."""
    restore_matrix_bak()
    left = [str(b) for d in _cookie_dirs() for b in d.glob("*" + BAK)]
    if left:
        raise Limit(f"leftover cookie backup(s) could not be restored: {left}")
    kill_engine_edge()
    time.sleep(1)
    moved = []
    try:
        for d in _cookie_dirs():
            for name in COOKIE_FILES:
                f, bak = d / name, d / (name + BAK)
                if f.exists():
                    if bak.exists():
                        raise Limit(f"{bak} appeared: not overwritten")
                    f.rename(bak)                  # rename never replaces an existing file on Windows
                    moved.append(str(f))
        yield moved
    finally:
        restore_matrix_bak()


def engine_hidden_excel_pid(pid: int) -> int:
    """`pid` if it is the hidden Excel of one of the matrix's own servers (pids.json entry with a matching create
    time, owned by a matrix child_server.py, started '-Embedding'); else AssertionError."""
    import psutil

    if pid in engine_excel_pids(_scratch_root()):
        try:
            if "-embedding" in " ".join(psutil.Process(pid).cmdline()).lower():
                return pid
        except psutil.Error:
            pass
    raise AssertionError(f"Excel pid {pid} is not the hidden Excel of a matrix server (pids.json "
                         f"{[(i.get('pid'), i.get('owner')) for i in pid_registry()]}, owners {sorted(matrix_owners())})")


def server_pids(client) -> list[int]:
    """The real MCP server behind the relay: child_server.py processes under the client's run_server.py."""
    import psutil

    out = []
    for c in psutil.Process(client.p.pid).children(recursive=True):
        try:
            if "child_server.py" in " ".join(c.cmdline()):
                out.append(c.pid)
        except psutil.Error:
            continue
    return out


def _suspend_marker() -> Path:
    from smartass.core import config

    return Path(config.HOME) / "matrix-suspended.json"


def resume_leftover_suspended() -> list[int]:
    """A matrix run killed during S8.sleep left its servers suspended (holding drill_edge.lock etc.): resume every
    PID in the marker that is still the same process (auditlib.stale_suspended), then delete the marker."""
    import json

    import psutil

    m = _suspend_marker()
    if not m.exists():
        return []
    try:
        marker = json.loads(m.read_text())
    except (OSError, ValueError):
        marker = []
    running = {}
    for i in marker if isinstance(marker, list) else []:
        try:
            running[i["pid"]] = psutil.Process(i["pid"]).create_time()
        except (psutil.Error, KeyError, TypeError, ValueError):
            continue
    resumed = []
    for pid in stale_suspended(marker if isinstance(marker, list) else [], running):
        try:
            psutil.Process(pid).resume()
            resumed.append(pid)
        except psutil.Error:
            continue
    m.unlink(missing_ok=True)
    return resumed


@contextlib.contextmanager
def suspended(pids: list[int]):
    """Suspend `pids`; a marker file ({pid, ctime}) lets the next matrix start resume them if this run dies."""
    import json

    import psutil

    procs = [psutil.Process(pid) for pid in pids]
    m = _suspend_marker()
    m.parent.mkdir(parents=True, exist_ok=True)
    m.write_text(json.dumps([{"pid": p.pid, "ctime": p.create_time()} for p in procs]))
    done = []
    try:
        for p in procs:
            p.suspend()
            done.append(p)
        yield done
    finally:
        for p in done:
            try:
                p.resume()
            except psutil.Error:
                pass
        m.unlink(missing_ok=True)


def app_of(pid: int, timeout_ms: int = 3000):
    """(Application, main hwnd) of Excel `pid` through its workbook window (the caller has initialised COM)."""
    import pythoncom
    import win32com.client
    import win32gui

    for hwnd, _pid, cls, _t in _windows({pid}, visible=False):
        if cls != "XLMAIN":
            continue
        desk = win32gui.FindWindowEx(hwnd, 0, "XLDESK", None)
        ex7 = win32gui.FindWindowEx(desk, 0, "EXCEL7", None) if desk else 0
        if not ex7:
            continue
        _, lres = win32gui.SendMessageTimeout(ex7, WM_GETOBJECT, 0, OBJID_NATIVEOM, SMTO_ABORTIFHUNG, timeout_ms)
        win = win32com.client.Dispatch(pythoncom.ObjectFromLresult(lres, pythoncom.IID_IDispatch, 0))
        return win.Application, hwnd
    return None, None


def _book_names(app) -> list[str]:
    return [str(app.Workbooks(i).FullName) for i in range(1, app.Workbooks.Count + 1)]


def _login_open(pid: int) -> bool:
    return any(t == "Login" for _h, _p, _c, t in _windows({pid}))


@contextlib.contextmanager
def scratch_excel(path, watch=None, timeout: float = 240):
    """Our own Excel ('/x <scratch workbook>': engine-owned by auditlib.is_engine_excel) with `path` open; yields its
    pid. The workbook is first closed in any other Excel. Afterwards our Excel is closed, or killed if it hangs (an
    Excel that holds a workbook from outside _matrix is left alone)."""
    import pythoncom
    from smartass.core import config

    p = str(path)
    if not _IN_MATRIX.search(p):
        raise ValueError(f"not a scratch workbook: {p}")
    for _ in range(3):
        if not close_scratch_book(p, save=True):
            break
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 1
    proc = subprocess.Popen([config.EXCEL_EXE, "/x", p], startupinfo=si)
    try:
        with watch.ignoring(proc.pid) if watch else contextlib.nullcontext():   # its start-up login: SSO watcher
            pythoncom.CoInitialize()
            t0, want = time.monotonic(), Path(p).name.lower()
            while True:
                if proc.poll() is not None:
                    raise AssertionError(f"our scratch Excel (pid {proc.pid}) exited before it opened {want}")
                if time.monotonic() - t0 > timeout:
                    raise AssertionError(f"our scratch Excel (pid {proc.pid}) did not open {want} in {timeout:.0f}s")
                time.sleep(1)
                if _login_open(proc.pid):
                    continue
                try:
                    app, _ = app_of(proc.pid)
                    if app is not None and app.Ready and \
                            any(Path(n.replace("/", "\\")).name.lower() == want for n in _book_names(app)):
                        break
                except Exception:
                    continue
            time.sleep(3)
            while _login_open(proc.pid) and time.monotonic() - t0 < timeout:
                time.sleep(1)
        yield proc.pid
    finally:
        _quit_scratch_excel(proc)


def _quit_scratch_excel(proc) -> None:
    import pythoncom

    foreign: list[str] = []
    if proc.poll() is None:
        try:                                       # nothing modal may block the close (or outlive the kill)
            close_own_dialogs(_dialogs_of(proc.pid), proc.pid)
        except Exception:
            pass

    def quit_():
        pythoncom.CoInitialize()
        try:
            app, _ = app_of(proc.pid, 2000)
            if app is None:
                return
            app.DisplayAlerts = False
            for i in range(app.Workbooks.Count, 0, -1):
                wb = app.Workbooks(i)
                if _IN_MATRIX.search(str(wb.FullName)):
                    wb.Close(SaveChanges=False)
                else:
                    foreign.append(str(wb.Name))
            if not foreign:
                app.Quit()
        except Exception:
            pass
        finally:
            pythoncom.CoUninitialize()
    t = threading.Thread(target=quit_, daemon=True)
    t.start()
    t.join(30)
    if foreign:
        return
    try:
        proc.wait(20)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.wait(10)


def _dialogs_of(pid: int) -> list[int]:
    return [h for h, _p, cls, _t in _windows({pid}) if cls in DIALOG_CLASSES]


def _post_key(hwnd: int, vk: int) -> None:
    """A key press posted straight to one window of OUR Excel (no focus needed, so it can never reach another app)."""
    import win32api
    import win32con
    import win32gui

    scan = win32api.MapVirtualKey(vk, 0)
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 1 | scan << 16)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, 1 | scan << 16 | 0xC0000000)


class Limit(Exception):
    """The injection cannot run safely here; matrix.py scores it LIMIT (no retry)."""


def dismiss_dialog(hwnd: int) -> str | None:
    """Close one dialog without focus or keystrokes into the foreground: BM_CLICK OK/No/Close on a '#32770', or
    WM_CLOSE on an Excel 'NUIDialog' (its buttons are not windows; for an ALERT, close = OK). Returns what was done."""
    import win32con
    import win32gui

    try:
        cls = win32gui.GetClassName(hwnd)
    except Exception:
        return None
    if cls == "#32770":
        return click_button(hwnd)
    if cls == "NUIDialog":
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        return "WM_CLOSE"
    return None


def _gone(hwnds, pid: int, wait: float) -> bool:
    t0 = time.monotonic()
    while time.monotonic() - t0 < wait:
        if not set(hwnds) & set(_dialogs_of(pid)):
            return True
        time.sleep(0.5)
    return not set(hwnds) & set(_dialogs_of(pid))


def close_own_dialogs(hwnds, pid: int) -> str:
    """Make sure the dialogs `hwnds` of OUR Excel `pid` are gone: dismiss (BM_CLICK OK / WM_CLOSE), then Enter
    posted to the dialog, then kill that (engine-owned) Excel. Returns how they went; raises if they survive."""
    import win32con

    hwnds = [h for h in hwnds if h in _dialogs_of(pid)]
    if not hwnds:
        return "already gone"
    for h in hwnds:
        dismiss_dialog(h)
    if _gone(hwnds, pid, 5):
        return "dismissed"
    for h in hwnds:
        for c in [h] + [ch for ch, _cls, _t in _children(h)]:
            _post_key(c, win32con.VK_RETURN)
    if _gone(hwnds, pid, 5):
        return "dismissed with Enter"
    if _kill_engine_excel(pid, _scratch_root()) and _gone(hwnds, pid, 10):
        return "Excel killed"
    raise AssertionError(f"dialog(s) {hwnds} of Excel {pid} are still open")


@contextlib.contextmanager
def modal_dialog(pid: int, watch=None):
    """An Excel 4 ALERT("matrix") in OUR engine-owned scratch Excel `pid`, from a helper thread. Afterwards exactly
    that alert (the dialog window that appeared after it started) is closed and checked gone (close_own_dialogs).
    Refuses (Limit) a non-engine Excel, and refuses an Excel that already shows a dialog (no second alert)."""
    import pythoncom

    if pid not in engine_excel_pids(_scratch_root()):
        raise Limit(f"Excel {pid} is not engine-owned (user's Excel?): no alert injected")
    if _dialogs_of(pid):
        raise AssertionError(f"Excel {pid} already shows a dialog {_dialogs_of(pid)}: no second alert")
    state: dict = {}

    def run():
        pythoncom.CoInitialize()
        try:
            app, _ = app_of(pid)
            state["started"] = True
            state["result"] = app.ExecuteExcel4Macro('ALERT("matrix")')
        except Exception as e:
            state["error"] = repr(e)[:200]
        finally:
            pythoncom.CoUninitialize()
    with watch.ignoring(pid) if watch else contextlib.nullcontext():
        t = threading.Thread(target=run, daemon=True, name="matrix-alert")
        t.start()
        ours: list[int] = []
        try:
            t0 = time.monotonic()
            while not ours and time.monotonic() - t0 < 20:
                time.sleep(0.5)
                ours = _dialogs_of(pid)
            if not ours:
                raise AssertionError(f"the ALERT dialog did not appear in Excel {pid}: {state}")
            state["dialog"] = ours
            yield state
        finally:
            if not ours:                           # it may still come up after the 20 s wait
                time.sleep(2)
                ours = _dialogs_of(pid)
            state["dismissed"] = close_own_dialogs(ours, pid)
            t.join(15)


@contextlib.contextmanager
def edit_mode(pid: int, sheet: str | None = None):
    """Cell edit mode in OUR scratch Excel `pid`: activate the sheet, select an empty cell, F2 + 'x'; afterwards ESC
    (nothing is entered). The keys are posted to that Excel's grid window (EXCEL7), not typed into the foreground."""
    import pythoncom
    import win32con
    import win32gui

    if pid not in engine_excel_pids(_scratch_root()):
        raise Limit(f"Excel {pid} is not engine-owned (user's Excel?): no edit mode injected")
    pythoncom.CoInitialize()
    app, hwnd = app_of(pid)
    wb = next(app.Workbooks(i) for i in range(1, app.Workbooks.Count + 1)
              if _IN_MATRIX.search(str(app.Workbooks(i).FullName)))
    wb.Activate()
    ws = wb.Worksheets(sheet) if sheet else wb.ActiveSheet
    ws.Activate()
    ws.Range("AZ500").Select()
    app = wb = ws = None
    grid = win32gui.FindWindowEx(win32gui.FindWindowEx(hwnd, 0, "XLDESK", None), 0, "EXCEL7", None)
    _post_key(grid, win32con.VK_F2)
    time.sleep(0.5)
    win32gui.PostMessage(grid, win32con.WM_CHAR, ord("x"), 1)
    time.sleep(1.5)
    state = {"in_edit": None}
    try:
        try:
            a, _ = app_of(pid, 2000)
            state["in_edit"] = "no COM object (window busy)" if a is None else not a.Ready
        except Exception as e:
            state["in_edit"] = f"COM refused: {type(e).__name__}"
        if state["in_edit"] is False:
            raise AssertionError(f"Excel {pid} did not enter edit mode")
        yield state
    finally:
        _post_key(grid, win32con.VK_ESCAPE)
        time.sleep(1)
        try:
            a, _ = app_of(pid, 2000)
            state["escaped"] = bool(a.Ready)
        except Exception as e:
            state["escaped"] = f"COM refused: {type(e).__name__}"


def _scratch_root() -> Path:
    from smartass.core import config

    return Path(config.ANALYSER_DIR) / "_matrix"


# ================================================================== part 3: S9 fresh profile and S10 feature helpers
REPO = Path(__file__).resolve().parents[3]
FRESH_ROOT = REPO / ".matrix-fresh"        # not %TEMP%: a repo folder (git-ignored), deleted after the run
FRESH_PORT = "47899"
_DROP_ENV = ("SV_", "SMARTASS_", "SVKIT_")


def _uv_dir(var: str, args: list[str]) -> str:
    """The REAL uv cache / Python folder, read before LOCALAPPDATA/APPDATA are overridden (reuse, no re-download)."""
    v = os.environ.get(var)
    if v:
        return v
    return subprocess.run(["uv", *args], capture_output=True, text=True, check=True).stdout.strip()


def fresh_paths(root: Path = FRESH_ROOT) -> dict[str, Path]:
    home = root / "local" / "smartass"
    engine = home / "app" / "engine"
    return {"root": root, "home": home, "app": home / "app", "engine": engine,
            "py": engine / ".venv" / "Scripts" / "python.exe", "config": home / "config.toml",
            "pids": home / "pids.json", "pathfile": root / "user_path.txt", "release": root / "release",
            "copilot": root / "profile" / ".copilot" / "mcp-config.json", "claude": root / "profile" / ".claude.json",
            "startup": root / "roaming" / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"}


def fresh_profile(root: Path = FRESH_ROOT, version: str = "0.2.0") -> dict:
    """A new user's profile under `root`: a `git archive` of HEAD as release `version`, a copy of the real atlas.db,
    empty agent configs. Returns the env dict for the child installer/CLI/MCP processes (the parent's own
    environment is never changed). The user PATH goes to a file (SMARTASS_USER_PATH_FILE), the HTTP API to its own
    port, so nothing reaches the owner's registry, profile or running `serve`."""
    import json
    import os
    import sqlite3
    import zipfile

    from smartass.core import config

    remove_fresh(root)
    p = fresh_paths(root)
    for d in ("local", "roaming", "profile/.copilot", "onedrive", "release/atlas"):
        (root / d).mkdir(parents=True, exist_ok=True)
    p["copilot"].write_text("{}", encoding="utf-8")
    p["claude"].write_text("{}", encoding="utf-8")
    (p["release"] / "latest.json").write_text(json.dumps({"version": version}), encoding="utf-8")
    z = root / "rel.zip"
    subprocess.run(["git", "-C", str(REPO), "archive", "--format", "zip", "-o", str(z), "HEAD"], check=True,
                   capture_output=True)
    with zipfile.ZipFile(z) as zf:
        zf.extractall(p["release"] / version)
    z.unlink()
    su = (p["release"] / version / "engine" / "smartview_mcp" / "setup_user.py").read_text(encoding="utf-8")
    if "SMARTASS_USER_PATH_FILE" not in su:
        raise AssertionError("HEAD has no SMARTASS_USER_PATH_FILE override: its setup would write the real user PATH")
    src = sqlite3.connect(f"file:{config.ATLAS_DB}?mode=ro", uri=True)
    dst = sqlite3.connect(p["release"] / "atlas" / "atlas.db")
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    uv_cache = _uv_dir("UV_CACHE_DIR", ["cache", "dir"])
    uv_python = _uv_dir("UV_PYTHON_INSTALL_DIR", ["python", "dir"])
    env = {k: v for k, v in os.environ.items()
           if not k.upper().startswith(_DROP_ENV) and k.upper() not in ("VIRTUAL_ENV", "UV_PROJECT_ENVIRONMENT")}
    env.update({"UV_CACHE_DIR": uv_cache, "UV_PYTHON_INSTALL_DIR": uv_python,
                "LOCALAPPDATA": str(root / "local"), "APPDATA": str(root / "roaming"),
                "USERPROFILE": str(root / "profile"), "ONEDRIVECOMMERCIAL": str(root / "onedrive"),
                "ONEDRIVE": str(root / "onedrive"), "SMARTASS_RELEASE": str(p["release"]),
                "SMARTASS_HTTP_PORT": FRESH_PORT, "SMARTASS_USER_PATH_FILE": str(p["pathfile"]),
                "PYTHONIOENCODING": "utf-8"})
    return env


def _run(cmd: list[str], env: dict, cwd=None, timeout: float = 900) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, env=env, cwd=cwd, capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=timeout)


def fresh_install(env: dict, version: str = "0.2.0", root: Path = FRESH_ROOT) -> subprocess.CompletedProcess:
    """The release's installer in a child `powershell -NoProfile -File` with the fresh env."""
    ps1 = fresh_paths(root)["release"] / version / "installer" / "install.ps1"
    return _run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps1), "-NoDoctor"], env)


def fresh_cli(env: dict, *args: str, root: Path = FRESH_ROOT, timeout: float = 900) -> subprocess.CompletedProcess:
    """`smartass <args>` of the fresh install (its venv python, the fresh env)."""
    p = fresh_paths(root)
    return _run([str(p["py"]), "-m", "smartview_mcp.cli", *args], env, cwd=p["engine"], timeout=timeout)


def fresh_py(env: dict, code: str, root: Path = FRESH_ROOT) -> subprocess.CompletedProcess:
    p = fresh_paths(root)
    return _run([str(p["py"]), "-c", code], env, cwd=p["engine"], timeout=120)


def make_release(version: str, base: str = "0.2.0", root: Path = FRESH_ROOT) -> Path:
    """release\\<version>: a copy of release\\<base> with the version bumped (pyproject + the smartass uv.lock entry),
    then latest.json points to it."""
    import json

    rel = fresh_paths(root)["release"]
    dst = rel / version
    shutil.rmtree(dst, ignore_errors=True)
    shutil.copytree(rel / base, dst)
    for f, pat in ((dst / "engine" / "pyproject.toml", r'(?m)^version = "[^"]+"'),
                   (dst / "engine" / "uv.lock", r'(name = "smartass"\r?\nversion = )"[^"]+"')):
        text = f.read_text(encoding="utf-8")
        new, n = re.subn(pat, (f'version = "{version}"' if f.suffix == ".toml" else rf'\g<1>"{version}"'), text,
                         count=1)
        if n != 1:
            raise AssertionError(f"no version line in {f}")
        f.write_text(new, encoding="utf-8")
    (rel / "latest.json").write_text(json.dumps({"version": version}), encoding="utf-8")
    return dst


def _under(path: str | None, root: Path) -> bool:
    return bool(path) and str(path).lower().startswith(str(root).lower().rstrip("\\") + "\\")


def _fresh_env(p, root: Path) -> bool:
    """The process runs with the fresh profile's environment. COM starts Excel with the CALLER's environment, so the
    fresh server's hidden Excels have it too, even those the engine never registered (DispatchEx failed)."""
    import psutil

    try:
        return _under(p.environ().get("LOCALAPPDATA"), root)
    except (psutil.Error, OSError):
        return False


def stop_fresh_processes(root: Path = FRESH_ROOT) -> list[int]:
    """Kill the processes of the fresh install only: command line, exe, working folder or LOCALAPPDATA under `root`
    (its `serve`, the venv launcher's base-python child, a leftover installer, Excels started by its server), and
    their children."""
    import psutil

    me = {os.getpid(), *(p.pid for p in psutil.Process().parents())}
    hits: dict[int, psutil.Process] = {}
    for p in psutil.process_iter(["cmdline", "exe"]):
        if p.pid in me:
            continue
        try:
            cmd = subprocess.list2cmdline(p.info["cmdline"] or [])
            if str(root).lower() in cmd.lower() or _under(p.info["exe"], root) or _under(p.cwd(), root) \
                    or _fresh_env(p, root):
                hits[p.pid] = p
                for c in p.children(recursive=True):
                    if c.pid not in me:
                        hits[c.pid] = c
        except (psutil.Error, OSError):
            continue
    for p in hits.values():
        try:
            p.kill()
        except psutil.Error:
            pass
    psutil.wait_procs(list(hits.values()), timeout=10)
    return sorted(hits)


def fresh_owners(client) -> dict[int, float]:
    """{pid: create time} of the fresh client's child_server.py processes (owners of its hidden Excel)."""
    import psutil

    out = {}
    for pid in server_pids(client):
        try:
            out[pid] = psutil.Process(pid).create_time()
        except psutil.Error:
            continue
    return out


def kill_fresh_excels(owners: dict[int, float], root: Path = FRESH_ROOT) -> list[int]:
    """Kill the hidden Excels the fresh install's pids.json records for `owners` (the fresh client's servers): the
    same strict rule as engine_excel_pids, with the fresh registry; plus any Excel running with the fresh profile's
    environment (see _fresh_env). Normally the server already quit them."""
    import json

    import psutil

    try:
        registry = json.loads(fresh_paths(root)["pids"].read_text())
    except (OSError, ValueError):
        registry = []
    registry = [i for i in registry if isinstance(i, dict)] if isinstance(registry, list) else []
    killed = []
    for p in psutil.process_iter(["name", "cmdline", "create_time"]):
        try:
            if (p.info["name"] or "").lower() == "excel.exe" and (
                    is_engine_excel(subprocess.list2cmdline(p.info["cmdline"] or []), "", p.pid,
                                    p.info["create_time"], registry, owners) or _fresh_env(p, root)):
                p.kill()
                killed.append(p.pid)
        except (psutil.Error, TypeError):
            continue
    return killed


def _rm_junctions(root: Path) -> None:
    """Remove the junctions under `root` (the skill links) as links, so deleting the tree never follows one."""
    import os

    for dirpath, dirnames, _files in os.walk(root):
        for d in list(dirnames):
            full = os.path.join(dirpath, d)
            if os.path.isjunction(full) or os.path.islink(full):
                os.rmdir(full)
                dirnames.remove(d)


def remove_fresh(root: Path = FRESH_ROOT) -> None:
    """Stop the fresh processes and delete the fresh profile (only ever the .matrix-fresh folder)."""
    import os
    import stat

    if root.name != ".matrix-fresh" or not root.exists():
        return
    stop_fresh_processes(root)
    _rm_junctions(root)

    def _ro(func, path, _exc):
        os.chmod(path, stat.S_IWRITE)
        func(path)
    for _ in range(5):
        try:
            shutil.rmtree(root, onexc=_ro)
        except OSError:
            pass
        if not root.exists():
            return
        time.sleep(2)
    raise OSError(f"could not delete {root}")


def doctor_lines(text: str) -> dict[str, tuple[str, str]]:
    """{check: (status, detail)} from `smartass doctor` text output."""
    out = {}
    for ln in text.splitlines():
        m = re.match(r"\s{2}(OK|WARN|FAIL)\s+(\S+)\s*(.*)$", ln)
        if m:
            out[m.group(2)] = (m.group(1), m.group(3))
    return out


# ---------------------------------------------------------------- S10 helpers
SV_OPTIONS_KEY = r"Software\Hyperion Solutions\HyperionSmartView\Options"
SV_OPTION_NAMES = ("messageDisplay", "bDisplayHomePanelOnStartup", "bCommentDlgEnable")  # GLOBAL_OPTIONS 104/109/110


def _sv_options_raw() -> tuple[str, int]:
    import winreg

    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, SV_OPTIONS_KEY) as k:
        return winreg.QueryValueEx(k, "CAOptionsXML")


def sv_options() -> dict[str, str | None]:
    """The Smart View global options the engine changes (messages, home panel, comment dialog), from CAOptionsXML."""
    xml, _ = _sv_options_raw()
    out = {}
    for name in SV_OPTION_NAMES:
        m = re.search(rf'<option key="{name}" value="([^"]*)"\s*/>', xml)
        out[name] = m.group(1) if m else None
    return out


def write_sv_options(values: dict[str, str | None]) -> None:
    """Put `values` back into CAOptionsXML (only those keys; the value type is kept)."""
    import winreg

    xml, kind = _sv_options_raw()
    for name, v in values.items():
        if v is not None:
            xml = re.sub(rf'(<option key="{name}" value=")[^"]*("\s*/>)', rf"\g<1>{v}\g<2>", xml)
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, SV_OPTIONS_KEY, 0, winreg.KEY_SET_VALUE) as k:
        winreg.SetValueEx(k, "CAOptionsXML", 0, kind, xml)


def visible_excel_pids() -> set[int]:
    """Excel processes with a visible main window (XLMAIN): the user's and the matrix's interactive Excels."""
    import psutil

    pids = set()
    for p in psutil.process_iter(["name"]):
        if (p.info["name"] or "").lower() == "excel.exe":
            pids.add(p.pid)
    return {pid for _h, pid, cls, _t in _windows(pids) if cls == "XLMAIN"}


class _PsProc:
    """psutil.Process with the subprocess.Popen bits _quit_scratch_excel uses."""

    def __init__(self, pid: int):
        import psutil

        self.pid, self._p = pid, psutil.Process(pid)

    def poll(self):
        return None if self._p.is_running() else 0

    def wait(self, timeout):
        import psutil

        try:
            return self._p.wait(timeout)
        except psutil.TimeoutExpired:
            raise subprocess.TimeoutExpired("excel", timeout) from None

    def kill(self):
        import psutil

        try:
            self._p.kill()
        except psutil.Error:
            pass


def quit_owned_excel(pid: int, scratch) -> None:
    """Close an interactive Excel the matrix owns ('/x <scratch workbook>', engine_excel_pids): its _matrix books
    unsaved, then Quit (killed if it hangs). Any other Excel raises; one holding a non-matrix book is left open."""
    if pid not in engine_excel_pids(scratch):
        raise AssertionError(f"Excel {pid} is not owned by the matrix")
    _quit_scratch_excel(_PsProc(pid))


def close_scratch(t) -> None:
    """Spike cleanup: close the scratch workbook, quit its Excel if nothing but add-ins is left open (an Excel the
    spike started; a user's Excel always has their own books), then delete the file (retried: Excel/OneDrive let go
    of it a moment after the close)."""
    from smartview_mcp.excel.jobs import com
    path = Path(str(t.path))
    app = t.app
    com(lambda: t.wb.Close(False))
    try:
        if all(com(lambda: wb.IsAddin) for wb in com(lambda: list(app.Workbooks))):
            com(app.Quit)
    except Exception:   # noqa: BLE001 - Excel may already be gone
        pass
    for _ in range(20):
        try:
            path.unlink(missing_ok=True)
            return
        except PermissionError:
            time.sleep(0.5)
    print(f"cleanup: could not delete {path.name}")
