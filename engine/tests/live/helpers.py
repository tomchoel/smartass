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

from smartview_mcp.auditlib import is_engine_excel, poll_kill, stale_suspended

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
    """The engine's pids.json (written by host._PidRegistry: [{pid, ctime, owner, owner_ctime}])."""
    import json

    from smartview_mcp import config

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


def kill_engine_edge_when_seen(done, timeout: float = 60, interval: float = 0.5) -> tuple[list[int], bool]:
    """Poll for the drill Edge for up to `timeout` s and kill it the moment it appears; stops early once `done()`
    (the call under test finished). Returns (killed pids, the call was still running at the kill)."""
    return poll_kill(kill_engine_edge, done, timeout, interval, time.sleep, time.monotonic)


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


# ================================================================== S8 injections (each undoes itself in finally)
BAK = ".matrix-bak"
COOKIE_FILES = ("Cookies", "Cookies-journal")


def _cookie_dirs() -> list[Path]:
    """Cookie folders of the engine's drill Edge profiles (drill_edge, drill_edge_<channel>)."""
    from smartview_mcp import config

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
    from smartview_mcp import config

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
    from smartview_mcp import config

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
    from smartview_mcp import config

    return Path(config.ANALYSER_DIR) / "_matrix"
