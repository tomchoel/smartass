"""Working inside the USER's interactive Excel (any open workbook), on the server's COM thread.

Proven on 2026-10-13 (spikes/spike_hwnd.py):
- attach through the Excel main window (AccessibleObjectFromWindow on XLMAIN>XLDESK>EXCEL7), never GetObject(path) -
  that spawns a hidden '-Embedding' Excel when the workbook is not in the ROT yet;
- load the VBA bridge as a hidden add-in (svbridge.xlam) into that Excel; add-ins resolve by name, they are not enumerated;
- never send COM calls while Smart View's 'Login' dialog is open in that Excel (deadlock) - wait for the SSO watcher;
- Smart View acts on the ACTIVE workbook/sheet, so the target sheet is activated (the user's view jumps to it).
The user's Excel is NEVER killed or saved by this module.
"""
from __future__ import annotations

import contextlib
import datetime as dt
import hashlib
import json
import logging
import os
import re
import subprocess
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path

import psutil
import pythoncom
import pywintypes
import win32com.client as wc
import win32file
import win32gui
import win32process

from . import config
from .grid import Axis, Layout, full_pov, parse
from .grid import verify_labels as _verify_labels   # pure; moved to grid so the HTTP route (svhttp) uses it too
from .host import ExcelHost, HostError, InputError

WM_GETOBJECT, OBJID_NATIVEOM = 0x003D, -16
RETRY = {-2147418111, -2147417846}            # call rejected / retry later (edit mode, busy)
EDIT_WAIT_S = 30


log = logging.getLogger("smartview.userexcel")

# ------------------------------------------------------------------ low level
VBA_E_IGNORE = -2146777998                   # 0x800AC472: Excel busy, delivered inside DISP_E_EXCEPTION

# Every user-Excel job runs on its own STA thread with its own token. When the server abandons a job stuck in the
# user's Excel, only that job's token is aborted: its next COM call raises (so it rolls back instead of silently
# finishing and duplicating a retried analysis), and jobs running in parallel in other Excels are not affected.
_job = threading.local()
_all_tokens: "list[JobToken]" = []


class JobAborted(HostError):
    pass


class JobToken:
    def __init__(self):
        self.aborted = False


def bind_job(token: "JobToken | None" = None) -> None:
    _job.token = token or JobToken()
    _job.turns = []
    _all_tokens.append(_job.token)
    del _all_tokens[:-50]


def abort_running_jobs() -> None:
    for tk in _all_tokens:
        tk.aborted = True


def _check_alive():
    tk = getattr(_job, "token", None)
    if tk is not None and tk.aborted:
        raise JobAborted("This request was abandoned after a timeout; its partial sheets are being removed.")


def _busy(e: pywintypes.com_error) -> bool:
    inner = e.excepinfo[5] if e.excepinfo and len(e.excepinfo) > 5 else None
    return e.hresult in RETRY or e.hresult == VBA_E_IGNORE or inner == VBA_E_IGNORE


def com(call, limit: float = 60.0, check: bool = True):
    if check:
        _check_alive()
    t0, d = time.monotonic(), 0.1
    while True:
        try:
            return call()
        except pywintypes.com_error as e:
            if _busy(e) and time.monotonic() - t0 < limit:
                pythoncom.PumpWaitingMessages()
                time.sleep(d)
                d = min(d * 2, 2.0)
                continue
            if _busy(e):
                raise HostError("Excel is busy - probably a cell is being edited or a dialog is open. Press Enter or "
                                "Esc in Excel and ask again.") from None
            raise


TURN_WAIT_S = 600
MAX_EXCELS = 4                     # the user's Excel(s) + extra windows started for parallel agents


class Turn:
    """Turn-taking for ONE Excel process, across threads, servers and agents (a lock file per Excel pid). One Excel
    runs all automation on its single UI thread and Smart View works on the active sheet, so two writers in the same
    Excel would only interleave and could put numbers on each other's sheets. Different Excels run in parallel."""

    def __init__(self, pid: int):
        import msvcrt
        self._m = msvcrt
        config.SVKIT.mkdir(parents=True, exist_ok=True)
        self.pid = pid
        self.f = open(config.SVKIT / f"user_excel_{pid}.lock", "a+b")
        self.held = False

    def try_acquire(self) -> bool:
        try:
            self.f.seek(0)
            self._m.locking(self.f.fileno(), self._m.LK_NBLCK, 1)
            self.held = True
        except OSError:
            pass
        return self.held

    def acquire(self, what: str) -> float:
        t0 = time.monotonic()
        while not self.try_acquire():
            if time.monotonic() - t0 > TURN_WAIT_S:
                self.close()
                raise HostError(f"Another agent has been working in that Excel for {TURN_WAIT_S // 60} min; "
                                f"{what} was not started. Try again shortly.")
            time.sleep(0.25)
        return time.monotonic() - t0

    def close(self):
        if self.held:
            try:
                self.f.seek(0)
                self._m.locking(self.f.fileno(), self._m.LK_UNLCK, 1)
            except OSError:
                pass
            self.held = False
        self.f.close()


def hold_turn(pid: int, wait: bool = True, what: str = "this request") -> bool:
    """Take this job's turn in Excel `pid` (kept until the job ends). wait=False: only if it is free right now."""
    turns = getattr(_job, "turns", None)
    if turns is None:
        turns = _job.turns = []
    if any(tt.pid == pid for tt in turns):
        return True
    tt = Turn(pid)
    if not wait:
        if not tt.try_acquire():
            tt.close()
            return False
    else:
        from .timing import stage
        with stage("wait for that Excel (another agent)"):
            waited = tt.acquire(what)
        if waited > 1:
            log.info("waited %.1fs for Excel pid %s (another agent)", waited, pid)
    turns.append(tt)
    return True


def release_turns() -> None:
    for tt in getattr(_job, "turns", None) or []:
        tt.close()
    _job.turns = []


@contextlib.contextmanager
def excel_turn(what: str = "a request", pid: int = 0):
    """Context form of a turn (tests, tools that hold one Excel for a block)."""
    tt = Turn(pid)
    waited = tt.acquire(what)
    try:
        yield waited
    finally:
        tt.close()


@contextlib.contextmanager
def fast_excel(app):
    """No screen redraw and no automatic recalculation while we write/format (the user's Excel otherwise repaints
    and recalculates after every single cell: measured 17 s vs 1.6 s for one answer sheet). The user's own settings
    are restored afterwards, also on failure; the caller recalculates explicitly when it needs results."""
    prev_su = prev_calc = None
    try:
        prev_su = com(lambda: app.ScreenUpdating, check=False)
        prev_calc = com(lambda: app.Calculation, check=False)
        com(lambda: setattr(app, "ScreenUpdating", False), check=False)
        if prev_calc != -4135:                                  # xlCalculationManual
            com(lambda: setattr(app, "Calculation", -4135), check=False)
    except Exception:   # noqa: BLE001 - e.g. no workbook open: just run normally
        pass
    try:
        yield
    finally:
        for name, val in (("Calculation", prev_calc), ("ScreenUpdating", prev_su)):
            if val is not None:
                try:
                    com(lambda: setattr(app, name, val), limit=30, check=False)
                except Exception:   # noqa: BLE001
                    log.warning("could not restore Application.%s=%s", name, val)


def rollback(t: "Target") -> list[str]:
    """Delete the sheets this request added (after a failure), so no half-built analysis is left behind.
    Runs even for an abandoned job (no generation check) and restores the user's DisplayAlerts setting."""
    removed = []
    prev = None
    try:
        prev = com(lambda: t.app.DisplayAlerts, limit=30, check=False)
        com(lambda: setattr(t.app, "DisplayAlerts", False), limit=30, check=False)
        for name in reversed(t.sheets_added):
            try:
                com(lambda: t.wb.Worksheets(name).Delete(), limit=30, check=False)
                removed.append(name)
            except Exception:   # noqa: BLE001
                pass
    except Exception:   # noqa: BLE001
        pass
    finally:
        try:
            if prev is not None:
                com(lambda: setattr(t.app, "DisplayAlerts", prev), limit=10, check=False)
            com(lambda: setattr(t.app, "StatusBar", False), limit=5, check=False)
        except Exception:   # noqa: BLE001
            pass
    t.sheets_added.clear()
    return removed


def _is_embedding(pid: int) -> bool:
    try:
        return "-embedding" in " ".join(psutil.Process(pid).cmdline()).lower()
    except psutil.Error:
        return True


def login_open(pid: int) -> bool:
    found = []

    def cb(h, _):
        if (win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == "#32770"
                and win32gui.GetWindowText(h) == "Login" and win32process.GetWindowThreadProcessId(h)[1] == pid):
            found.append(h)
        return True
    win32gui.EnumWindows(cb, None)
    return bool(found)


BOX_POLL_S = 0.5


class SmartViewBoxes:
    """While the engine drives Smart View in a visible Excel: a thread clicks OK on Smart View's modal message boxes
    in that Excel (they would block the call until the user clicks) and keeps their texts for the caller."""

    def __init__(self, pid: int):
        self.pid, self.texts = pid, []
        self._stop, self._seen = threading.Event(), set()
        self._thread = threading.Thread(target=self._run, name=f"sv-boxes-{pid}", daemon=True)

    def _run(self):
        from . import dialogs
        while not self._stop.wait(BOX_POLL_S):
            try:
                boxes = dialogs.sv_boxes(dialogs.scan(self.pid))
                self._seen &= {d.hwnd for d in boxes}     # a reused window handle is a new box
                for d in boxes:
                    if d.hwnd not in self._seen:
                        self._seen.add(d.hwnd)
                        if dialogs.press_ok(d.hwnd):      # information boxes only, never a question
                            self.texts.append(d.text)
                            log.warning("Smart View message in Excel pid %s dismissed: %s", self.pid, d.text)
            except Exception:   # noqa: BLE001 - the guard must never break the pull
                log.debug("Smart View box guard failed", exc_info=True)

    def __enter__(self):
        self._thread.start()
        return self

    def __exit__(self, *exc):
        self._stop.set()
        self._thread.join(timeout=2)


NOT_SIGNED_IN = ("Smart View in this Excel is not signed in correctly: EPM answers 'Empty response from provider' to "
                 "every connect, and Smart View shows no Login to click. Ask the user to sign in once in that Excel "
                 "(Smart View ribbon -> Panel -> connect FinMax, or close that Excel so the next call starts a fresh "
                 "one), then ask again. If every Excel does this and no Login appears anywhere, Smart View's shared "
                 "sign-in is stale: close all Excels, rename the Cookies files in %APPDATA%\\Oracle\\SmartView\\WebView"
                 "\\EBWebView\\Default\\Network, then retry (docs\\troubleshooting.md, Known limits).")


RECOVER_SETTLE_S = 12         # after a login recovery: time for Smart View to reopen the login before we go on


def recover_ie_login(pid: int, hwnds: list[int]) -> dict:
    """The login in Excel `pid` is on the Internet Explorer fallback because an earlier, unfinished WebView2 login
    still holds Smart View's WebView2 folder. End that stale WebView2 (this Excel's own, or an orphan - never another
    app's) and close the broken IE windows; Smart View then opens a new login in WebView2. Excel itself is untouched."""
    from . import dialogs
    ended = dialogs.end_stale_sv_webview2(pid)
    dialogs.close(hwnds)
    return {"ended": ended, "closed": hwnds}


def wait_login_cleared(pid: int, appear_grace: float = 0.0, limit: float = 90.0) -> None:
    """Block until no Smart View Login dialog is open in `pid` (the global SSO watcher clicks it). A login that fell
    back to Internet Explorer ('Script Error') never completes: it is recovered (recover_ie_login), at most twice."""
    from . import dialogs
    t0 = time.monotonic()
    seen = False
    recoveries, t_rec = 0, None
    while time.monotonic() - t0 < limit:
        now = time.monotonic()
        if login_open(pid):
            broken = dialogs.ie_fallback(dialogs.scan(pid))
            if not broken:                        # a real (WebView2) login: the watcher clicks it
                seen = True
            # a second recovery only after RECOVER_SETTLE_S: a new WebView2 login of this Excel must not be ended
            elif recoveries < 2 and (t_rec is None or now - t_rec >= RECOVER_SETTLE_S):
                recoveries += 1
                res = recover_ie_login(pid, broken)
                log.warning("Smart View login in Excel pid %s fell back to Internet Explorer; ended stale WebView2 %s "
                            "and closed %s so it can sign in again", pid, res["ended"], res["closed"])
                seen, t_rec = False, time.monotonic()   # Smart View reopens the login (in WebView2): wait for it
                continue
        elif t_rec is not None:
            if seen or now - t_rec >= RECOVER_SETTLE_S:   # the new login came and went, or none came back
                return
        elif seen or now - t0 >= appear_grace:
            return
        time.sleep(0.5)
    raise HostError("The Oracle EPM login dialog is still open in your Excel and the automatic SSO click did not "
                    "complete. Ask the user to click 'Telenor SSO' in it, then ask again. If it keeps coming back, "
                    "Smart View's WebView2 may be held by another Office app (often an old PowerPoint): closing that "
                    "app usually fixes it.")


@dataclass
class XlInstance:
    """One interactive Excel process. `app` is None when that Excel did not answer COM (a dialog is open, a cell is
    being edited, or it is busy); its workbook window captions (`captions`) are known either way."""
    pid: int
    hwnd: int                                        # front-most main window
    captions: list[str]
    app: object = None


def _excel_windows() -> list[tuple[int, int, list[tuple[int, str]]]]:
    """Thin Win32 part: [(pid, front-most XLMAIN hwnd, [(EXCEL7 hwnd, caption)])] of the visible Excel main windows,
    grouped by process, front-most process first (EnumWindows walks top-level windows in z-order)."""
    by_pid: dict[int, tuple[int, list[tuple[int, str]]]] = {}

    def cb(h, _):
        try:
            if win32gui.GetClassName(h) != "XLMAIN" or not win32gui.IsWindowVisible(h):
                return True
            pid = win32process.GetWindowThreadProcessId(h)[1]
            desk = win32gui.FindWindowEx(h, 0, "XLDESK", None)
            ex7 = win32gui.FindWindowEx(desk, 0, "EXCEL7", None) if desk else 0
            entry = by_pid.setdefault(pid, (h, []))
            if ex7:
                entry[1].append((ex7, win32gui.GetWindowText(ex7)))
        except pywintypes.error:
            pass
        return True
    win32gui.EnumWindows(cb, None)
    return [(pid, h, ex7s) for pid, (h, ex7s) in by_pid.items()]


def _app_from_window(ex7: int):
    """Excel.Application behind a workbook window (AccessibleObjectFromWindow), or None if that Excel refuses."""
    try:
        _, lres = win32gui.SendMessageTimeout(ex7, WM_GETOBJECT, 0, OBJID_NATIVEOM, 2, 3000)
        win = wc.Dispatch(pythoncom.ObjectFromLresult(lres, pythoncom.IID_IDispatch, 0))
        return win.Application
    except Exception:   # noqa: BLE001 - modal dialog, cell edit mode, busy, or not a workbook window
        return None


def engine_excel_pids() -> set[int]:
    """The hidden automation Excels the smartass servers started (pids.json): never a user's Excel."""
    from .host import _PidRegistry
    return {int(i["pid"]) for i in _PidRegistry().load() if isinstance(i, dict) and i.get("pid")}


def excel_instances(_retry_login: bool = True) -> list[XlInstance]:
    """Every visible, user-facing Excel process (not the engine's hidden ones), most recently active first, including
    Excels that do not answer COM right now (app=None). An Excel showing the EPM Login dialog is waited for (the SSO
    watcher clicks it) before it is touched."""
    out, pending = [], []
    engine = engine_excel_pids()
    for pid, h, ex7s in _excel_windows():
        if pid in engine or _is_embedding(pid):
            continue
        if login_open(pid):
            pending.append(pid)                   # no COM into an Excel whose EPM login is open (deadlock)
            continue
        if not ex7s:
            continue
        app = None
        for ex7, _cap in ex7s:                    # some windows (e.g. an add-in shown read-only) give no object
            app = _app_from_window(ex7)
            if app is not None:
                break
        out.append(XlInstance(pid, h, [c for _e, c in ex7s], app))
    if pending and _retry_login:
        from . import sso
        sso.ensure_global_watcher()
        for pid in pending:
            wait_login_cleared(pid)
        return excel_instances(_retry_login=False)
    return out


def interactive_excels(_retry_login: bool = True) -> list[tuple[int, int, object]]:
    """[(pid, main hwnd, Application)] of the visible, user-started Excel instances that answer COM, most recently
    active first (excel_instances() also lists the ones that are busy)."""
    return [(i.pid, i.hwnd, i.app) for i in excel_instances(_retry_login) if i.app is not None]


def _caption_book(caption: str) -> str:
    """Workbook name in a window caption: 'Book  -  Read-Only  -  Compatibility Mode' / 'Book.xlsx:2' -> 'book'."""
    s = str(caption).split("  -  ")[0].strip()
    s = re.sub(r":\d+$", "", s)
    return re.sub(r"\.(xlsx|xlsm|xlsb|xls|xltx|xltm)$", "", s, flags=re.I).lower()


def busy_holder(instances: list[XlInstance], name: str) -> XlInstance | None:
    """The Excel that does not answer COM but shows a window for workbook `name` (a file name or path)."""
    want = _caption_book(Path(str(name)).name)
    return next((i for i in instances if i.app is None and any(_caption_book(c) == want for c in i.captions)), None)


def busy_error(pid: int, name: str, found: list | None = None) -> HostError:
    """Name what blocks the Excel holding `name`: a dialog (dialogs.scan) or else cell edit mode."""
    from . import dialogs
    if found is None:
        try:
            found = dialogs.scan(pid)
        except Exception:   # noqa: BLE001
            found = []
    if found:
        return HostError(f"'{name}' is open in an Excel (pid {pid}) that has a dialog open "
                         f"({dialogs.describe(found[0])}). Close that dialog in Excel, then ask again.")
    return HostError(f"'{name}' is open in an Excel (pid {pid}) that is in cell edit mode or busy (no dialog seen). "
                     "Press Enter or Esc in that Excel window, then ask again.")


SHARING_VIOLATION = {32, 33}        # ERROR_SHARING_VIOLATION, ERROR_LOCK_VIOLATION


def _file_in_use(path: Path) -> bool:
    """Another process holds the file open (Excel locks the workbooks it has open): only a sharing/lock violation
    counts. A read-only file (access denied, e.g. a SharePoint file with read rights only) is NOT in use - Excel
    opens it read-only."""
    try:
        h = win32file.CreateFile(str(path), win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                                 win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE,
                                 None, win32file.OPEN_EXISTING, 0, None)
    except pywintypes.error as e:
        return e.winerror in SHARING_VIOLATION
    h.Close()
    return False


def _locked_error(path: Path) -> HostError:
    return HostError(f"'{path.name}' is open in another program, or in an Excel this engine cannot reach (the file is "
                     "locked). Close it there (or save and close it in that Excel), then ask again.")


ARCHIVE_URL_MARK = "/sites/analysetelenornorgeas/"     # SharePoint site behind the Produktregnskapet archive


def _onedrive_mounts() -> list[tuple[str, str]]:
    """[(url namespace, local mount)] from OneDrive's sync registry, longest URL first."""
    import winreg
    out = []
    try:
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\SyncEngines\Providers\OneDrive")
    except OSError:
        return out
    i = 0
    while True:
        try:
            sub = winreg.EnumKey(root, i)
        except OSError:
            break
        i += 1
        try:
            k = winreg.OpenKey(root, sub)
            url = winreg.QueryValueEx(k, "UrlNamespace")[0]
            mount = winreg.QueryValueEx(k, "MountPoint")[0]
            if url and mount:
                out.append((url, mount))
        except OSError:
            continue
    return sorted(set(out), key=lambda x: -len(x[0]))


def local_path(full_name: str) -> str:
    """Workbook.FullName of a OneDrive/SharePoint file is an https URL; map it to the synced local path."""
    s = str(full_name)
    if not s.lower().startswith("http"):
        return s
    from urllib.parse import unquote
    for url, mount in _onedrive_mounts():
        if s.lower().startswith(url.lower()):
            rest = unquote(s[len(url):]).replace("/", os.sep)
            return os.path.join(mount, rest)
    return s


def _norm(p) -> str:
    return os.path.normcase(os.path.abspath(local_path(str(p))))


def is_archive(path) -> bool:
    s = str(path)
    if ARCHIVE_URL_MARK in s.replace("\\", "/").lower():
        return True
    local = local_path(s)
    if not os.path.isabs(local):                  # e.g. an unsaved 'Book1': not a file anywhere
        return False
    try:
        return _norm(local).startswith(_norm(config.ARCHIVE_ROOT) + os.sep)
    except Exception:   # noqa: BLE001
        return False


# ------------------------------------------------------------------ bridge add-in
def _bridge_hash() -> str:
    h = hashlib.sha1()
    for p in config.BAS_FILES:
        h.update(p.read_bytes())
    return h.hexdigest()[:12]


def bridge_path(ver: str) -> Path:
    """Versioned file name: a new build never collides with an older add-in still loaded in some Excel."""
    return config.SVKIT / f"svbridge_{ver}.xlam"


def ensure_bridge_file(h: ExcelHost) -> str:
    """Build svbridge_<hash>.xlam with the hidden server Excel when the .bas sources changed; prune old builds."""
    ver = _bridge_hash()
    path = bridge_path(ver)
    for old in config.SVKIT.glob("svbridge*.xlam"):
        if old != path:
            try:
                old.unlink()
            except OSError:
                pass                                 # still loaded in an Excel; removed next time
    if path.exists():
        return ver
    return h.on_host(lambda: _build_bridge(h, ver, path))   # the hidden Excel lives on the host thread


def _build_bridge(h: ExcelHost, ver: str, path: Path) -> str:
    if path.exists():                                # a parallel job built it meanwhile
        return ver
    from .ops import _label
    wb = h.com(h.app.Workbooks.Add)
    try:
        for p in config.BAS_FILES:
            m = wb.VBProject.VBComponents.Add(1)
            m.Name = "SmartViewVBA" if p.stem == "smartview_body" else p.stem
            m.CodeModule.AddFromString(p.read_text(encoding="cp1252"))
        m = wb.VBProject.VBComponents.Add(1)
        m.Name = "modBridgeVersion"
        m.CodeModule.AddFromString(f'Public Function B_BridgeVersion() As String\n    B_BridgeVersion = "{ver}"\nEnd Function\n')
        _label(wb)
        h.com(lambda: wb.SaveAs(str(path), 55))
    finally:
        h.com(lambda: wb.Close(False))
    if not path.exists():
        raise HostError("Could not build the Smart View bridge add-in (svbridge.xlam).")
    return ver


# ------------------------------------------------------------------ target workbook
@dataclass
class Target:
    app: object
    wb: object
    pid: int
    hwnd: int
    bridge: object = None
    created: bool = False
    sheets_added: list[str] = field(default_factory=list)

    @property
    def name(self) -> str:
        return self.wb.Name

    @property
    def path(self) -> str:
        return local_path(self.wb.FullName)

    @property
    def full_name(self) -> str:
        return str(self.wb.FullName)

    def run(self, fn: str, *args):
        """Bridge call. Smart View acts on the ACTIVE workbook/sheet, so the target sheet is activated right before
        every call - if the user clicked into another workbook meanwhile, we act on ours, not theirs."""
        if args and isinstance(args[0], str) and fn not in ("B_Ping", "B_SetOption", "B_GetOption", "B_LastError"):
            name = args[0]
            try:
                ws = com(lambda: self.wb.Worksheets(name))
                com(self.wb.Activate)
                com(ws.Activate)
            except pywintypes.com_error:
                pass
        return com(lambda: self.app.Run(f"'{self.bridge.Name}'!modBridge.{fn}", *args))

    def status(self, text: str | None):
        try:
            com(lambda: setattr(self.app, "StatusBar", ("Claude: " + text) if text else False), limit=5)
        except Exception:   # noqa: BLE001 - cosmetic
            pass


def _wait_ready(app, limit: float = EDIT_WAIT_S, pid: int | None = None, name: str | None = None):
    t0 = time.monotonic()
    while time.monotonic() - t0 < limit:
        try:
            if app.Ready:
                return
        except pywintypes.com_error as e:
            if not _busy(e):
                raise
        time.sleep(0.5)
    if pid:
        raise busy_error(pid, name or "The workbook")
    raise HostError("Your Excel is in edit mode (a cell is being typed in) or showing a dialog. Press Enter or Esc, "
                    "then ask again.")


def _attach_bridge(h: ExcelHost, t: Target):
    ver = ensure_bridge_file(h)
    path = bridge_path(ver)
    br = None
    try:
        br = com(lambda: t.app.Workbooks(path.name), limit=20)    # this exact version already loaded
    except pywintypes.com_error as e:
        if e.hresult != -2147352567:                 # anything but "not in collection" is a real error
            raise
    if br is None:
        br = com(lambda: t.app.Workbooks.Open(str(path)), limit=120)
    t.bridge = br
    if t.run("B_Ping") != "pong":
        raise HostError("The Smart View bridge add-in did not load in your Excel.")


ARCHIVE_REFUSAL = ("'{name}' is in the Produktregnskapet archive (audit record) and is not written to. "
                   "Use target='new', or allow it explicitly for this one request (allow_archive=true).")


def resolve_target(h: ExcelHost, target: str = "session", allow_archive: bool = False, title: str | None = None) -> Target:
    """target: 'session' (default: this session's own analysis workbook - created on its first question, reopened
    if closed), 'active' (the workbook the user is working in), 'new' (a new file, which becomes the session's
    workbook), or an absolute file path."""
    from . import session, sso
    sso.ensure_global_watcher()
    if target == "session":
        wb_path = session.workbook()
        if wb_path:
            t = resolve_target(h, wb_path, allow_archive, title)
            t.created = True                          # the session's own file: saved after each answer
            return t
        target = "new"
    if target not in ("active", "new"):
        p = Path(target)
        if not p.is_absolute():
            raise InputError(f"Give an absolute path for target (got '{target}'), or use 'session' / 'active' / 'new'.")
        if is_archive(p) and not allow_archive:       # refuse BEFORE opening anything
            raise InputError(ARCHIVE_REFUSAL.format(name=p.name))
    from .timing import stage
    with stage("find Excel windows"):
        instances = excel_instances()                 # waits for any open EPM login before touching an Excel
    excels = [(i.pid, i.hwnd, i.app) for i in instances if i.app is not None]
    if target == "active":
        for pid, hwnd, app in excels:
            try:
                wb = com(lambda: app.ActiveWorkbook, limit=20)
            except pywintypes.com_error:
                wb = None
            if wb is not None and not str(wb.Name).lower().endswith(".xlam"):
                hold_turn(pid, what="this request")   # same workbook: agents take turns
                return _finish_target(h, Target(app, wb, pid, hwnd), allow_archive)
        target = "new"                                # nothing open: fall back to a new analysis workbook
    if target == "new":
        with stage("new workbook"):
            t = _new_target(h, title or "Analyse", excels, {i.pid for i in instances})
        with stage("ready + bridge"):
            t = _finish_target(h, t, allow_archive)
        session.set_workbook(t.path)                  # a new file becomes this session's workbook
        return t
    with stage("open workbook"):
        t = _open_path(h, Path(target), excels, instances)
    with stage("ready + bridge"):
        return _finish_target(h, t, allow_archive)


def _new_target(h: ExcelHost, title: str, excels, known: set[int] | None = None) -> Target:
    """A new analysis workbook in a FREE Excel, so parallel agents do not wait for each other: the first Excel whose
    turn is free; else (below MAX_EXCELS) a new Excel window started on the TN template; else wait for the first.
    If it cannot be saved, the new workbook is closed unsaved (and an Excel started for it is closed again)."""
    from . import styling
    from .timing import stage
    tpl = styling.template_path()
    choice = next((e for e in excels if hold_turn(e[0], wait=False)), None)
    wb = None
    spawned = False
    if choice is None and excels and len(excels) < MAX_EXCELS and tpl.exists():
        with stage("start another Excel (parallel agent)"):
            choice, wb = _spawn_excel(tpl, {e[0] for e in excels} | (known or set()))
        spawned = True
    if choice is None and excels:
        hold_turn(excels[0][0], what="this request")
        choice = excels[0]
    if choice is None:                                # no Excel at all: the hidden one creates the file, /x opens it
        t = _open_path(h, _new_analysis_file(h, title), excels)
        t.created = True
        return t
    pid, hwnd, app = choice
    try:
        if wb is None:
            with stage("add from template"):
                try:                                  # from the TN template: the styles come for free (saves ~11 s)
                    tpl_s = str(styling.ensure_template(app))
                    wb = com(lambda: app.Workbooks.Add(tpl_s))
                except Exception:   # noqa: BLE001 - template unavailable: plain workbook, styles installed later
                    wb = com(lambda: app.Workbooks.Add())
        with stage("save as"):
            _save_new(app, wb, title)
    except BaseException:
        _discard_new(app, wb, spawned)
        raise
    t = Target(app, wb, pid, hwnd)
    t.created = True
    return t


def _save_new(app, wb, title: str) -> Path:
    """Name the first sheet 'Kilder', label and SaveAs the new workbook under a file name that is free on disk and
    not the name of a workbook already open in that Excel (Excel refuses both with 0x800A03EC)."""
    from .ops import _label
    try:
        taken = {str(b.Name).lower() for b in com(lambda: app.Workbooks)}
    except Exception:   # noqa: BLE001
        taken = set()
    path = _new_path(title, taken)
    com(lambda: setattr(wb.Worksheets(1), "Name", "Kilder"))
    com(lambda: _label(wb))
    try:
        com(lambda: wb.SaveAs(str(path), 51))
    except pywintypes.com_error as e:
        raise HostError(f"Excel could not save the new workbook as {path} ({_com_text(e)}). Nothing was left open; "
                        "check that the folder exists and is writable, then ask again.") from None
    if not path.exists():
        raise HostError(f"Could not create {path} (sensitivity label?).")
    return path


def _com_text(e: pywintypes.com_error) -> str:
    info = e.excepinfo if e.excepinfo and len(e.excepinfo) > 2 else None
    return str(info[2]) if info and info[2] else str(e)


def _discard_new(app, wb, spawned: bool) -> None:
    """Close a new workbook that could not be saved, without saving; quit an Excel started only for it."""
    if wb is not None:
        try:
            com(lambda: wb.Close(False), limit=30, check=False)
        except Exception:   # noqa: BLE001
            log.warning("could not close the unsaved new workbook")
    if spawned:
        try:
            left = [b for b in com(lambda: app.Workbooks, limit=10, check=False)
                    if not str(b.Name).lower().endswith((".xlam", ".xla"))]
            if not left:
                com(lambda: app.Quit(), limit=10, check=False)
        except Exception:   # noqa: BLE001
            pass


def _spawn_excel(tpl: Path, known: set[int]):
    """Start a new interactive Excel on the TN template (opens as a new unsaved workbook) and take its turn."""
    proc = _launch_excel(tpl)
    t0 = time.monotonic()
    while time.monotonic() - t0 < 120:
        time.sleep(1)
        for pid, hwnd, app in interactive_excels():
            if pid in known:
                continue
            for wb in app.Workbooks:
                if str(wb.Name).lower().startswith(tpl.stem.lower()):
                    hold_turn(pid, wait=False)
                    wait_login_cleared(pid, appear_grace=12)     # a fresh Excel may show the EPM login at start-up
                    log.info("started Excel pid %s for a parallel agent", pid)
                    return (pid, hwnd, app), wb
    raise HostError(f"A new Excel (pid {proc.pid}) did not open the TN template within 120 s.")


def _finish_target(h: ExcelHost, t: Target, allow_archive: bool) -> Target:
    wait_login_cleared(t.pid)                         # before any further COM into that Excel
    if (is_archive(t.path) or is_archive(t.full_name)) and not allow_archive:
        raise InputError(ARCHIVE_REFUSAL.format(name=t.name))
    _wait_ready(t.app, pid=t.pid, name=t.name)
    _attach_bridge(h, t)
    return t


def _find_open(excels, path: Path) -> Target | None:
    for pid, hwnd, app in excels:
        for wb in app.Workbooks:
            if _norm(wb.FullName) == _norm(path):
                return Target(app, wb, pid, hwnd)
    return None


BUSY_WAIT_S = 10                    # a busy Excel that holds the target gets this long to answer before we say why


def _wait_busy_holder(pid: int, path: Path) -> Target | None:
    """The target seems open in Excel `pid` (a window caption shows its name), which does not answer COM (dialog, cell
    edit mode, busy). Take that Excel's turn (another agent may be working in it), give it BUSY_WAIT_S to answer, else
    name what blocks it. None when it answers but does not hold `path` (a same-named workbook from another folder)."""
    hold_turn(pid, what="this request")
    t0 = time.monotonic()
    while True:
        insts = [i for i in excel_instances() if i.pid == pid]
        if not insts:
            raise HostError(f"The Excel (pid {pid}) that had '{path.name}' open has closed. Ask again.")
        i = insts[0]
        if i.app is not None:
            return _find_open([(i.pid, i.hwnd, i.app)], path)
        if time.monotonic() - t0 >= BUSY_WAIT_S:
            raise busy_error(pid, path.name)
        time.sleep(0.5)


def _open_path(h: ExcelHost, path: Path, excels, instances: list[XlInstance] | None = None) -> Target:
    """The workbook at `path`, wherever it is open: any Excel that answers, else a busy Excel showing its window
    (which is named if it stays busy); else opened in the user's running Excel; else in a new Excel."""
    if not path.exists():
        raise InputError(f"File not found: {path}")
    t = _find_open(excels, path)
    if t is not None:
        hold_turn(t.pid, what="this request")
        return t
    if instances is None:
        instances = excel_instances()
    holder = busy_holder(instances, path.name)
    if holder is not None:
        t = _wait_busy_holder(holder.pid, path)
        if t is not None:
            return t
    if _file_in_use(path):
        raise _locked_error(path)
    if excels:                                  # open it in the user's Excel that is already running
        pid, hwnd, app = excels[0]
        hold_turn(pid, what="this request")
        try:
            wb = com(lambda: app.Workbooks.Open(str(path)), limit=180)
        except pywintypes.com_error as e:
            if "cannot access" in str(e).lower():
                raise _locked_error(path) from None
            raise
        return Target(app, wb, pid, hwnd)
    proc = _launch_excel(path)
    t0 = time.monotonic()
    while time.monotonic() - t0 < OPEN_TIMEOUT_S:
        time.sleep(1)
        for pid, hwnd, app in interactive_excels():
            for wb in app.Workbooks:
                if _norm(wb.FullName) == _norm(path):
                    hold_turn(pid, what="this request")
                    wait_login_cleared(pid, appear_grace=12)     # a fresh Excel shows the EPM login at start-up
                    return Target(app, wb, pid, hwnd)
        if proc.poll() is not None:
            raise HostError(f"Excel (pid {proc.pid}) closed before it opened {path.name}. "
                            "Next step: open Excel yourself (any workbook), then ask again.")
    raise HostError(f"Excel did not open {path} within {OPEN_TIMEOUT_S} s (pid {proc.pid}). "
                    "Next step: check that Excel window for a sign-in or OneDrive prompt, answer it, and ask again.")


# A cold Excel loads Smart View and the OneDrive file (sync, sensitivity label) before the workbook appears;
# on a busy PC that takes well over a minute.
OPEN_TIMEOUT_S = 240


def _launch_excel(path: Path) -> subprocess.Popen:
    """Start a new interactive Excel on path with a normal, visible window, so any start-up prompt can be seen
    (inherited start-up info otherwise left it minimised, hiding prompts from the user)."""
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 1                                # SW_SHOWNORMAL
    return subprocess.Popen([config.EXCEL_EXE, "/x", str(path)], startupinfo=si)


def _new_path(title: str, taken: set[str] | frozenset = frozenset()) -> Path:
    """analyser\\<project of this session>\\<date time> <title>.xlsx, with ' (2)', ' (3)' ... when that file exists
    (the same question again within the minute, e.g. a retry) or its name is in `taken` (open workbooks, lower case)."""
    from . import session
    safe = re.sub(r'[\\/:*?"<>|]+', "_", title).strip()[:60] or "Analyse"   # only characters Windows forbids
    stem = f"{dt.datetime.now():%Y-%m-%d_%H%M} {safe}"
    folder = session.project_dir()
    n = 1
    while True:
        p = folder / (f"{stem}.xlsx" if n == 1 else f"{stem} ({n}).xlsx")
        if not p.exists() and p.name.lower() not in taken:
            return p
        n += 1


def _new_analysis_file(h: ExcelHost, title: str) -> Path:
    path = _new_path(title)
    return h.on_host(lambda: _new_analysis_file_host(h, path))


def _new_analysis_file_host(h: ExcelHost, path: Path) -> Path:
    from .ops import _label
    wb = h.com(h.app.Workbooks.Add)
    try:
        wb.Worksheets(1).Name = "Kilder"
        _label(wb)
        h.com(lambda: wb.SaveAs(str(path), 51))
    finally:
        h.com(lambda: wb.Close(False))
    if not path.exists():
        raise HostError(f"Could not create {path} (sensitivity label?).")
    return path


# ------------------------------------------------------------------ question numbering and sheets
_Q = re.compile(r"^Q(\d+) (Svar|Answer|FinMax|FinLite|FinPlan|Linjer)\b")   # our sheets only, not the user's "Q3 Forecast"


def next_question(t: Target) -> int:
    nums = [int(m.group(1)) for s in t.wb.Worksheets if (m := _Q.match(s.Name))]
    return max(nums, default=0) + 1


def park_sheets(t: Target, names: list[str]) -> dict[str, str]:
    """Rename these sheets out of the way ('~<name>'). Returns {parked name: original name}."""
    parked = {}
    for name in names:
        if not _sheet_exists(t, name):
            continue
        s = com(lambda name=name: t.wb.Worksheets(name))
        tmp = unique_sheet_name(t, ("~" + name)[:31])
        com(lambda s=s, tmp=tmp: setattr(s, "Name", tmp))
        parked[tmp] = name
    return parked


def drop_parked(t: Target, parked: dict[str, str]) -> None:
    prev = com(lambda: t.app.DisplayAlerts, check=False)
    com(lambda: setattr(t.app, "DisplayAlerts", False), check=False)
    try:
        for tmp in parked:
            com(lambda tmp=tmp: t.wb.Worksheets(tmp).Delete(), check=False)
    finally:
        com(lambda: setattr(t.app, "DisplayAlerts", prev), check=False)


def restore_parked(t: Target, parked: dict[str, str]) -> None:
    for tmp, orig in parked.items():
        if _sheet_exists(t, tmp) and not _sheet_exists(t, orig):
            com(lambda tmp=tmp, orig=orig: setattr(t.wb.Worksheets(tmp), "Name", orig), check=False)


def _sheet_exists(t: Target, name: str) -> bool:
    return any(s.Name == name for s in t.wb.Worksheets)


def unique_sheet_name(t: Target, base: str) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", base)[:31]
    name, i = base, 2
    while _sheet_exists(t, name):
        suffix = f" {i}"
        name = base[:31 - len(suffix)] + suffix
        i += 1
    return name


def add_sheet(t: Target, name: str, after_q: int | None = None, color: int | None = None):
    """Add a sheet after the last sheet of question `after_q` (or at the end, before 'Kilder')."""
    sheets = list(t.wb.Worksheets)
    anchor = None
    if after_q is not None:
        group = [s for s in sheets if (m := _Q.match(s.Name)) and int(m.group(1)) == after_q]
        anchor = group[-1] if group else None
    if anchor is None:
        non_kilder = [s for s in sheets if s.Name not in ("Kilder", "Kilder SV", META_SHEET)]
        anchor = non_kilder[-1] if non_kilder else None
    # positional (Before, After): the keyword After= is silently ignored through late-bound COM here (measured
    # 15.10.2026: the sheet landed before the ACTIVE sheet instead). Only 'Kilder'/meta sheets yet: go first.
    ws = com(lambda: t.wb.Worksheets.Add(None, anchor) if anchor is not None else t.wb.Worksheets.Add(sheets[0]))
    t.sheets_added.append(name)
    com(lambda: setattr(ws, "Name", name))
    if color is not None:
        com(lambda: setattr(ws.Tab, "Color", color))
    return ws


# ------------------------------------------------------------------ data pull into the user's workbook
@dataclass
class Pull:
    sheet: str
    connection: str
    layout: Layout
    parsed: object
    seconds: float
    row_labels: dict          # tuple(requested members) -> tuple(labels as shown in the grid)
    col_labels: dict
    positional: bool = False  # labels not unique -> formulas reference cells by position


# Grid options forced on every data sheet: the layout is mapped by POSITION, so Smart View must not drop rows/columns
# (the user's own defaults may suppress missing/zero/no-access rows - that would silently shift every label).
DATA_SHEET_OPTIONS = {5: 0, 6: False, 7: False, 9: False, 41: False, 42: False, 43: False, 44: False,
                      13: "#NumericZero", 14: "#NoAccess", 30: True}


META_SHEET = "_svmeta"


def save_meta(t: Target, p: "Pull", extra: dict | None = None) -> None:
    """Persist the pull's layout in a hidden sheet, so answers can be (re)built after a server restart."""
    if not _sheet_exists(t, META_SHEET):
        ws = com(lambda: t.wb.Worksheets.Add(None, t.wb.Worksheets(t.wb.Worksheets.Count)))   # positional: see add_sheet
        com(lambda: setattr(ws, "Name", META_SHEET))
        com(lambda: setattr(ws, "Visible", 0))         # xlSheetHidden
    ws = com(lambda: t.wb.Worksheets(META_SHEET))
    blob = json.dumps({"connection": p.connection, "pov": p.layout.pov,
                       "rows": [[a.dimension, a.members] for a in p.layout.rows],
                       "cols": [[a.dimension, a.members] for a in p.layout.cols],
                       "warnings": [w for w in p.parsed.pov_problems if w.startswith("Smart View rejected")],
                       **(extra or {})},
                      ensure_ascii=False)
    chunks = [blob[i:i + 30000] for i in range(0, len(blob), 30000)]   # a cell holds max 32 767 characters
    r = com(lambda: ws.Cells(ws.Rows.Count, 1).End(-4162).Row) + 1
    com(lambda: setattr(ws.Cells(r, 1), "Value", "'" + p.sheet))
    for k, c in enumerate(chunks):
        com(lambda: setattr(ws.Cells(r, 2 + k), "Value", "'" + c))


def save_spec(t: Target, q: int, spec: dict, ids: dict[str, str] | None = None, owned: dict | None = None) -> None:
    """Keep the question's last answer spec (unexpanded) + pull ids + the tool-owned range of its answer sheet
    ({"rows", "cols"}) in the hidden sheet, for patch-only follow-ups."""
    if not _sheet_exists(t, META_SHEET):
        return
    ws = com(lambda: t.wb.Worksheets(META_SHEET))
    key = f"SPEC Q{q}"
    used = com(lambda: ws.UsedRange)
    vals = com(lambda: used.Value)
    vals = vals if isinstance(vals, tuple) else ((vals,),)
    top = int(com(lambda: used.Row))                   # UsedRange starts at the first used row (save_meta: row 2)
    r = next((top + i for i, row in enumerate(vals) if row and row[0] == key), None)
    if r is None:
        r = com(lambda: ws.Cells(ws.Rows.Count, 1).End(-4162).Row) + 1
    else:
        com(lambda: ws.Rows(r).ClearContents())
    blob = json.dumps({"spec": spec, "ids": ids or {}, "owned": owned}, ensure_ascii=False)
    com(lambda: setattr(ws.Cells(r, 1), "Value", "'" + key))
    for k in range(0, len(blob), 30000):
        com(lambda: setattr(ws.Cells(r, 2 + k // 30000), "Value", "'" + blob[k:k + 30000]))


def _stored_specs(t: Target) -> dict[int, dict]:
    """Every SPEC Q<n> row of the hidden sheet: {q: {"spec", "ids", "owned"}}."""
    if not _sheet_exists(t, META_SHEET):
        return {}
    vals = com(lambda: t.wb.Worksheets(META_SHEET).UsedRange.Value)
    vals = vals if isinstance(vals, tuple) else ((vals,),)
    out = {}
    for row in vals:
        m = re.match(r"^SPEC Q(\d+)$", row[0]) if row and isinstance(row[0], str) else None
        if m:
            d = json.loads("".join(str(x) for x in row[1:] if x))
            out[int(m.group(1))] = {"spec": d.get("spec") or {}, "ids": d.get("ids") or {}, "owned": d.get("owned")}
    return out


def load_answer(t: Target, q: int) -> dict | None:
    """The stored answer of question q: {"spec", "ids", "owned"} (owned None for answers before phase 5)."""
    return _stored_specs(t).get(q)


def stored_questions(t: Target) -> list[int]:
    """The questions with a stored answer spec v1 (drawn by sv_answer), ascending."""
    return sorted(q for q, a in _stored_specs(t).items() if a["spec"].get("version") == 1)


def load_pulls(t: Target, q: int) -> dict:
    """Rebuild Pull objects of question q from the hidden metadata + the live data sheets."""
    from .ops import _used_values
    if not _sheet_exists(t, META_SHEET):
        return {}
    ws = com(lambda: t.wb.Worksheets(META_SHEET))
    vals = com(lambda: ws.UsedRange.Value)
    vals = vals if isinstance(vals, tuple) else ((vals,),)
    out = {}
    for row in vals:
        sheet = row[0]
        if not isinstance(sheet, str) or not sheet.startswith(f"Q{q} ") or not _sheet_exists(t, sheet):
            continue
        meta = json.loads("".join(str(x) for x in row[1:] if x))
        lay = Layout([tuple(x) for x in meta["pov"]], [Axis(d, m) for d, m in meta["rows"]],
                     [Axis(d, m) for d, m in meta["cols"]])
        parsed = parse(lay, _used_values(com(lambda: t.wb.Worksheets(sheet))), False, False)
        parsed.pov_problems[:0] = meta.get("warnings", [])
        out[sheet] = _make_pull(sheet, meta["connection"], lay, parsed, 0.0)
    return out


def fte_rows(lay: Layout, parsed) -> list[int]:
    """Sheet rows (1-based) of FTE / man-year lines: they are styled as decimals, not NOK."""
    return [2 + lay.nc + i for i, labs in enumerate(parsed.row_labels)
            if any(isinstance(x, str) and ("STA" in x or "man-years" in x.lower()) for x in labs)]


def _make_pull(name, connection, lay, parsed, secs) -> "Pull":
    row_map = {tuple(combo): tuple(parsed.row_labels[i]) for i, combo in enumerate(lay.row_tuples)
               if i < len(parsed.row_labels)}
    col_map = {tuple(combo): tuple(str(x).strip() if x is not None else "" for x in parsed.col_labels[j])
               for j, combo in enumerate(lay.col_tuples)}
    positional = len(set(row_map.values())) < len(row_map) or len(set(col_map.values())) < len(col_map)
    return Pull(name, connection, lay, parsed, secs, row_map, col_map, positional)


def pull(h: ExcelHost, t: Target, q: int, connection: str, pov: dict, rows: list[Axis], cols: list[Axis],
         title: str | None = None) -> Pull:
    """A live Smart View grid in the user's workbook; Smart View's modal message boxes in that Excel are dismissed
    while it runs (SmartViewBoxes), so they cannot block the pull."""
    with SmartViewBoxes(t.pid) as boxes:
        return _pull(h, t, q, connection, pov, rows, cols, title, boxes)


def _pull(h: ExcelHost, t: Target, q: int, connection: str, pov: dict, rows: list[Axis], cols: list[Axis],
          title: str | None, boxes: "SmartViewBoxes") -> Pull:
    from . import styling                         # provided by the styling adapter
    from .grid import validate_axes
    from .ops import _dims_cache, _grid_labels, _used_values, canon_dim, dims_known, is_rc1000, learn_dims
    from .timing import stage

    def layout():
        p_ = {canon_dim(connection, d): m for d, m in pov.items()}
        r_ = [Axis(canon_dim(connection, a.dimension), a.members) for a in rows]
        c_ = [Axis(canon_dim(connection, a.dimension), a.members) for a in cols]
        dims = _dims_cache[connection]
        problems = validate_axes(dims, p_, r_, c_)
        if problems:
            raise InputError("Invalid pull: " + "; ".join(problems))
        return Layout(full_pov(dims, p_, r_, c_), r_, c_)

    # dimension labels come from the on-disk cache, else from this sheet's own default grid below - never from the
    # hidden server Excel (a second Excel + login, and a connect there once hung a pull for minutes)
    lay = layout() if dims_known(connection) else None
    name = unique_sheet_name(t, title or f"Q{q} {connection}")
    t.status(f"henter {connection} til arket '{name}' ...")
    with stage("add sheet"):
        com(t.wb.Activate)
        ws = add_sheet(t, name, after_q=q, color=0xA6A6A6)
        com(ws.Activate)
    with stage("connect"):
        from .dialogs import not_signed_in
        failed: list[list[str]] = []              # box texts of each failed attempt
        for attempt in range(3):                  # EPM sometimes answers a connect with its extension manifest
            n0 = len(boxes.texts)
            st = t.run("B_Connect", name, connection)
            if st == "OK":
                break
            failed.append(boxes.texts[n0:])
            if not_signed_in(failed):             # retrying cannot fix a session EPM keeps rejecting
                raise HostError(NOT_SIGNED_IN)
            if attempt == 2:
                from .host import _explain
                raise HostError(_explain("Connect", st))
            time.sleep(4 * (attempt + 1))
    # Always let Smart View open its own default grid first (1-2 s). Skipping it was tried on 2026-10-14 and was
    # followed by a 752 error and two Excel crashes (HsAddin.dll access violation) - not worth the time.
    with stage("open default grid"):
        _open_default_grid(t, name)
    if lay is None:                               # first pull of this cube ever: learn its dimensions here
        with stage("learn dimensions"):
            r = t.run("B_GetDimensions", name)
            labels = _grid_labels(com(lambda: _used_values(ws)), list(r[2])) if isinstance(r, tuple) else None
            if not labels:
                raise HostError(f"Could not read the {connection} dimensions from the default grid on '{name}'.")
            learn_dims(connection, list(r[1]), list(r[2]), labels)
            lay = layout()
    with stage("sheet options"):
        for k, v in DATA_SHEET_OPTIONS.items():
            t.run("B_SetOption", k, v, name)
    rejected: list[tuple[str, str]] = []
    for attempt in range(2):
        with stage("write layout + retrieve"):
            st, secs = _write_and_retrieve(t, ws, name, lay)
        if st != "OK" and is_rc1000(st) and not rejected:
            with stage("isolate rejected members"):
                lay, rejected = _drop_rejected(t, ws, name, lay, connection)
            if rejected:
                with stage("write layout + retrieve"):
                    st, secs = _write_and_retrieve(t, ws, name, lay)
        if st != "OK":
            raise HostError(f"Retrieve on '{name}' failed: {st}")
        with stage("read + verify"):
            vals = com(lambda: _used_values(ws))
            parsed = parse(lay, vals, suppress_missing=False, suppress_zeros=False)
            bad = _verify_labels(lay, parsed)
        if not bad or attempt or not _fix_sheet_mode(t, name):
            break                                 # labels off because the sheet opened in alias 'none': fixed, redo
    if bad:
        raise HostError("Smart View returned a different layout than requested, so the numbers cannot be mapped "
                        "safely: " + "; ".join(bad[:5]) + (" ..." if len(bad) > 5 else ""))
    if rejected:
        from .ops import _explain_reject
        parsed.pov_problems.insert(0, "Smart View rejected these members (rc=1000), so they were left out: "
                                   + "; ".join(_explain_reject(connection, d, m_) for d, m_ in rejected)
                                   + ". Pick a more specific member with sv_find_member / sv_members.")
    fte = fte_rows(lay, parsed)
    t.status(f"formaterer '{name}' ...")
    p = _make_pull(name, connection, lay, parsed, secs)
    with fast_excel(t.app):
        with stage("style data sheet"):
            styling.apply_data_sheet(ws, header_rows=1 + lay.nc, label_cols=lay.nr, fte_rows=fte)
        with stage("save meta"):
            try:
                save_meta(t, p)
            except Exception:   # noqa: BLE001 - metadata only helps after restarts
                pass
    return p


def refresh_pull(t: Target, p: Pull) -> Pull:
    """Retrieve an existing data sheet again IN PLACE (same layout): its name, formulas pointing at it and its
    formatting stay; only the numbers are refreshed. Smart View's message boxes are dismissed meanwhile."""
    with SmartViewBoxes(t.pid):
        return _refresh_pull(t, p)


def _refresh_pull(t: Target, p: Pull) -> Pull:
    from .ops import _used_values
    from .timing import stage
    ws = com(lambda: t.wb.Worksheets(p.sheet))
    t.status(f"oppdaterer '{p.sheet}' ...")
    with stage("retrieve in place"):
        t0 = time.monotonic()
        st = t.run("B_Retrieve", p.sheet)
        secs = time.monotonic() - t0
    if st != "OK":
        raise HostError(f"Retrieve on '{p.sheet}' failed: {st}")
    with stage("read + verify"):
        parsed = parse(p.layout, com(lambda: _used_values(ws)), suppress_missing=False, suppress_zeros=False)
        bad = _verify_labels(p.layout, parsed)
    if bad:
        raise HostError(f"'{p.sheet}' no longer has the expected layout: " + "; ".join(bad[:5])
                        + ". Re-run sv_answer for this question.")
    parsed.pov_problems[:0] = [w for w in p.parsed.pov_problems if w.startswith("Smart View rejected")]
    return _make_pull(p.sheet, p.connection, p.layout, parsed, secs)


def _write_and_retrieve(t: Target, ws, name: str, lay: Layout) -> tuple[str, float]:
    com(lambda: ws.Cells.ClearContents())
    m = lay.matrix()
    height, width = len(m), len(m[0])
    com(lambda: setattr(ws.Range(ws.Cells(1, 1), ws.Cells(1 + lay.nc, width)), "NumberFormat", "@"))
    com(lambda: setattr(ws.Range(ws.Cells(2 + lay.nc, 1), ws.Cells(height, lay.nr)), "NumberFormat", "@"))
    com(lambda: setattr(ws.Range(ws.Cells(1, 1), ws.Cells(height, width)), "Value", m))
    t0 = time.monotonic()
    st = t.run("B_Retrieve", name)
    return st, time.monotonic() - t0


def _drop_rejected(t: Target, ws, name: str, lay: Layout, connection: str) -> tuple[Layout, list[tuple[str, str]]]:
    """rc=1000: bisect each axis on this sheet to find the members Smart View rejects; return the layout without
    them (unchanged + [] if none could be isolated or an axis would be empty)."""
    from .ops import MAX_BISECT_RETRIEVES, is_rc1000
    axes = lay.rows + lay.cols
    budget = [MAX_BISECT_RETRIEVES]

    def ok(i, members):
        if budget[0] <= 0:
            raise HostError(f"Retrieve on '{name}' was rejected (rc=1000) and the bad member could not be isolated "
                            f"within {MAX_BISECT_RETRIEVES} test retrieves; split the pull.")
        budget[0] -= 1
        trial = [Axis(a.dimension, members if j == i else a.members[:1]) for j, a in enumerate(axes)]
        lt = Layout(lay.pov, trial[:len(lay.rows)], trial[len(lay.rows):])
        st, _ = _write_and_retrieve(t, ws, name, lt)
        if st != "OK" and not is_rc1000(st):
            raise HostError(f"Retrieve on '{name}' failed: {st}")
        return st == "OK"

    def bad(i, members):
        if len(axes[i].members) <= 1 or ok(i, members):
            return []
        if len(members) == 1:
            return members
        mid = len(members) // 2
        return bad(i, members[:mid]) + bad(i, members[mid:])

    rej = [(a.dimension, m) for i, a in enumerate(axes) for m in bad(i, list(a.members))]
    drop = set(rej)
    new = [Axis(a.dimension, [m for m in a.members if (a.dimension, m) not in drop]) for a in axes]
    if not rej or any(not a.members for a in new):
        return lay, []
    return Layout(lay.pov, new[:len(lay.rows)], new[len(lay.rows):]), rej


def _sheet_mode_bad(t: Target, name: str) -> bool:
    info = t.run("B_SheetInfo", name)
    meta = dict(zip(info[1], info[2])) if isinstance(info, tuple) and info[0] == "OK" else {}
    return "Standard" not in str(meta.get("Sheet Type", "")) or str(meta.get("Alias Table", "")).lower() == "none"


def _open_default_grid(t: Target, name: str) -> None:
    st = t.run("B_Retrieve", name)
    if st != "OK":
        raise HostError(f"Smart View could not open a grid on '{name}': {st}")
    if _sheet_mode_bad(t, name):                  # seen once: FinMax opened as plain 'Ad hoc', alias none
        t.run("B_SetAliasTable", name, "Default")
        t.run("B_Retrieve", name)


def _fix_sheet_mode(t: Target, name: str) -> bool:
    """After a label mismatch: if the sheet opened in a non-standard mode / alias 'none', switch to Default."""
    if not _sheet_mode_bad(t, name):
        return False
    t.run("B_SetAliasTable", name, "Default")
    return True


def _q(s: str) -> str:
    # the formulas compare against TRIM(sheet label), which collapses double spaces
    return " ".join(s.split()).replace('"', '""')


def lookup_formula(p: Pull, row_members: list[str], col_members: list[str]) -> str:
    """Live label-keyed lookup into a pull sheet (skill tn-formelstil), robust to rows moving on refresh:
        =XLOOKUP("C30 REVENUES (ac)",'Q1 FinMax'!$A:$A,XLOOKUP("Sep 26",'Q1 FinMax'!$A$2:$M$2,'Q1 FinMax'!$A:$M))
    Rows are whole columns (never a guessed $A$5:$A$400); the header row and the block share the exact data width,
    so the arrays always match. Several column dimensions -> XLOOKUP(1,(row2="a")*(row3="b"),...). Several ROW
    dimensions need bounded criteria (whole-column array tests take seconds per sheet) -> INDEX/XMATCH over trim
    refs ($A:.$A), whose index is independent of the block size. A missing label shows #N/A, never a wrong number."""
    rlab = p.row_labels.get(tuple(row_members))
    clab = p.col_labels.get(tuple(col_members))
    if rlab is None or clab is None:
        raise InputError(f"No cell for rows {row_members} / columns {col_members} in '{p.sheet}'. "
                         f"Available rows: {list(p.row_labels)[:6]} columns: {list(p.col_labels)[:6]}")
    s = "'" + p.sheet.replace("'", "''") + "'!"
    nr, nc = p.layout.nr, p.layout.nc
    from openpyxl.utils import get_column_letter as L
    if p.positional:                                   # duplicate labels: a label lookup could hit the wrong row
        i = p.layout.row_tuples.index(tuple(row_members))
        j = p.layout.col_tuples.index(tuple(col_members))
        return f"={s}${L(nr + 1 + j)}${2 + nc + i}"
    W = L(nr + len(p.layout.col_tuples))                 # last data column (suppression is off: width is fixed)
    lab = lambda x: '"' + str(x).strip().replace('"', '""') + '"'   # exact label (no TRIM in these formulas)
    if nc == 1:
        col_lookup, col_arr = lab(clab[0]), f"{s}$A$2:${W}$2"
    else:
        col_lookup, col_arr = "1", "*".join(f"({s}$A${2 + d}:${W}${2 + d}={lab(clab[d])})" for d in range(nc))
    if nr == 1:
        return f"=XLOOKUP({lab(rlab[0])},{s}$A:$A,XLOOKUP({col_lookup},{col_arr},{s}$A:${W}))"
    row_match = "XMATCH(1," + "*".join(f"({s}${L(1 + d)}:.${L(1 + d)}={lab(rlab[d])})" for d in range(nr)) + ")"
    col_match = (f"XMATCH({col_lookup},{col_arr})" if nc == 1 else f"XMATCH(1,{col_arr})")
    return f"=INDEX({s}$A:${W},{row_match},{col_match})"


def _kilder_name(t: Target) -> str:
    """'Kilder', unless the user already has their own sheet with that name - then 'Kilder SV'."""
    if _sheet_exists(t, "Kilder"):
        a1 = com(lambda: t.wb.Worksheets("Kilder").Range("A1").Value)
        used = com(lambda: t.wb.Worksheets("Kilder").UsedRange.Address)
        if a1 != "Hentet" and used not in ("$A$1",):
            return "Kilder SV"
    return "Kilder"


def log_source(t: Target, q: int, question: str, p: Pull, notes: str = "") -> None:
    """Append a row to the sources sheet (created if missing, kept last)."""
    with fast_excel(t.app):
        _log_source(t, q, question, p, notes)


def _log_source(t: Target, q: int, question: str, p: Pull, notes: str) -> None:
    kname = _kilder_name(t)
    if not _sheet_exists(t, kname):
        ws = com(lambda: t.wb.Worksheets.Add(None, t.wb.Worksheets(t.wb.Worksheets.Count)))   # positional: see add_sheet
        com(lambda: setattr(ws, "Name", kname))
    ws = com(lambda: t.wb.Worksheets(kname))
    if com(lambda: ws.Range("A1").Value) != "Hentet":
        com(lambda: setattr(ws.Range("A1:H1"), "Value", (("Hentet", "Spørsmål nr", "Spørsmål", "Kube", "POV",
                                                          "Rader × kolonner", "Merknader", "Ark"),)))
        com(lambda: setattr(ws.Range("A1:H1").Font, "Bold", True))
        com(lambda: setattr(ws.Range("A1:H1").Font, "Color", 0x64381F))
    r = com(lambda: ws.Cells(ws.Rows.Count, 1).End(-4162).Row) + 1
    pov = " · ".join(f"{m}" for d, m in p.layout.pov if m != d)
    axes = " × ".join(f"{a.dimension}[{len(a.members)}]" for a in p.layout.rows + p.layout.cols)
    row = ((dt.datetime.now().strftime("%Y-%m-%d %H:%M"), f"Q{q}", question, p.connection, pov, axes, notes),)
    com(lambda: setattr(ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)), "Value", row))
    # positional (Anchor, Address, SubAddress, ScreenTip, TextToDisplay): keyword TextToDisplay= was ignored
    com(lambda: ws.Hyperlinks.Add(ws.Cells(r, 8), "", f"'{p.sheet}'!A1", p.sheet, p.sheet))
    com(lambda: ws.Range("A1:H1").EntireColumn.AutoFit())
    if com(lambda: ws.Columns("C").ColumnWidth) > 60:
        com(lambda: setattr(ws.Columns("C"), "ColumnWidth", 60))


def wait_calculated(app, limit: float = 60.0) -> None:
    """Wait until Excel has finished recalculating (CalculationState == xlDone)."""
    t0 = time.monotonic()
    while time.monotonic() - t0 < limit:
        try:
            if app.CalculationState == 0:
                return
        except pywintypes.com_error as e:
            if not _busy(e):
                raise
        time.sleep(0.3)
