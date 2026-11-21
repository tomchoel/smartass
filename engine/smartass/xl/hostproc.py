"""Our own Excel: processes (verified by PID + creation time + '-Embedding'), the STA COM thread, HiddenExcel."""
from __future__ import annotations

import concurrent.futures as cf
import contextlib
import json
import logging
import os
import queue
import threading
import time
import traceback
from typing import Any, Callable

import psutil
import pythoncom
import pywintypes
import win32api
import win32con

from ..core import config

log = logging.getLogger("smartview.host")
RETRY_HRESULTS = {-2147418111, -2147417846}        # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER
VBA_E_IGNORE = -2146777998                         # 0x800AC472: Excel busy (recalculating), inside DISP_E_EXCEPTION


def is_busy(e: BaseException) -> bool:
    """Excel answered 'call rejected / retry later' (busy opening, recalculating, a cell in edit mode)."""
    hr, info = getattr(e, "hresult", None), getattr(e, "excepinfo", None)
    return hr in RETRY_HRESULTS or VBA_E_IGNORE in (hr, info[5] if info and len(info) > 5 else None)


def retry_busy(call: Callable[[], Any], limit_s: float = 120, alive: Callable[[], Any] | None = None,
               rejected: Callable[[float], Any] | None = None):
    """Run any COM call, retrying while Excel is busy (is_busy) for up to limit_s; then the COM error is raised.
    pywin32 has no IMessageFilter, so this is the retry loop instead. alive: called before every try (jobs.com: the
    job's abort, which raises). rejected: told at the end the seconds Excel rejected the call (the rejected tries and
    the pauses after them; not the try that answered, however long it took)."""
    delay, t0, waited = 0.1, time.monotonic(), 0.0
    try:
        while True:
            if alive is not None:
                alive()
            t_try = time.monotonic()
            try:
                return call()
            except pywintypes.com_error as e:
                if not is_busy(e):
                    raise
                if time.monotonic() - t0 >= limit_s:
                    waited += time.monotonic() - t_try
                    raise
                pythoncom.PumpWaitingMessages()
                time.sleep(delay)
                delay = min(delay * 2, 2.0)
                waited += time.monotonic() - t_try
    finally:
        if rejected is not None and waited:
            rejected(waited)


def _excel_pids() -> set[int]:
    return {p.pid for p in psutil.process_iter(["name"]) if (p.info["name"] or "").lower() == "excel.exe"}


def _is_ours(pid: int, ctime: float) -> bool:
    try:
        p = psutil.Process(pid)
        return abs(p.create_time() - ctime) < 1.0 and "-embedding" in " ".join(p.cmdline()).lower()
    except psutil.Error:
        return False


def windows_of(pid: int) -> list[str]:
    """Visible top-level windows of a process ('title [class]') - tells us which dialog blocked Excel."""
    from .dialogs import windows
    return [f"{title!r} [{cls}]" for _h, cls, title in windows(pid=pid)]


def kill_ours(pid: int, ctime: float) -> bool:
    if not _is_ours(pid, ctime):
        return False
    try:
        h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(h, 1)
        win32api.CloseHandle(h)
        return True
    except pywintypes.error:
        return False


class _PidRegistry:
    """Persist the Excel PIDs we started, with the owning server process, so a restarted server cleans up only
    orphans whose owner is gone - never the live Excel of another running server (e.g. a test harness)."""

    LOCK_WAIT_S = 10.0
    _tlock = threading.RLock()                     # threads of this process; the byte lock covers other processes

    def __init__(self, path=config.PID_FILE):
        self.path = path

    @contextlib.contextmanager
    def _locked(self):
        """Exclusive lock for a read-modify-write of the file: a thread lock plus an msvcrt byte lock on
        <file>.lock (Windows frees it when the process dies). Not re-entrant across processes - only the public
        methods take it."""
        import msvcrt
        with self._tlock:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.path.with_name(self.path.name + ".lock"), "a+b") as f:
                t0 = time.monotonic()
                while True:
                    try:
                        f.seek(0)
                        msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                        break
                    except OSError:
                        if time.monotonic() - t0 > self.LOCK_WAIT_S:
                            break                  # a stuck holder must not block Excel start-up: write anyway
                        time.sleep(0.02)
                try:
                    yield
                finally:
                    try:
                        f.seek(0)
                        msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                    except OSError:
                        pass

    def load(self) -> list[dict]:
        try:
            return json.loads(self.path.read_text())
        except (OSError, ValueError):
            return []

    def _write(self, items: list[dict]) -> None:
        """Atomic: write <file>.<pid>.tmp, then os.replace (retried while a reader briefly holds the target)."""
        tmp = self.path.with_name(f"{self.path.name}.{os.getpid()}.tmp")
        tmp.write_text(json.dumps(items))
        for attempt in range(200):
            try:
                os.replace(tmp, self.path)
                return
            except PermissionError:
                if attempt == 199:
                    tmp.unlink(missing_ok=True)
                    raise
                time.sleep(0.005)

    def save(self, items: list[dict]) -> None:
        with self._locked():
            self._write(items)

    @staticmethod
    def _owner_alive(item: dict) -> bool:
        owner, octime = item.get("owner"), item.get("owner_ctime")
        if not owner:
            return False
        try:
            return abs(psutil.Process(owner).create_time() - octime) < 1.0
        except psutil.Error:
            return False

    def add(self, pid: int, ctime: float) -> None:
        me = psutil.Process()
        with self._locked():
            keep = [i for i in self.load() if psutil.pid_exists(i["pid"])]
            self._write(keep + [{"pid": pid, "ctime": ctime, "owner": me.pid, "owner_ctime": me.create_time()}])

    def remove(self, pid: int) -> None:
        with self._locked():
            self._write([i for i in self.load() if i["pid"] != pid])

    def sweep(self) -> list[int]:
        killed, keep = [], []
        with self._locked():
            for i in self.load():
                if self._owner_alive(i):
                    keep.append(i)                 # belongs to another live server process
                elif kill_ours(i["pid"], i["ctime"]):
                    killed.append(i["pid"])
            self._write(keep)
        return killed


class ComThread(threading.Thread):
    """Single STA thread that executes callables in order and pumps messages while idle. A job's error leaves the
    thread without the locals of the frames it passed through (_release_frames): the job's COM objects are released
    here, on the thread whose apartment made them."""

    def __init__(self, name: str = "excel-com"):
        super().__init__(daemon=True, name=name)
        self.q: queue.Queue = queue.Queue()
        self.busy_since: float | None = None
        self.job_no = 0                           # counts started jobs: tells a reset that the blocked job has ended
        self.waiting_reset = False                # the current job sits in ExcelHost._wait_reset (not blocked)

    def run(self):
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        try:
            while True:
                try:
                    fn, fut = self.q.get(timeout=0.05)
                except queue.Empty:
                    pythoncom.PumpWaitingMessages()
                    continue
                if fn is None:
                    break
                if not fut.set_running_or_notify_cancel():
                    continue
                self.job_no += 1
                self.busy_since = time.monotonic()
                try:
                    fut.set_result(fn())
                except BaseException as e:   # noqa: BLE001 - forwarded to the awaiting coroutine
                    _release_frames(e)
                    fut.set_exception(e)
                finally:
                    self.busy_since = None
        finally:
            pythoncom.CoUninitialize()

    def submit(self, fn: Callable[[], Any]) -> cf.Future:
        fut: cf.Future = cf.Future()
        self.q.put((fn, fut))
        return fut


def _release_frames(e: BaseException) -> None:
    """The locals of every finished frame e's traceback holds - and those of its cause and context - released now, on
    this thread. An error handed to the event loop otherwise carries the failed job's COM objects (its Target's app
    and workbook, a batch's ranges) to that thread, which releases them when it drops the error: a proxy released on
    another thread while the apartment that made it is alive calls that Excel from there and waits for it (Task D,
    21.11.2026: for ever, when that Excel has stopped responding). The line numbers stay: the error logs as before."""
    todo, seen = [e], set()
    while todo:
        x = todo.pop()
        if x is None or id(x) in seen:
            continue
        seen.add(id(x))
        traceback.clear_frames(x.__traceback__)
        todo += [x.__cause__, x.__context__]


def on_host_thread() -> bool:
    """True on an ExcelHost's own COM thread (named 'excel-com'): work there must not wait on threads that call
    on_host, or it deadlocks."""
    t = threading.current_thread()
    return isinstance(t, ComThread) and t.name == "excel-com"


class HiddenExcel:
    """ExcelHost's hidden Excel on the host COM thread: start check, VBA bridge, Smart View options, sheets."""

    def on_host(self, fn: Callable[[], Any], timeout: float | None = None):
        """Run fn on the host COM thread (where the hidden Excel's COM objects live), starting that Excel if needed.
        Called from a user-Excel job thread, it waits for the result; on the host thread it just runs."""
        if threading.current_thread() is self.thread:
            self._ensure_started()
            return fn()

        def job():
            self._wait_reset()
            self._ensure_started()
            return fn()
        return self.thread.submit(job).result(timeout or config.CALL_TIMEOUT_S)

    def _ensure_started(self):
        if self.app is not None:
            alive = self.pid is not None and self.ctime is not None and _is_ours(self.pid, self.ctime)
            try:
                if alive:
                    _ = self.app.Hwnd
                    return
            except Exception:   # noqa: BLE001 - a dead proxy raises com_error or AttributeError
                pass
            self._mark_dead("Excel was gone before the call (killed or crashed); restarting")
            self._drop_refs()
        self._start()

    def _load_bridge(self, wb):
        for path in config.BAS_FILES:
            mod = wb.VBProject.VBComponents.Add(1)
            mod.Name = "SmartViewVBA" if path.stem == "smartview_body" else path.stem
            mod.CodeModule.AddFromString(path.read_text(encoding="cp1252"))

    def _apply_global_options(self):
        """The user's own values are recorded ONCE in a file before we change them, so a crash (no clean shutdown)
        never turns our settings into the 'originals' of the next run. While the old svkit engine may still have its
        own values applied, its file holds the real originals: copy them (the legacy file stays for that engine)."""
        orig_file = config.SVKIT / "smartview_options_original.json"
        legacy = config.LEGACY_HOME / "smartview_options_original.json" if config.LEGACY_HOME else None
        saved, src = None, None
        for f in [p for p in (orig_file, legacy) if p]:
            try:
                saved, src = {int(k): v for k, v in json.loads(f.read_text()).items()}, f
                break
            except (OSError, ValueError):
                pass
        if saved is None:
            saved = {}
            for k in config.GLOBAL_OPTIONS:
                got = self._run("B_GetOption", k, "")
                if got and got[0] == "OK":
                    saved[k] = got[1]
        if src == orig_file:
            # A clean shutdown deletes our file, so it is left over from a crashed run: put the user's originals
            # back first (our values were never restored), then apply ours below as on any start.
            log.info("smartview options: restoring the originals left by a crashed run")
            for k, v in saved.items():
                try:
                    self._run("B_SetOption", k, v, "")
                except pywintypes.com_error as e:
                    log.info("restoring option %s failed: %s", k, str(e)[:200])
        else:
            orig_file.parent.mkdir(parents=True, exist_ok=True)
            orig_file.write_text(json.dumps({str(k): v for k, v in saved.items()}))
        self.saved_options = saved
        for k, v in config.GLOBAL_OPTIONS.items():
            self._run("B_SetOption", k, v, "")

    def _restore_global_options(self):
        ok = True
        for k, v in self.saved_options.items():
            try:
                self._run("B_SetOption", k, v, "")
            except pywintypes.com_error:
                ok = False
        if ok:
            try:
                (config.SVKIT / "smartview_options_original.json").unlink()
            except OSError:
                pass

    def _run(self, fn: str, *args):
        """Application.Run on the bridge, with retry on 'call rejected'."""
        return self.com(lambda: self.app.Run(f"'{self.wb.Name}'!modBridge.{fn}", *args), limit_s=60)

    com = staticmethod(retry_busy)

    def sheet(self, name: str, fresh: bool = False):
        ws = None
        for s in self.wb.Worksheets:
            if s.Name == name:
                ws = s
                break
        if ws is not None and fresh:
            self.app.DisplayAlerts = False
            ws.Delete()
            self.connected.pop(name, None)
            ws = None
        if ws is None:
            ws = self.wb.Worksheets.Add()
            ws.Name = name
        return ws

    def connect_sheet(self, name: str, connection: str):
        if self.connected.get(name) == connection:
            return
        self.wb.Worksheets(name).Activate()
        self.run("B_Connect", name, connection)
        self.connected[name] = connection

    def status(self) -> dict:
        return {
            "excel_running": self.app is not None,
            "pid": self.pid,
            "visible": self.visible,
            "started_at": self.started_at,
            "connected_sheets": dict(self.connected),
            "restarts": self.restarts,
            "smartview_version": self.sv_version,
            "last_error": self.last_error,
        }

    def _drop_refs(self):
        self.app = self.wb = None
        self.connected.clear()

    def _mark_dead(self, why: str):
        self.last_error = why
        self.restarts += 1
        log.warning("Excel marked dead: %s", why)
