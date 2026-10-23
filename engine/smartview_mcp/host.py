"""The Excel host: one dedicated, hidden Excel driven from one STA thread.

All COM happens on ComThread. Calls carry a deadline; when one passes, the watchdog terminates
*our own* Excel process (verified by PID + creation time + "-Embedding"), which makes the blocked
COM call return with an RPC error. The next call starts a fresh Excel. Other Excel processes
(the user's, excel-mcp's) are never touched.
"""
from __future__ import annotations

import concurrent.futures as cf
import json
import logging
import queue
import re
import threading
import time
from typing import Any, Callable

import psutil
import pythoncom
import pywintypes
import win32api
import win32con
import win32com.client as wc
import win32process

from . import config, dialogs, sso

log = logging.getLogger("smartview.host")

RETRY_HRESULTS = {-2147418111, -2147417846}        # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER
DEAD_HRESULTS = {-2147023174, -2147023170, -2147417848}   # RPC unavailable, call failed, object disconnected


class HostError(Exception):
    """Raised for failures the model should see (message is meant to be actionable)."""


class HostTimeout(HostError):
    pass


class SignInBlocked(HostError):
    """Smart View's sign-in is stuck in our hidden Excel (an Internet Explorer Login window, a WebView2-fallback box,
    or a Login window the SSO watcher could not finish). ExcelHost.call kills that Excel and retries once."""

    def __init__(self, msg: str, dialog: dialogs.Dialog, pid: int | None = None):
        super().__init__(msg)
        self.dialog, self.pid = dialog, pid


def signin_message(d: dialogs.Dialog, again: bool, holders: list[str]) -> str:
    if d.kind == "webview2_fallback" or d.ie:
        cause = ("Smart View showed its sign-in window (WebView2 failed and it fell back to Internet Explorer, whose "
                 "Oracle login page doesn't work)")
    else:
        cause = (f"Smart View's sign-in window stayed open for more than {config.LOGIN_GRACE_S:.0f}s (the automatic "
                 "SSO click did not complete)")
    hold = (f" Smart View's WebView2 folder is in use by {', '.join(holders)}; closing that app (and its Smart View) "
            "may let WebView2 start again." if holders else "")
    if not again:
        return f"{cause}. Dialog: {dialogs.describe(d)}. The engine restarted its hidden Excel; retry.{hold}"
    return (f"{cause}. Dialog: {dialogs.describe(d)}. The engine restarted Excel once; it happened again.{hold} "
            "Next step: sign in to Smart View once in your own Excel (Smart View ribbon → Panel → connect), then "
            "retry; if it repeats, run `smartass doctor --bundle`.")


class InputError(ValueError):
    """Bad input from the caller (spec, address, name, path) - not a host failure, so no restart advice."""


def rewrap(e: Exception, msg: str) -> Exception:
    """The same kind of error with more context: input errors stay input errors."""
    return InputError(msg) if isinstance(e, InputError) else HostError(msg)


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
    import win32gui
    found: list[str] = []

    def cb(hwnd, _):
        try:
            if win32gui.IsWindowVisible(hwnd) and win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
                found.append(f"{win32gui.GetWindowText(hwnd)!r} [{win32gui.GetClassName(hwnd)}]")
        except pywintypes.error:
            pass
        return True
    try:
        win32gui.EnumWindows(cb, None)
    except pywintypes.error:
        pass
    return found


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

    def __init__(self, path=config.PID_FILE):
        self.path = path

    def load(self) -> list[dict]:
        try:
            return json.loads(self.path.read_text())
        except (OSError, ValueError):
            return []

    def save(self, items: list[dict]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(items))

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
        keep = [i for i in self.load() if psutil.pid_exists(i["pid"])]
        self.save(keep + [{"pid": pid, "ctime": ctime, "owner": me.pid, "owner_ctime": me.create_time()}])

    def remove(self, pid: int) -> None:
        self.save([i for i in self.load() if i["pid"] != pid])

    def sweep(self) -> list[int]:
        killed, keep = [], []
        for i in self.load():
            if self._owner_alive(i):
                keep.append(i)                     # belongs to another live server process
            elif kill_ours(i["pid"], i["ctime"]):
                killed.append(i["pid"])
        self.save(keep)
        return killed


class ComThread(threading.Thread):
    """Single STA thread that executes callables in order and pumps messages while idle."""

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
                    fut.set_exception(e)
                finally:
                    self.busy_since = None
        finally:
            pythoncom.CoUninitialize()

    def submit(self, fn: Callable[[], Any]) -> cf.Future:
        fut: cf.Future = cf.Future()
        self.q.put((fn, fut))
        return fut


class ExcelHost:
    def __init__(self, visible: bool = False):
        self.visible = visible
        self.thread = ComThread()
        self.thread.start()
        self.registry = _PidRegistry()
        self.app = None
        self.wb = None
        self.pid: int | None = None
        self.ctime: float | None = None
        self.started_at: float | None = None
        self.saved_options: dict[int, Any] = {}
        self.connected: dict[str, str] = {}          # sheet -> connection
        self.restarts = 0
        self.last_error: str | None = None
        self._known_before: set[int] = _excel_pids()
        self._start_wall = time.time()
        self.sv_version = sso.smartview_version()
        self.detect: Callable[[int], list[dialogs.Dialog]] = dialogs.scan
        self._resetting = threading.Event()       # set while _hard_reset runs: queued jobs wait, scans pause
        self._reset_lock = threading.Lock()
        self.reset_gen = 0
        swept = self.registry.sweep()
        if swept:
            log.warning("killed own orphaned Excel from a previous run: %s", swept)

    # ---------------- async API ----------------
    async def call(self, fn: Callable[["ExcelHost"], Any], timeout: float | None = None, what: str = "call",
                   user_excel: bool = False):
        """Run fn on the COM thread. user_excel=True: fn also drives the USER's Excel, which is never killed - on a
        timeout the blocked COM thread is abandoned instead (and our own hidden Excel restarted with a fresh thread).
        A stuck Smart View sign-in in our hidden Excel (SignInBlocked) kills that Excel; the call is retried once."""
        import asyncio
        for attempt in (1, 2):
            try:
                return await self._call_once(fn, timeout, what, user_excel)
            except SignInBlocked as e:
                d, pid = e.dialog, e.pid
                await asyncio.to_thread(self._hard_reset, f"{what}: Smart View sign-in stuck at {dialogs.describe(d)}",
                                        pid)
                if attempt == 2 or user_excel:
                    holders = await asyncio.to_thread(dialogs.webview2_holders, exclude_parent=pid)
                    raise SignInBlocked(signin_message(d, attempt == 2, holders), d) from None
                log.warning("Smart View sign-in stuck during %s (%s); restarted its Excel, retrying once",
                            what, dialogs.describe(d))

    async def _call_once(self, fn: Callable[["ExcelHost"], Any], timeout: float | None, what: str, user_excel: bool):
        import asyncio
        from . import userexcel
        timeout = timeout or (config.USER_CALL_TIMEOUT_S if user_excel else config.CALL_TIMEOUT_S)
        started = threading.Event()
        token = userexcel.JobToken()

        def job():
            started.set()
            userexcel.bind_job(token)             # an abandoned job stops at its next COM call (token check)
            if not user_excel:                    # jobs in the user's Excel start the hidden one only on demand
                self._wait_reset()
                self._ensure_started()            # (h.on_host): it costs ~10 s + an SSO login per server start
                return fn(self)
            try:
                return fn(self)
            finally:
                userexcel.release_turns()         # the per-Excel turns taken while resolving the target

        if user_excel:
            # its own STA thread per job: parallel tool calls (e.g. subagents of one session) run at the same time
            # in different Excels; work in ONE Excel still takes turns (userexcel.Turn)
            worker = ComThread(name=f"user-job-{what}")
            worker.start()
            fut = worker.submit(job)
            worker.q.put((None, None))            # the thread ends after this job
        else:
            fut = self.thread.submit(job)
        wrapped = asyncio.wrap_future(fut)
        wrapped.add_done_callback(lambda f: f.cancelled() or f.exception())   # an abandoned job's error is expected
        q0 = time.monotonic()                     # the deadline starts when the job STARTS, not while it queues
        while not started.is_set() and not fut.done():
            if time.monotonic() - q0 > config.QUEUE_TIMEOUT_S:
                fut.cancel()
                raise HostError(f"Smart View is still busy with an earlier request after {config.QUEUE_TIMEOUT_S:.0f}s "
                                f"(e.g. an atlas build); {what} was not started. Try again shortly.")
            await asyncio.sleep(0.1)
        try:
            return await self._await_job(wrapped, timeout, watch_dialogs=not user_excel)
        except asyncio.TimeoutError:
            if user_excel:
                token.aborted = True              # only this job: it rolls back when that Excel frees up
                self.last_error = f"{what} exceeded {timeout:.0f}s in the user's Excel"
                raise HostTimeout(
                    f"Your Excel did not respond within {timeout:.0f}s during {what} (a dialog, edit mode, or a large "
                    "refresh). Nothing was closed. Check Excel for an open dialog, press Esc/Enter, and ask again.") from None
            pid = self.pid
            dlg = await asyncio.to_thread(self._open_dialog_note)
            await asyncio.to_thread(self._hard_reset, f"{what} exceeded {timeout:.0f}s", pid)
            raise HostTimeout(
                f"Smart View did not respond within {timeout:.0f}s during {what}{dlg}; its Excel (pid {pid}) was "
                "terminated and will restart on the next call (SSO is clicked automatically). "
                f"Retry, or use a smaller grid. Diagnosis: {self.last_error}") from None
        except pywintypes.com_error as e:
            if e.hresult in DEAD_HRESULTS:
                self._mark_dead(f"Excel disconnected: {e}")
                raise HostError("The Smart View Excel process died during the call; it will restart on the next "
                                "call. Retry.") from None
            log.warning("COM error during %s", what, exc_info=True)
            inner = e.excepinfo[5] if e.excepinfo and len(e.excepinfo) > 5 else None
            if e.hresult in RETRY_HRESULTS or -2146777998 in (e.hresult, inner):    # 0x800AC472 VBA_E_IGNORE
                dlg = "" if user_excel else await asyncio.to_thread(self._open_dialog_note)
                raise HostError(f"Excel was busy during {what}{dlg or ' - probably a cell is being edited or a dialog is open'}"
                                ". Press Enter or Esc in Excel and ask again.") from None
            raise HostError(f"COM error during {what}: {e}") from None
        except (HostError, InputError):
            raise
        except Exception as e:   # noqa: BLE001 - always give the model a readable reason
            log.exception("unexpected error during %s", what)
            raise HostError(f"Unexpected {type(e).__name__} during {what}: {e}. Retry once; if it repeats, run "
                            "sv_restart.") from None

    async def _await_job(self, wrapped, timeout: float, watch_dialogs: bool):
        """Wait for the job; from DIALOG_CHECK_AFTER_S on, every DIALOG_POLL_S, look for a stuck Smart View sign-in in
        our hidden Excel (the COM thread is blocked by its modal window, so only this side can see it)."""
        import asyncio
        t0, first_seen = time.monotonic(), {}
        while True:
            elapsed = time.monotonic() - t0
            if elapsed >= timeout:
                raise asyncio.TimeoutError
            nxt = max(config.DIALOG_CHECK_AFTER_S, elapsed + config.DIALOG_POLL_S)
            done, _ = await asyncio.wait({wrapped}, timeout=min(nxt, timeout) - elapsed)
            if done:
                return wrapped.result()
            pid = self.pid                        # read at scan time: a reset may have replaced the Excel
            if watch_dialogs and pid and not self._resetting.is_set() \
                    and time.monotonic() - t0 >= config.DIALOG_CHECK_AFTER_S:
                found = await asyncio.to_thread(self._dialogs, pid)
                d = dialogs.blocking(found, first_seen, time.monotonic(), config.LOGIN_GRACE_S)
                if d and pid == self.pid:
                    raise SignInBlocked(f"Smart View sign-in stuck: {dialogs.describe(d)}", d, pid)

    def _dialogs(self, pid: int) -> list[dialogs.Dialog]:
        try:
            return self.detect(pid)
        except Exception:   # noqa: BLE001 - diagnosis only; never mask the real outcome
            log.warning("dialog scan of pid %s failed", pid, exc_info=True)
            return []

    def _open_dialog_note(self) -> str:
        """' - a dialog is open in the hidden Excel: 'title: text'' when one is, else ''."""
        found = self._dialogs(self.pid) if self.pid else []
        return f" - a dialog is open in the hidden Excel: {dialogs.describe(found[0])}" if found else ""

    def check_signin(self) -> None:
        """On the COM thread, after a transient Smart View error: raise SignInBlocked if a Login window or a
        WebView2-fallback box is open in our Excel (retrying in that Excel is pointless; ExcelHost.call restarts it)."""
        if self.pid is None:
            return
        d = dialogs.blocking(self._dialogs(self.pid), {}, 0.0, 0.0)
        if d:
            raise SignInBlocked(f"Smart View sign-in stuck: {dialogs.describe(d)}", d, self.pid)

    def _wait_reset(self) -> None:
        """A job queued behind a reset waits for it, so it never starts an Excel the reset then forgets."""
        deadline = time.monotonic() + config.RESET_WAIT_S + 5
        me = threading.current_thread()
        mark = isinstance(me, ComThread)
        if mark:
            me.waiting_reset = True
        try:
            while self._resetting.is_set() and time.monotonic() < deadline:
                time.sleep(0.05)
        finally:
            if mark:
                me.waiting_reset = False

    # ---------------- worker-thread side ----------------
    def ensure(self) -> "ExcelHost":
        """Start the hidden server Excel if it is not running. Only valid ON the host COM thread; user-Excel jobs run
        on their own threads and must use on_host() for anything that touches the hidden Excel."""
        self._ensure_started()
        return self

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

    def _start(self):
        t0 = time.monotonic()
        before = _excel_pids()
        self._known_before = before
        self._start_wall = time.time()
        sso.ensure_global_watcher()            # .NET UIA watcher clicks the Login dialog of every Excel
        app = wc.DispatchEx("Excel.Application")
        _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
        if pid in before:
            raise HostError("Refusing to use an Excel that already existed (DispatchEx attached to it).")
        self.app, self.pid = app, pid
        self.ctime = psutil.Process(pid).create_time()
        self.registry.add(pid, self.ctime)
        app.Visible = self.visible
        # Files the user double-clicks must open in THEIR Excel, not in this hidden one (verified 2026-10-13 that the
        # setting is per instance and does not persist into the user's Excel options).
        app.IgnoreRemoteRequests = True
        app.DisplayAlerts = False
        app.AutomationSecurity = 1
        app.AskToUpdateLinks = False
        self.wb = app.Workbooks.Add()
        self._load_bridge(self.wb)
        if self._run("B_Ping") != "pong":
            raise HostError("VBA bridge did not load.")
        self._apply_global_options()
        self.started_at = time.time()
        self.connected.clear()
        log.info("Excel started pid=%s in %.1fs", pid, time.monotonic() - t0)

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
        if src != orig_file:
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

    @staticmethod
    def com(call: Callable[[], Any], limit_s: float = 120):
        """Run any COM call, retrying while Excel answers 'call rejected / retry later' (busy opening,
        recalculating). pywin32 has no IMessageFilter, so this is the retry loop instead."""
        delay, t0 = 0.1, time.monotonic()
        while True:
            try:
                return call()
            except pywintypes.com_error as e:
                inner = e.excepinfo[5] if e.excepinfo and len(e.excepinfo) > 5 else None
                busy = e.hresult in RETRY_HRESULTS or -2146777998 in (e.hresult, inner)   # 0x800AC472 VBA_E_IGNORE
                if busy and time.monotonic() - t0 < limit_s:
                    pythoncom.PumpWaitingMessages()
                    time.sleep(delay)
                    delay = min(delay * 2, 2.0)
                    continue
                raise

    def run(self, fn: str, *args):
        """Run a bridge function and raise HostError on an ERR| status string."""
        r = self._run(fn, *args)
        status = r[0] if isinstance(r, tuple) else r
        if isinstance(status, str) and status.startswith("ERR|"):
            raise HostError(_explain(fn, status))
        return r

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

    def shutdown(self):
        def job():
            if self.app is None:
                return
            self._restore_global_options()
            try:
                self._run("B_DisconnectAll")
                self.wb.Close(False)
                self.app.Quit()
            except pywintypes.com_error:
                pass
            pid, ctime = self.pid, self.ctime
            self._drop_refs()
            time.sleep(1.5)
            if pid and kill_ours(pid, ctime):
                log.info("force-closed own Excel pid=%s after Quit", pid)
            if pid:
                self.registry.remove(pid)
        try:
            self.thread.submit(job).result(timeout=30)
        except Exception:   # noqa: BLE001 - best effort on shutdown
            if self.pid and self.ctime:
                kill_ours(self.pid, self.ctime)
        self.thread.q.put((None, None))

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

    # ---------------- recovery ----------------
    def _drop_refs(self):
        self.app = self.wb = None
        self.connected.clear()

    def _mark_dead(self, why: str):
        self.last_error = why
        self.restarts += 1
        log.warning("Excel marked dead: %s", why)

    def _abandon_thread(self, why: str):
        """The COM thread is stuck inside the USER's Excel: leave it, kill only our own hidden Excel (its proxies
        belong to the stuck apartment), and start a fresh COM thread."""
        from . import userexcel
        userexcel.abort_running_jobs()            # the stuck job stops (and rolls back) when the user's Excel frees up
        self._mark_dead(why)
        pid, ctime = self.pid, self.ctime
        if pid and ctime and kill_ours(pid, ctime):
            self.registry.remove(pid)
        old = self.thread
        old.q.put((None, None))                   # sentinel: the old thread exits once its blocked call returns
        self.thread = ComThread()
        self.thread.start()
        self.app = self.wb = None
        self.connected.clear()
        self.pid = self.ctime = None

    def _hard_reset(self, why: str, expected_pid: int | None = None):
        """Called from the event-loop thread while the COM thread is blocked. expected_pid: the Excel the caller saw
        stuck - if another reset already replaced it, there is nothing to do. One reset at a time; jobs queued behind
        it wait (_wait_reset), so the Excel they start is never cleared or orphaned by this reset."""
        with self._reset_lock:
            if expected_pid is not None and self.pid != expected_pid:
                log.info("reset for Excel pid=%s skipped: already replaced by pid=%s", expected_pid, self.pid)
                return
            thread, job_no = self.thread, self.thread.job_no    # before the flag: a job starting after it is not ours
            self._resetting.set()
            self.reset_gen += 1
            try:
                self._do_hard_reset(why, thread, job_no)
            finally:
                self._resetting.clear()

    def _do_hard_reset(self, why: str, thread: "ComThread", job_no: int):
        pid0 = pid = self.pid
        ctime = self.ctime
        if pid is None:
            # Still inside DispatchEx: only if exactly ONE new Excel appeared within 5 s of our start is it ours.
            # Anything else (excel-mcp, the user) may have started one in the meantime - leave it alone.
            cands = [p for p in psutil.process_iter(["name", "create_time"])
                     if (p.info["name"] or "").lower() == "excel.exe" and p.pid not in self._known_before
                     and self._start_wall - 1 <= p.info["create_time"] <= self._start_wall + 5]
            if len(cands) == 1:
                pid, ctime = cands[0].pid, cands[0].info["create_time"]
            elif cands:
                log.warning("hang during start-up but %d candidate Excels; killing none", len(cands))
        wins = windows_of(pid) if pid else []
        why = f"{why}; open windows: {', '.join(wins) or 'none visible'}"
        self._mark_dead(why)
        if pid and ctime and kill_ours(pid, ctime):
            log.warning("terminated own Excel pid=%s (%s)", pid, why)
            self.registry.remove(pid)
        deadline = time.monotonic() + config.RESET_WAIT_S

        def blocked() -> bool:                    # the job that was running when we killed is still running
            return thread.busy_since is not None and thread.job_no == job_no and not thread.waiting_reset
        while blocked() and time.monotonic() < deadline:
            time.sleep(0.05)
        if blocked() and self.thread is thread:
            log.error("COM thread still blocked after kill; replacing it")
            self.thread = ComThread()
            self.thread.start()
        if self.pid == pid0:                      # COM refs are dead anyway; the worker recreates on next call
            self.app = self.wb = None
            self.connected.clear()
            self.pid = self.ctime = None


def _explain(fn: str, status: str) -> str:
    parts = status.split("|", 3)
    rc = parts[1] if len(parts) > 1 else "?"
    msg = parts[2] if len(parts) > 2 else ""
    detail = parts[3].strip() if len(parts) > 3 else ""
    low = detail.lower()
    if "cannot connect to provider" in low or "<update>" in low or "empty response" in low:
        hint = "EPM Cloud did not accept the connection (transient; usually works on retry)"
    elif rc == "-15" and re.search(r"sid", low):
        hint = ("Smart View is not signed in yet (no session id): wait until the EPM login has completed, or ask the "
                "user to sign in once in their Excel, then retry")
    elif rc == "57704" and "mbrs" in low:
        hint = ("the EPM Cloud Planning provider does not support this member query; use sv_find_member / sv_members "
                "(atlas) instead")
    else:
        hint = {
            "-15": "invalid parameter",
            "-41": "not supported by the Planning provider (official)",
            "1000": "Smart View rejected the grid; check member names with sv_find_member",
            "-71": "invalid member name",
        }.get(rc, "")
    if "<update>" in detail:                     # EPM sent its extension-update manifest instead of a reply
        detail = detail.split("<?xml")[0].strip() + " [extension-update manifest omitted]"
    return f"{fn.removeprefix('B_')} failed (rc={rc} {msg}{'; ' + hint if hint else ''}). Server says: {detail or '-'}"
