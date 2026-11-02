"""The Excel host: one dedicated, hidden Excel driven from one STA thread.

All COM happens on ComThread. Calls carry a deadline; when one passes, the watchdog terminates
*our own* Excel process (verified by PID + creation time + "-Embedding"), which makes the blocked
COM call return with an RPC error. The next call starts a fresh Excel. Other Excel processes
(the user's, excel-mcp's) are never touched.
"""
from __future__ import annotations

import logging
import re
import threading
import time
from typing import Any, Callable

import psutil
import pywintypes
import win32com.client as wc
import win32process

from . import config, dialogs, sso
from .hostproc import RETRY_HRESULTS, HiddenExcel, _excel_pids, _PidRegistry, ComThread, kill_ours, windows_of

log = logging.getLogger("smartview.host")

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


class ProviderEmpty(HostError):
    """EPM kept answering 'Empty response from provider' (rc=752/57704) after the in-place retry: the Smart View
    session of the engine's hidden Excel (pid) is broken. notify = the job's progress callback, kept for call()."""

    def __init__(self, msg: str, pid: int | None, notify=None):
        super().__init__(msg)
        self.pid, self.notify = pid, notify


PROVIDER_EMPTY = ("EPM kept answering 'Empty response from provider' after a restart of the engine's Excel. This "
                  "usually clears within a minute: wait, then ask again. If it persists, sign in to Smart View once in "
                  "your own Excel (Smart View ribbon -> Panel -> FinMax) and run `smartass doctor --bundle`.")


def signin_message(d: dialogs.Dialog, again: bool, holders: list[str]) -> str:
    if d.kind == "webview2_fallback" or d.ie:
        cause = ("Smart View showed its sign-in window (WebView2 failed and it fell back to Internet Explorer, whose "
                 "Oracle login page doesn't work)")
    elif d.hung:
        cause = "Smart View's sign-in window stopped responding (Windows marks it 'Not responding')"
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


class ExcelHost(HiddenExcel):
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
        A stuck Smart View sign-in in our hidden Excel (SignInBlocked) kills that Excel; the call is retried once.
        ProviderEmpty restarts only the hidden Excel, with a note to the user, and retries once (a user-Excel job is
        not retried: the caller is asked to ask again)."""
        import asyncio

        from .timing import _send
        for attempt in (1, 2):
            try:
                return await self._call_once(fn, timeout, what, user_excel)
            except ProviderEmpty as e:
                _send(e.notify, "EPM gave no response; restarting the engine's hidden Excel (yours is untouched)")
                await asyncio.to_thread(self._hard_reset, f"{what}: Empty response from provider", e.pid)
                if user_excel:
                    raise HostError("EPM answered 'Empty response from provider'; the engine restarted its own "
                                    "hidden Excel (yours is untouched) - ask again.") from None
                if attempt == 2:
                    raise HostError(PROVIDER_EMPTY) from None
                log.warning("Empty response from provider during %s; restarted the hidden Excel, retrying once", what)
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
        from .excel import jobs
        timeout = timeout or (config.USER_CALL_TIMEOUT_S if user_excel else config.CALL_TIMEOUT_S)
        started = threading.Event()
        token = jobs.JobToken()

        def job():
            started.set()
            jobs.bind_job(token)             # an abandoned job stops at its next COM call (token check)
            if not user_excel:                    # jobs in the user's Excel start the hidden one only on demand
                self._wait_reset()
                self._ensure_started()            # (h.on_host): it costs ~10 s + an SSO login per server start
                return fn(self)
            try:
                return fn(self)
            finally:
                jobs.release_turns()         # the per-Excel turns taken while resolving the target

        if user_excel:
            # its own STA thread per job: parallel tool calls (e.g. subagents of one session) run at the same time
            # in different Excels; work in ONE Excel still takes turns (jobs.Turn)
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

    def run(self, fn: str, *args):
        """Run a bridge function and raise HostError on an ERR| status string."""
        r = self._run(fn, *args)
        status = r[0] if isinstance(r, tuple) else r
        if isinstance(status, str) and status.startswith("ERR|"):
            raise HostError(_explain(fn, status))
        return r

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

    # ---------------- recovery ----------------
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
    elif rc == "-15" and re.search(r"\bsid\b", low):
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
