"""Our own Excel's processes (verified by PID + creation time + '-Embedding') and the STA COM thread a job runs on -
the user's Excel's jobs (xl/userjobs.py) and the engine's hidden Excel (xl/hidden.py, xl/host.py) alike."""
from __future__ import annotations

import concurrent.futures as cf
import contextlib
import gc
import json
import logging
import os
import queue
import threading
import time
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
    """Single STA thread that executes callables in order and pumps messages while idle.

    The rule (Task D2, D3): the COM objects a job makes are released on its own thread, the one whose apartment made
    them. A proxy released on another thread while that apartment is alive calls that Excel from there and waits for
    it: for ever, when that Excel has stopped responding (live, 21.11.2026: stack A, between one job's thread and the
    next one's). After each job (_run_job) its outcome goes to the caller first, then the objects it made are released
    here, and only then does the job's 'after' run (host._call_once: its turns in the user's Excel released; a release
    that does not return is watched: host._watch_release).
    What is enforced:
    - the job's result holds no COM object (host.plain), and the frames its error carries are cleared here (_frames,
      _clear): their locals go here, not where the error is read;
    - the garbage collector runs at a job's end only when no other COM job is running in this process (_running), so
      a job's end never releases a cycle a job still running on another thread made. A job that stays blocked (given
      up on in an Excel that does not answer; on a host thread replaced after a reset) counts as running until it
      returns: no job's end collects meanwhile;
    - the engine's own code makes no reference cycle that holds a job's COM objects: the ones found (an error kept in
      a local, a parameter or a dict of a frame its own traceback holds: exceljob._raise, redraw_job and _bring_in,
      runner.execute, _register_xll and _readable) are broken where they are made, so a job's COM objects go by
      reference counting, on its thread, as it lets them go - with no collector at all.
    What is only made unlikely: a cycle that still holds a COM object (made somewhere not found: pywin32, a library)
    is freed by whichever collection meets it first - Python's automatic collector, which runs on whatever thread
    happens to allocate, or the end of a later job that runs alone, on that job's thread. So is garbage a job made in
    the few milliseconds it ran beside another job's collect."""

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
                    fn, fut, *rest = self.q.get(timeout=0.05)
                except queue.Empty:
                    pythoncom.PumpWaitingMessages()
                    continue
                if fn is None:
                    break
                if not fut.set_running_or_notify_cancel():
                    continue
                self.job_no += 1
                self.busy_since = time.monotonic()
                after = rest[0] if rest else None
                try:
                    _run_job(fn, fut)
                finally:
                    fn = fut = rest = None
                    try:
                        if after is not None:
                            after()
                    except Exception:   # noqa: BLE001 - the end of a job must not end this thread
                        log.exception("the end of a job on %s failed", self.name)
                    after = None
                    self.busy_since = None
        finally:
            pythoncom.CoUninitialize()

    def submit(self, fn: Callable[[], Any], after: Callable[[], Any] | None = None) -> cf.Future:
        """fn's outcome as a future; after (if any) runs on this thread once the objects fn made are released."""
        fut: cf.Future = cf.Future()
        self.q.put((fn, fut, after))
        return fut


_running = 0                                      # COM jobs running in this process now (_run_job)
_running_lock = threading.Lock()


def _run_job(fn: Callable[[], Any], fut: cf.Future) -> None:
    """fn() on this thread: its outcome handed to fut first, then the objects it made released here (a release that
    blocks never keeps the outcome from the caller: Task D review, Minor 3) - the frames its error carries cleared,
    and the garbage collector run here only when no other COM job is running in this process (_running): a collect is
    process-wide, and would release a running job's cycles on this thread (Task D2 review, Important 2)."""
    global _running
    with _running_lock:
        _running += 1
    try:
        frames: list = []
        try:
            out = fn()
        except BaseException as e:   # noqa: BLE001 - forwarded to the awaiting coroutine
            frames = _frames(e)                  # before the error leaves this thread: its frames are all ours
            fut.set_exception(e)
        else:
            fut.set_result(out)
            out = None
        _clear(frames)
        frames = None
        with _running_lock:
            alone = _running == 1
        if alone:
            gc.collect()                          # what is left in cycles, released on this thread
    finally:
        with _running_lock:
            _running -= 1


def _frames(e: BaseException) -> list:
    """The frames e's traceback holds - and those of its cause and context -, read while e has not left this thread.
    Once handed over, e's traceback grows by the frames of the thread that raises it again (the event loop's
    coroutines), which must never be cleared from here."""
    frames, todo, seen = [], [e], set()
    while todo:
        x = todo.pop()
        if x is None or id(x) in seen:
            continue
        seen.add(id(x))
        tb = x.__traceback__
        while tb is not None:
            frames.append(tb.tb_frame)
            tb = tb.tb_next
        todo += [x.__cause__, x.__context__]
    return frames


def _clear(frames: list) -> None:
    """The locals of every finished frame of a failed job released now, on this thread. An error handed to the event
    loop otherwise carries the failed job's COM objects (its Target's app and workbook, a batch's ranges) to that
    thread, which releases them when it drops the error (Task D, 21.11.2026). The line numbers stay: the error logs as
    before. A frame still running (this thread's own) is skipped, as traceback.clear_frames skips it."""
    for f in frames:
        try:
            f.clear()
        except RuntimeError:
            pass


def on_host_thread() -> bool:
    """True on an ExcelHost's own COM thread (named 'excel-com'): work there must not wait on threads that call
    on_host, or it deadlocks."""
    t = threading.current_thread()
    return isinstance(t, ComThread) and t.name == "excel-com"
