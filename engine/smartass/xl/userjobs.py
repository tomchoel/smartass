"""Jobs in the USER's Excel (serverkit._call_user): each on an STA thread of its own, so parallel tool calls run at
the same time in different Excels (work in ONE Excel takes turns: jobs.Turn); a deadline after which the job is given
up on - never killed: the user's Excel is never closed; the objects a job made released on its own thread before its
turns are freed (hostproc.ComThread, the 0.5.0 rule); every failure a readable text. Split from xl/host.py's
ExcelHost.call(user_excel=True) in light-fast phase 2 Task 5, with the same behaviour; nothing here imports the
engine's hidden Excel or starts it.

Two things a user job may still hand to the hidden Excel, when it needs them (UserJobs.on_host and app, which reach
it through serverkit.host): the Smart View bridge add-in's build when its file for the current sources is missing
(target._attach_bridge -> instances.ensure_bridge_file), and a new analysis file when no Excel is running and no
labelled seed exists (target._new_analysis_file). Task 6 replaces both."""
from __future__ import annotations

import asyncio
import logging
import threading
import time
from typing import Any, Callable, Protocol

import pywintypes

from ..core import config
from ..core.errors import HostError, HostTimeout, InputError
from . import jobs
from .hostproc import DEAD_HRESULTS, ComThread, is_busy, plain, retry_busy

log = logging.getLogger("smartview.host")
RELEASE_WAIT_S = 10.0       # a job's objects not released this long after its answer: the job is given up on


class JobHost(Protocol):
    """What a user job's code (xl/target.py, xl/instances.py) uses of the object the runner hands it (h): Excel's
    busy retry (com), and the hidden Excel's Application and thread for what it still makes there (app, on_host)."""

    @property
    def app(self) -> Any: ...

    def com(self, call: Callable[[], Any], limit_s: float = 120, alive: Callable[[], Any] | None = None,
            rejected: Callable[[float], Any] | None = None) -> Any: ...

    def on_host(self, fn: Callable[[], Any], timeout: float | None = None) -> Any: ...


class UserJobs:
    """The runner of jobs in the user's Excel. hidden: how to reach the engine's hidden Excel (serverkit.host), asked
    only when a job uses app or on_host - never when the runner is made or a job runs."""

    def __init__(self, hidden: Callable[[], Any] | None = None):
        self._hidden = hidden
        self.last_error: str | None = None
        self.restarts = 0                         # Excels that disconnected during a job (_mark_dead)

    com = staticmethod(retry_busy)

    @property
    def app(self):
        """The hidden Excel's Application, for the code on_host runs there (its own thread)."""
        return self._hidden().app

    def on_host(self, fn: Callable[[], Any], timeout: float | None = None):
        """fn on the hidden Excel's thread, that Excel started if needed (HiddenExcel.on_host)."""
        return self._hidden().on_host(fn, timeout)

    def _mark_dead(self, why: str) -> None:
        self.last_error = why
        self.restarts += 1
        log.warning("Excel marked dead: %s", why)

    async def call(self, fn: Callable[[UserJobs], Any], timeout: float | None = None, what: str = "call"):
        """fn(self) as a job in the user's Excel (run)."""
        return await run(self, fn, timeout, what)


async def run(h, fn: Callable[[Any], Any], timeout: float | None, what: str, release_wait: float | None = None):
    """fn(h) on an STA thread of its own, as a job in the USER's Excel. h: what the job is handed, which also records
    a timeout (last_error) and an Excel that disconnected (_mark_dead) - a UserJobs, or an ExcelHost for its own
    call(user_excel=True). On a timeout the blocked thread is abandoned (the job given up on: jobs.give_up), never
    killed. release_wait: how long the release of a job's objects after its answer may take (RELEASE_WAIT_S)."""
    timeout = timeout or config.USER_CALL_TIMEOUT_S
    started = threading.Event()
    released = threading.Event()
    token = jobs.JobToken(what)

    def job():
        started.set()
        jobs.bind_job(token)             # an abandoned job stops at its next COM call (token check)
        return plain(fn(h), what)        # the hidden Excel only on demand (h.on_host): ~10 s + an SSO login

    def after():                         # on the job's thread, once the objects it made are released
        try:
            jobs.release_turns()         # the per-Excel turns taken while resolving the target
            jobs.unbind_job()
        finally:
            released.set()

    # its own STA thread per job: parallel tool calls (e.g. subagents of one session) run at the same time in
    # different Excels; work in ONE Excel still takes turns (jobs.Turn)
    worker = ComThread(name=f"user-job-{what}")
    worker.start()
    fut = worker.submit(job, after)
    worker.q.put((None, None))           # the thread ends after this job
    wrapped = asyncio.wrap_future(fut)
    wrapped.add_done_callback(lambda f: f.cancelled() or f.exception())   # an abandoned job's error is expected
    q0 = time.monotonic()                # the deadline starts when the job STARTS, not while it queues
    while not started.is_set() and not fut.done():
        if time.monotonic() - q0 > config.QUEUE_TIMEOUT_S:
            fut.cancel()
            raise HostError(f"Smart View is still busy with an earlier request after {config.QUEUE_TIMEOUT_S:.0f}s "
                            f"(e.g. an atlas build); {what} was not started. Try again shortly.")
        await asyncio.sleep(0.1)
    try:
        try:
            done, _ = await asyncio.wait({wrapped}, timeout=timeout)     # wait, never wait_for: no cancel
            if not done:
                raise asyncio.TimeoutError
            return wrapped.result()
        finally:
            if wrapped.done():           # it answered (a result or its error): now its release
                _watch_release(released, what, token, RELEASE_WAIT_S if release_wait is None else release_wait)
    except asyncio.TimeoutError:
        jobs.give_up(token)              # only this job: it rolls back when that Excel frees up
        h.last_error = f"{what} exceeded {timeout:.0f}s in the user's Excel"
        raise HostTimeout(
            f"Your Excel did not respond within {timeout:.0f}s during {what} (a dialog, edit mode, or a large "
            "refresh). Nothing was closed. Check Excel for an open dialog, press Esc/Enter, and ask again.") from None
    except pywintypes.com_error as e:
        if e.hresult in DEAD_HRESULTS:
            h._mark_dead(f"Excel disconnected: {e}")
            raise HostError("The Smart View Excel process died during the call; it will restart on the next "
                            "call. Retry.") from None
        log.warning("COM error during %s", what, exc_info=True)
        if is_busy(e):
            raise HostError(f"Excel was busy during {what} - probably a cell is being edited or a dialog is open"
                            ". Press Enter or Esc in Excel and ask again.") from None
        raise HostError(f"COM error during {what}: {e}") from None
    except (HostError, InputError):
        raise
    except Exception as e:   # noqa: BLE001 - always give the model a readable reason
        log.exception("unexpected error during %s", what)
        raise HostError(f"Unexpected {type(e).__name__} during {what}: {e}. Retry once; if it repeats, run "
                        "fin_status(restart=True).") from None


def _watch_release(released: threading.Event, what: str, token: jobs.JobToken, wait: float) -> None:
    """The job has answered; its objects are being released on its thread (ComThread), which holds its turns
    meanwhile. A release that has not returned after wait seconds is into an Excel that stopped answering: the job is
    given up on (jobs.give_up), so a job waiting for that Excel's turn is told so (jobs.ABANDONED) instead of waiting
    silently; the turn itself is freed only when the release returns (never two jobs in one Excel). Watched by a
    thread of its own, not the event loop: the loop of a script's asyncio.run is gone by then."""
    def watch():
        if released.wait(wait):
            return
        if not token.aborted:
            log.warning("%s answered, but the Excel objects it made were not released within %.0fs (that Excel does "
                        "not answer): it is given up on", what, wait)
            jobs.give_up(token)
    threading.Thread(target=watch, name=f"release-watch {what}", daemon=True).start()
