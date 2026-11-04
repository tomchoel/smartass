"""Server helpers shared by the v2 MCP server: the lifespan (drill warm-up, closing the hidden Excel), the tool guard
(any exception -> readable text with one next step), tool annotations, the hidden-Excel host and the user-Excel job
call with progress notes. Lifted from v1 server.py; nothing here imports a tool module."""
from __future__ import annotations

import contextvars
import logging
import re
import time
from contextlib import asynccontextmanager
from pathlib import Path

from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

from . import config, logsetup
from ..xl import instances
from ..xl.host import ExcelHost, HostError, InputError

logsetup.configure(config.SERVER_LOG)
logsetup.add_secret(config.HTTP_SECRET)
log = logging.getLogger("smartview")

READ = ToolAnnotations(read_only_hint=True, open_world_hint=True)          # reads EPM, writes nothing
LOCAL = ToolAnnotations(read_only_hint=True, open_world_hint=False)        # local atlas / status only
# Writes to the user's workbook or to files (never to EPM): not read-only, so clients do not auto-approve them.
WRITE = ToolAnnotations(read_only_hint=False, destructive_hint=False, idempotent_hint=False, open_world_hint=True)
ADMIN = ToolAnnotations(read_only_hint=False, destructive_hint=False, open_world_hint=False)

_host: ExcelHost | None = None


def host() -> ExcelHost:
    global _host
    if _host is None:
        _host = ExcelHost(visible=False)
    return _host


def _warm_up() -> None:
    """Start the hidden browser and sign in, in the background (never awaited; a failure is only logged)."""
    try:
        if not config.DRILL_KEEP_WARM:
            return
        from ..epm import drill_browser
        b = drill_browser.shared_browser()

        def warm(s):
            s.deadline = time.monotonic() + config.DRILL_SETUP_S
            s.warmup = True
            try:
                b._ensure(s)
            finally:
                s.warmup = False
        fut = b.submit(warm)
        fut.add_done_callback(lambda f: (None if f.cancelled() or f.exception() is None else
                                         log.info("drill warm-up failed: %s", str(f.exception())[:200])))
    except Exception as e:   # noqa: BLE001 - best effort
        log.info("drill warm-up not started: %s", str(e)[:200])


@asynccontextmanager
async def lifespan(_server):
    _warm_up()
    try:
        yield {}
    finally:
        try:
            if _host is not None:
                _host.shutdown()
        finally:
            from ..epm import drill_browser
            drill_browser.shutdown_shared()


def _err(e: Exception) -> ToolError:
    return ToolError(with_next_step(logsetup.redact(e), e))


ESCALATE = ("if it repeats, call fin_status (in a terminal: `smartass doctor`) and send the file from "
            "`smartass doctor --bundle` to the maintainer.")
GENERIC_NEXT = "Retry once; " + ESCALATE
NEXT_STEP = {
    "InputError": "Correct the input named above and call the tool again.",
    "HostError": "Call fin_status(restart=True) (restarts the server and its hidden Smart View Excel), then retry; " + ESCALATE,
    "DrillError": "Retry the drill once; " + ESCALATE,
    "CellTooWide": "Narrow the cell (e.g. child entities) and drill again.",
    "TimeoutError": "Retry with a smaller request (fewer members or periods).",
}
# A message that already says what to do ('Next step: ...', 'If it repeats, run fin_status(restart=True).', '... Try again
# shortly.', '; pass force=true') gets no second piece of advice.
_ADVISED = re.compile(
    r"next step|if it (?:repeats|persists)|(?:^|[.;:!?]\s+|,\s*(?:or|then)\s+|\bthen\s+|\s-\s+)"
    r"(?:try|retry|press|pass|use|run|re-run|call|click|open|give|narrow|split|choose|revert|see|check|correct|"
    r"fix|do|work on|ask|wait|select)\b", re.I)


def with_next_step(msg: str, e: BaseException | None = None) -> str:
    """msg with exactly one next step: its own advice if it has some, else the advice for e's class (most
    specific along the class hierarchy: HostTimeout -> HostError, CellTooWide before DrillError), else generic."""
    if _ADVISED.search(msg):
        return msg
    hint = next((NEXT_STEP[c.__name__] for c in type(e).__mro__ if c.__name__ in NEXT_STEP), None) \
        if e is not None else None
    return f"{msg.rstrip(' .')}. Next step: {hint or GENERIC_NEXT}"


def explain(tool_name: str, e: Exception) -> str:
    return with_next_step(f"{tool_name} failed - {type(e).__name__}: {logsetup.redact(e)}", e)


def guard(fn):
    """ANY exception reaches the model as readable text with a next step (the SDK otherwise reports only
    'Error executing tool X' for non-ToolError exceptions)."""
    import functools
    import inspect

    import anyio.to_thread

    is_async = inspect.iscoroutinefunction(fn)

    @functools.wraps(fn)          # keeps __wrapped__, so the SDK builds the same input schema
    async def guarded(*args, **kwargs):
        from mcp.server.mcpserver import Context
        # tools that declare ctx get progress notes from their slow stages (None for the others: no stale ctx)
        _ctx.set(next((v for v in kwargs.values() if isinstance(v, Context)), None))
        _t0.set(time.monotonic())
        try:
            if is_async:
                return await fn(*args, **kwargs)
            # The SDK runs sync tools in a worker thread; keep that.
            return await anyio.to_thread.run_sync(functools.partial(fn, *args, **kwargs))
        except ToolError as e:
            raise ToolError(with_next_step(logsetup.redact(e))) from None
        except Exception as e:   # noqa: BLE001
            log.exception("tool %s failed", fn.__name__)
            raise ToolError(explain(fn.__name__, e)) from None
    return guarded


_ctx: contextvars.ContextVar = contextvars.ContextVar("smartass_tool_ctx", default=None)
_t0: contextvars.ContextVar = contextvars.ContextVar("smartass_tool_t0", default=0.0)   # progress = s since call start


def progress_secs() -> float:
    """Seconds since this tool call started: the MCP progress value (it must only increase within a request)."""
    return round(time.monotonic() - _t0.get(), 1)


def _with_progress(fn):
    """fn(h) whose Timer sends 'still working: <stage>' notes to the client (MCP progress) while a stage runs
    long (timing.PROGRESS_AFTER_S). Only for tools that declare ctx; a client without a progress token ignores them."""
    import asyncio

    from .timing import Timer
    ctx = _ctx.get()
    if ctx is None:
        return fn
    loop = asyncio.get_running_loop()

    def notify(text: str) -> None:
        asyncio.run_coroutine_threadsafe(ctx.report_progress(progress_secs(), None, text), loop)

    def job(h):
        Timer.note_notify(notify)
        try:
            return fn(h)
        finally:                                  # no notes after the job (also when the call was abandoned)
            Timer.note_notify(None)
            t = Timer.current()
            if t is not None and t.notify is notify:
                t.notify = None
                t._ticker(False)
    return job


async def _call(fn, timeout: float | None = None, what: str = "call"):
    """host().call with every failure surfaced as a readable ToolError (with the advice for its class)."""
    try:
        return await host().call(_with_progress(fn), timeout=timeout, what=what)
    except (HostError, InputError, ToolError) as e:
        raise _err(e) from None


def _is_our_analysis_file(path: str) -> bool:
    """A file the MCP itself created in the analyser folder (config.ANALYSER_DIR; never a user's or an archive workbook)."""
    try:
        parent, root = Path(path).resolve().parent, config.ANALYSER_DIR.resolve()
        return (parent == root or parent.parent == root) and not instances.is_archive(path)   # analyser\[<project>\]
    except (OSError, ValueError):
        return False


async def _call_user(fn, what: str):
    """host().call for jobs that drive the user's Excel (never killed on timeout)."""
    try:
        return await host().call(_with_progress(fn), what=what, user_excel=True)
    except (HostError, InputError) as e:
        raise _err(e) from None
