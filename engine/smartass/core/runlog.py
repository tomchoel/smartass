"""Run log for the v2 answer tools: one JSON line per fin_analyze / fin_drill / fin_refresh call in config.RUN_LOG,
with the wall seconds of its spans (prepare, prepare.fetch, prepare.drill), the drill kinds and the Excel job's stage
timings. serverkit.guard begins and finishes the run; the collector is a contextvars value, so it follows the call
across await and asyncio.to_thread. Outside a run every function here does nothing. Writing never breaks a tool."""
from __future__ import annotations

import contextlib
import contextvars
import json
import logging
import threading
import time
from datetime import datetime

from . import config, logsetup

log = logging.getLogger("smartview")
FIELDS = ("q", "workbook", "question")         # what note() may set

_run: contextvars.ContextVar[dict | None] = contextvars.ContextVar("smartass_run", default=None)


def begin(tool: str) -> None:
    """Start collecting for this tool call (replaces any run left in this context)."""
    _run.set({"tool": tool, "t0": time.monotonic(), "timings": {}, "fields": {}, "lock": threading.Lock()})


def _merge(run: dict, timings: dict) -> None:
    with run["lock"]:
        for k, v in timings.items():
            run["timings"][k] = run["timings"].get(k, 0) + v


@contextlib.contextmanager
def span(name: str):
    """Adds the wall seconds of the block to name (summed when entered more than once). A no-op outside a run."""
    run = _run.get()
    if run is None:
        yield
        return
    t0 = time.monotonic()
    try:
        yield
    finally:
        _merge(run, {name: time.monotonic() - t0})


def add(timings: dict[str, float]) -> None:
    """Merge seconds (or COM-call counts) measured elsewhere: a Timer.summary(), drilltime.take(). Repeated keys sum."""
    run = _run.get()
    if run is not None:
        _merge(run, timings)


def note(**fields) -> None:
    """Set q, workbook and question for the run's line (the last value wins)."""
    run = _run.get()
    if run is not None:
        with run["lock"]:
            run["fields"].update({k: v for k, v in fields.items() if k in FIELDS})


def finish(ok: bool, error: str | None = None) -> dict:
    """End the run: append its line to config.RUN_LOG and return it ({} outside a run). An OSError is logged only."""
    run = _run.get()
    if run is None:
        return {}
    _run.set(None)
    with run["lock"]:
        timings = {k: round(v, 3) for k, v in run["timings"].items()}
        fields = dict(run["fields"])
    timings["_total"] = round(time.monotonic() - run["t0"], 3)
    line = {"ts": datetime.now().astimezone().isoformat(timespec="seconds"), "tool": run["tool"], "ok": ok,
            "error": logsetup.redact(error) if error is not None else None,
            **{k: fields.get(k) for k in FIELDS}, "timings": timings}
    try:
        path = config.RUN_LOG
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(line, ensure_ascii=False, default=str) + "\n")
    except OSError as e:
        log.info("run log not written: %s", logsetup.redact(e))
    return line
