"""Stage timer for one tool call: `with stage("retrieve"):` inside a job, `Timer.current()` to read the totals.
Stages nest by name ("pull.retrieve"); the tool result carries the summary so slow steps are visible per call."""
from __future__ import annotations

import contextlib
import logging
import threading
import time

_local = threading.local()
log = logging.getLogger("smartview.stage")


class Timer:
    def __init__(self):
        self.t0 = time.monotonic()
        self.stages: dict[str, float] = {}
        self.counts: dict[str, int] = {}
        self.stack: list[str] = []

    @staticmethod
    def start() -> "Timer":
        _local.timer = Timer()
        waited = getattr(_local, "pending_wait", None)        # set by the host before the job starts
        if waited:
            _local.timer.stages["wait for your Excel (another agent)"] = round(waited, 2)
            _local.timer.t0 -= waited
        _local.pending_wait = None
        return _local.timer

    @staticmethod
    def note_wait(seconds: float) -> None:
        _local.pending_wait = seconds

    @staticmethod
    def current() -> "Timer | None":
        return getattr(_local, "timer", None)

    def summary(self) -> dict:
        out = {k: round(v, 2) for k, v in sorted(self.stages.items(), key=lambda kv: -kv[1])}
        out["_total"] = round(time.monotonic() - self.t0, 2)
        return out


@contextlib.contextmanager
def stage(name: str):
    t = Timer.current()
    if t is None:
        yield
        return
    full = ".".join(t.stack + [name])
    t.stack.append(name)
    log.info("stage > %s", full)                  # in server.log: pins a crash/hang to the stage it happened in
    t0 = time.monotonic()
    try:
        yield
    finally:
        t.stack.pop()
        t.stages[full] = t.stages.get(full, 0.0) + time.monotonic() - t0
        t.counts[full] = t.counts.get(full, 0) + 1
