"""Stage timer for one tool call: `with stage("retrieve"):` inside a job, `Timer.current()` to read the totals.
Stages nest by name ("pull.retrieve"); the tool result carries the summary so slow steps are visible per call."""
from __future__ import annotations

import contextlib
import logging
import threading
import time

_local = threading.local()
log = logging.getLogger("smartview.stage")
COM_PREFIX = "com calls."      # summary keys with this prefix are call counts, not seconds
BATCH_PREFIX = "batches."      # ... and these batch counts (xl/inproc: one Excel run of a batch of operations)
PROGRESS_AFTER_S = 10          # a stage running this long sends 'still working: <stage>' (then every 10 s)


class Timer:
    def __init__(self):
        self.t0 = time.monotonic()
        self.stages: dict[str, float] = {}
        self.counts: dict[str, int] = {}
        self.com: dict[str, int] = {}            # COM calls per stage path (nested stages count in their parents)
        self.batches: dict[str, int] = {}        # batches run per stage path (as com)
        self.stack: list[str] = []
        self.notify = None                       # notify(text): progress note to the client (thread-safe)
        self._tick: threading.Timer | None = None
        self._lock = threading.Lock()

    @staticmethod
    def start() -> "Timer":
        _local.timer = Timer()
        _local.timer.notify, _local.pending_notify = getattr(_local, "pending_notify", None), None
        return _local.timer

    @staticmethod
    def note_notify(fn) -> None:
        """Progress callback for the next Timer.start() on this thread (set by the server before the job)."""
        _local.pending_notify = fn

    def _ticker(self, on: bool) -> None:
        """One ticker per timer while a stage runs: every PROGRESS_AFTER_S it names the innermost stage."""
        with self._lock:
            if not on:
                if self._tick:
                    self._tick.cancel()
                self._tick = None
            elif self.notify is not None and self._tick is None:
                self._arm()

    def _arm(self) -> None:                       # under self._lock
        t = threading.Timer(PROGRESS_AFTER_S, lambda: self._fire(t))
        t.daemon = True
        self._tick = t
        t.start()

    def _fire(self, me: threading.Timer) -> None:
        with self._lock:
            if self._tick is not me or not self.stack:
                return                            # stopped (or restarted) meanwhile
            text = f"still working: {'.'.join(self.stack)} ({time.monotonic() - self.t0:.0f}s)"
        _send(self.notify, text)
        with self._lock:
            if self._tick is me:
                self._arm()

    @staticmethod
    def current() -> "Timer | None":
        return getattr(_local, "timer", None)

    def summary(self) -> dict:
        out = {k: round(v, 2) for k, v in sorted(self.stages.items(), key=lambda kv: -kv[1])}
        out["_total"] = round(time.monotonic() - self.t0, 2)
        out.update({f"{COM_PREFIX}{k}": v for k, v in self.com.items()})     # counts, not seconds
        out.update({f"{BATCH_PREFIX}{k}": v for k, v in self.batches.items()})
        return out


def _send(notify, text: str) -> None:
    try:
        if notify is not None:
            notify(text)
    except Exception:   # noqa: BLE001 - a progress note must never break the job
        log.debug("progress note failed", exc_info=True)


def count_com() -> None:
    """One COM call in the current stage (and its parents), for the stage summary. No timer/stage: nothing."""
    _count("com")


def count_batch() -> None:
    """One batch run (xl/inproc/runner.run) in the current stage (and its parents), as count_com."""
    _count("batches")


def _count(what: str) -> None:
    t = Timer.current()
    if t is None or not t.stack:
        return
    counts = getattr(t, what)
    for k in range(1, len(t.stack) + 1):
        key = ".".join(t.stack[:k])
        counts[key] = counts.get(key, 0) + 1


def say(text: str) -> None:
    """A progress note to the user of the running job (if any), and the log. Never raises."""
    log.info("note: %s", text)
    t = Timer.current()
    _send(t and t.notify, text)


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
    t._ticker(True)
    try:
        yield
    finally:
        t.stack.pop()
        if not t.stack:
            t._ticker(False)
        t.stages[full] = t.stages.get(full, 0.0) + time.monotonic() - t0
        t.counts[full] = t.counts.get(full, 0) + 1
