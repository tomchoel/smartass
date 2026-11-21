"""Where a recipe's drills spend their time (drill links over HTTP or Excel, the token drill, derived pieces):
thread-safe totals per kind, read once per answer into its timings ('drills.<kind> (n)': seconds summed over the
calls; the drills and links overlap, so the sums do not add up to the wall time)."""
from __future__ import annotations

import threading
import time
from contextlib import contextmanager

_lock = threading.Lock()
_tot: dict[str, list] = {}          # kind -> [count, seconds]


def add(kind: str, secs: float) -> None:
    with _lock:
        e = _tot.setdefault(kind, [0, 0.0])
        e[0] += 1
        e[1] += secs


@contextmanager
def span(kind: str):
    t0 = time.monotonic()
    try:
        yield
    finally:
        add(kind, time.monotonic() - t0)


def timed(kind: str, fn, *args):
    with span(kind):
        return fn(*args)


def take() -> dict[str, float]:
    """{'<kind> (n)': seconds} since the last take(), and reset."""
    with _lock:
        out = {f"{k} ({n})": round(s, 1) for k, (n, s) in _tot.items()}
        _tot.clear()
    return out
