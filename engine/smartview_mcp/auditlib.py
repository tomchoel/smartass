"""Pure helpers for the live audit matrix (tests/live/matrix.py): error classes, verdicts, before/after."""
from __future__ import annotations

import re


def error_class(err: str) -> str:
    lines = (err or "").strip().splitlines()
    first = lines[0] if lines else ""
    return re.sub(r"\d+", "#", first)[:160]


def verdict(ok: bool, secs: float, budget: float, limit: str | None = None) -> str:
    if not ok:
        return "LIMIT" if limit else "FAIL"
    return "SLOW" if secs > budget else "PASS"


def merge_before(now: list[dict], before: list[dict]) -> list[dict]:
    prev = {r["id"]: r["result"] for r in before}
    return [{**r, "before": prev.get(r["id"], "")} for r in now]
