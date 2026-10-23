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


def _tokens(cmdline: str) -> list[str]:
    return [a if a else b for a, b in re.findall(r'"([^"]*)"|(\S+)', cmdline or "")]


def is_engine_excel(cmdline: str, scratch: str) -> bool:
    """An Excel the engine or the matrix owns: the hidden '/automation -Embedding' one, or one started with '/x' on a
    file in the matrix scratch folder. The user's own Excel and any other program are never engine-owned."""
    toks = [t.lower() for t in _tokens(cmdline)]
    if not toks:
        return False
    if toks[0].endswith(".exe") and toks[0].replace("/", "\\").rsplit("\\", 1)[-1] != "excel.exe":
        return False
    if "/automation" in toks and "-embedding" in toks:
        return True
    if "/x" not in toks or not scratch:
        return False
    root = str(scratch).replace("/", "\\").rstrip("\\").lower() + "\\"
    return any(t.replace("/", "\\").startswith(root) for t in toks[toks.index("/x") + 1:])
