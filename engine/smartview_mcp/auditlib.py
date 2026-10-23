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


def _same_time(a, b) -> bool:
    try:
        return abs(float(a) - float(b)) < 1.0
    except (TypeError, ValueError):
        return False


def _registered_to(pid: int, ctime: float, registry: list[dict], owners: dict[int, float]) -> bool:
    """`pid` (with that create time) is in the engine's pids.json, owned by one of `owners` (pid -> create time)."""
    for i in registry:
        try:
            if i["pid"] != pid or not _same_time(i["ctime"], ctime):
                continue
            owner = i["owner"]
            if owner in owners and _same_time(i["owner_ctime"], owners[owner]):
                return True
        except (KeyError, TypeError):
            continue
    return False


def is_engine_excel(cmdline: str, scratch: str, pid: int, ctime: float, registry: list[dict],
                    owners: dict[int, float]) -> bool:
    """An Excel the MATRIX owns: (a) an '/automation -Embedding' Excel recorded in the engine's pids.json (`registry`)
    with its create time and an owner that is one of the matrix's own child_server.py processes (`owners`: pid ->
    create time), or (b) one started with '/x' on a file in the matrix scratch folder. The user's Excel, the hidden
    Excels of other smartass sessions, the Excel MCP server, a DispatchEx script and any other program are not."""
    toks = [t.lower() for t in _tokens(cmdline)]
    if not toks:
        return False
    if toks[0].endswith(".exe") and toks[0].replace("/", "\\").rsplit("\\", 1)[-1] != "excel.exe":
        return False
    if "/automation" in toks and "-embedding" in toks:
        return _registered_to(pid, ctime, registry, owners)
    if "/x" not in toks or not scratch:
        return False
    root = str(scratch).replace("/", "\\").rstrip("\\").lower() + "\\"
    return any(t.replace("/", "\\").startswith(root) for t in toks[toks.index("/x") + 1:])


def stale_suspended(marker: list[dict], running: dict[int, float]) -> list[int]:
    """PIDs from a suspend marker ([{pid, ctime}]) that are still the same process (`running`: pid -> create time),
    so a crashed matrix run can resume them. A gone process, or a new one on a reused pid, is left alone."""
    out = []
    for i in marker:
        try:
            pid = i["pid"]
            if pid in running and _same_time(i["ctime"], running[pid]):
                out.append(pid)
        except (KeyError, TypeError):
            continue
    return out


def poll_kill(kill, done, timeout: float, interval: float, sleep, clock) -> tuple[list, bool]:
    """Call `kill()` every `interval` s until it kills something (returns a non-empty list), the call under test is
    `done()`, or `timeout` s pass. Returns (killed, call_was_running_at_the_kill)."""
    t0 = clock()
    while True:
        if done():
            return [], False
        killed = kill()
        if killed:
            return killed, True
        if clock() - t0 >= timeout:
            return [], True
        sleep(interval)
