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


def reg_snapshot(read) -> dict:
    """A registry value as JSON: {"present": True, "data"|"hex", "kind"} or {"present": False}. `read()` returns
    (data, kind) like winreg.QueryValueEx and raises FileNotFoundError when the value (or its key) is absent."""
    try:
        data, kind = read()
    except FileNotFoundError:
        return {"present": False}
    if isinstance(data, bytes):
        return {"present": True, "hex": data.hex(), "kind": kind}
    return {"present": True, "data": data, "kind": kind}


def reg_restore(snap: dict, read, write, delete) -> bool:
    """Put the value of `snap` (reg_snapshot) back: write(data, kind), or delete() when it was absent. Nothing is
    written when the value is already as snapped. Returns True when something was changed."""
    if reg_snapshot(read) == snap:
        return False
    if snap.get("present"):
        write(bytes.fromhex(snap["hex"]) if "hex" in snap else snap["data"], snap["kind"])
    else:
        delete()
    return True


SV_OPTION_IDS = {"messageDisplay": (104, int), "bDisplayHomePanelOnStartup": (109, bool),
                 "bCommentDlgEnable": (110, bool)}               # config.GLOBAL_OPTIONS ids in CAOptionsXML


def sv_originals(xml: str) -> dict[str, int | bool]:
    """host.py's smartview_options_original.json ({"104": int, "109": bool, "110": bool}) from a CAOptionsXML text;
    an option the XML lacks (or cannot be read) is left out."""
    out: dict[str, int | bool] = {}
    for name, (oid, kind) in SV_OPTION_IDS.items():
        m = re.search(rf'<option key="{name}" value="([^"]*)"\s*/>', xml or "")
        if not m:
            continue
        v = m.group(1).strip().lower()
        if kind is bool and v in ("0", "1", "true", "false"):
            out[str(oid)] = v in ("1", "true")
        elif kind is int and re.fullmatch(r"-?\d+", v):
            out[str(oid)] = int(v)
    return out


def sv_patch(xml: str, originals: dict) -> str:
    """CAOptionsXML with only the options in `originals` (sv_originals format: {"104": int, "109": bool, ...}) set;
    every other option, and an option the XML lacks, is left as it is. Booleans keep the XML's style (1/0, true/false)."""
    names = {str(oid): (name, kind) for name, (oid, kind) in SV_OPTION_IDS.items()}
    for oid, v in originals.items():
        if str(oid) not in names:
            continue
        name, kind = names[str(oid)]

        def put(m, v=v, kind=kind):
            if kind is bool:
                words = m.group(2).strip().lower() in ("true", "false")
                new = ("true" if v else "false") if words else ("1" if v else "0")
            else:
                new = str(int(v))
            return m.group(1) + new + m.group(3)
        xml = re.sub(rf'(<option key="{name}" value=")([^"]*)("\s*/>)', put, xml)
    return xml


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
