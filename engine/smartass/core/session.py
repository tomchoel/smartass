"""Which Claude Code / Copilot session this server belongs to, its project, and the session's analysis workbook.

Every client session starts its own relay (run_server.py) in the session's project folder; the relay outlives code
reloads of the server child, so (relay pid, start time) identifies the session and the relay's working folder names
the project. Each session keeps ONE analysis workbook (analyser\\<project>\\<date time> <first question>.xlsx) for
all its questions. After an /mcp reconnect the relay is new, so the last workbook of the same folder (cwd) is taken
over (the '_last:<cwd>' entry) - but only by the same client process (the relay's parent: the Claude Code or VS Code
process that started it). A new session in the same folder, e.g. Copilot after Claude Code, gets its own new workbook
(10.11: a Copilot session asked about Wholesale and got the folder's old Streamix workbook). Every write of the state file is atomic (a temp file, then os.replace)."""
from __future__ import annotations

import json
import os
import re
import time
from pathlib import Path

import psutil

from . import config

_STATE = config.HOME / "sessions.json"
RELAY_SCRIPT = "run_server.py"
LAST_PREFIX = "_last:"
_info: dict | None = None


def _sid(p) -> str:
    return f"{p.pid}-{int(p.create_time())}"


def info() -> dict:
    """{'id', 'project', 'cwd', 'client'} of this server's session (the relay ancestor; the server process itself if
    none). client = the relay's parent process (the agent that started it), '' when unknown."""
    global _info
    if _info is None:
        me = psutil.Process()
        relay = next((p for p in me.parents() if _alive_relay(p)), None)
        p = relay or me
        cwd = _cwd(p) or os.getcwd()
        _info = {"id": _sid(p), "cwd": cwd, "project": safe_name(Path(cwd).name), "client": _client(relay)}
    return _info


def _client(relay) -> str:
    try:
        parent = relay.parent() if relay is not None else None
        return _sid(parent) if parent is not None else ""
    except (psutil.Error, OSError):
        return ""


def _cmdline(p) -> list[str]:
    try:
        return p.cmdline()
    except (psutil.Error, OSError):
        return []


def _cwd(p) -> str | None:
    try:
        return p.cwd()
    except (psutil.Error, OSError):
        return None


def _alive_relay(p) -> bool:
    return any(Path(a).name == RELAY_SCRIPT for a in _cmdline(p))


def safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", s).strip()[:60] or "Analysis"


def project_dir() -> Path:
    d = config.ANALYSER_DIR / info()["project"]
    d.mkdir(parents=True, exist_ok=True)
    return d


def _load() -> dict:
    try:
        return json.loads(_STATE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save(data: dict) -> None:
    """Atomic: a reader (another session) sees the old file or the new one, never half of it."""
    tmp = _STATE.with_name(f"{_STATE.name}.{os.getpid()}.tmp")
    try:
        _STATE.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
        os.replace(tmp, _STATE)
    except OSError:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass


def workbook() -> str | None:
    """This session's analysis workbook, if it has one and the file still exists. After an /mcp reconnect the id
    lookup finds nothing: then the last workbook of the same folder (cwd), if the file still exists AND it was set by
    the same client process, re-bound under the new id. Another client (a new session) gets None: a new workbook."""
    me = info()
    data = _load()
    own = data.get(me["id"], {}).get("workbook")
    if own and Path(own).exists():
        return own
    entry = data.get(f"{LAST_PREFIX}{me['cwd']}", {})
    last = entry.get("workbook")
    same_client = bool(me.get("client")) and entry.get("client") == me.get("client")
    if last and same_client and Path(last).exists():
        set_workbook(last)
        return last
    return None


def set_workbook(path: str) -> None:
    """Remember this session's workbook (and the folder's last one). Ended sessions are forgotten; the live relays'
    entries and every '_last:' entry are kept, so concurrent sessions never prune each other."""
    me = info()
    data = _load()
    data[me["id"]] = {"workbook": str(path), "project": me["project"], "cwd": me["cwd"]}
    data[f"{LAST_PREFIX}{me['cwd']}"] = {"workbook": str(path), "ts": time.time(),      # survives /mcp reconnects
                                         "client": me.get("client", "")}                # ... of the same client only
    live = set()
    for p in psutil.process_iter():
        try:
            if _alive_relay(p):
                live.add(_sid(p))
        except (psutil.Error, OSError):
            pass
    _save({k: v for k, v in data.items() if k in live or k == me["id"] or k.startswith(LAST_PREFIX)})
