"""Which Claude Code / Copilot session this server belongs to, its project, and the session's analysis workbook.

Every client session starts its own relay (run_server.py) in the session's project folder; the relay outlives code
reloads of the server child, so (relay pid, start time) identifies the session and the relay's working folder names
the project. Each session keeps ONE analysis workbook (analyser\\<project>\\<date time> <first question>.xlsx) for
all its questions; two sessions never share a file unless the user asks for their open workbook."""
from __future__ import annotations

import json
import os
import re
from pathlib import Path

import psutil

from . import config

_STATE = config.SVKIT / "sessions.json"
_info: dict | None = None


def info() -> dict:
    """{'id', 'project', 'cwd'} of this server's session (the relay ancestor; the server process itself if none)."""
    global _info
    if _info is None:
        me = psutil.Process()
        relay = next((p for p in me.parents() if "run_server.py" in " ".join(_cmdline(p))), None)
        p = relay or me
        cwd = _cwd(p) or os.getcwd()
        _info = {"id": f"{p.pid}-{int(p.create_time())}", "cwd": cwd, "project": safe_name(Path(cwd).name)}
    return _info


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


def safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", s).strip()[:60] or "Analyse"


def project_dir() -> Path:
    d = config.ANALYSER_DIR / info()["project"]
    d.mkdir(parents=True, exist_ok=True)
    return d


def _load() -> dict:
    try:
        return json.loads(_STATE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def workbook() -> str | None:
    """This session's analysis workbook, if it has one and the file still exists."""
    w = _load().get(info()["id"], {}).get("workbook")
    return w if w and Path(w).exists() else None


def set_workbook(path: str) -> None:
    data = _load()
    data[info()["id"]] = {"workbook": str(path), "project": info()["project"], "cwd": info()["cwd"]}
    live = {f"{p.pid}-{int(p.create_time())}" for p in psutil.process_iter() if _alive_relay(p)}
    data = {k: v for k, v in data.items() if k in live or k == info()["id"]}   # forget ended sessions
    try:
        config.SVKIT.mkdir(parents=True, exist_ok=True)
        _STATE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


def _alive_relay(p) -> bool:
    try:
        return "run_server.py" in " ".join(p.cmdline())
    except (psutil.Error, OSError):
        return False
