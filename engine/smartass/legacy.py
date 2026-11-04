"""The one seam to v1 (smartview_mcp). Everything v2 borrows from v1 is imported here or from v1 modules named
in this docstring: excel.jobs, excel.target, excel.sheets, styling, sheetorder, answer_write._Cells, svhttp, grid,
atlas, server (host, guard, _call_user, lifespan). When v1 is deleted these are lifted into smartass/."""
from __future__ import annotations

import json
import os
from pathlib import Path

import psutil

from smartview_mcp import config as v1config
from smartview_mcp import session as v1session

V2_DIR_NAME = "v2"
RELAY_SCRIPT = "run_server_v2.py"
_configured = False


def _session_info() -> dict:
    me = psutil.Process()
    relay = next((p for p in me.parents() if RELAY_SCRIPT in " ".join(v1session._cmdline(p))), None)
    p = relay or me
    cwd = v1session._cwd(p) or os.getcwd()
    return {"id": f"v2-{p.pid}-{int(p.create_time())}", "cwd": cwd, "project": v1session.safe_name(Path(cwd).name)}


def _alive_relay_v2(p) -> bool:
    try:
        return RELAY_SCRIPT in " ".join(p.cmdline())
    except (psutil.Error, OSError):
        return False


def _set_workbook_v2(path: str) -> None:
    """v1 session.set_workbook, but the live ids carry v2's 'v2-' prefix and come from run_server_v2.py relays, so
    concurrent v2 sessions do not prune each other from sessions_v2.json."""
    info = v1session.info()
    data = v1session._load()
    data[info["id"]] = {"workbook": str(path), "project": info["project"], "cwd": info["cwd"]}
    live = {f"v2-{p.pid}-{int(p.create_time())}" for p in psutil.process_iter() if v1session._alive_relay(p)}
    data = {k: v for k, v in data.items() if k in live or k == info["id"]}
    try:
        v1config.SVKIT.mkdir(parents=True, exist_ok=True)
        v1session._STATE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


def configure_v2() -> None:
    """In this process only: new analysis files go to <analyser>/v2/<project>, and the session workbook is
    remembered in sessions_v2.json under v2's own relay, so v2 never writes into a v1 session workbook."""
    global _configured
    if _configured:
        return
    if v1config.ANALYSER_DIR.name != V2_DIR_NAME:
        v1config.ANALYSER_DIR = v1config.ANALYSER_DIR / V2_DIR_NAME
    v1session._STATE = v1config.SVKIT / "sessions_v2.json"
    v1session._info = _session_info()
    v1session._alive_relay = _alive_relay_v2
    v1session.set_workbook = _set_workbook_v2
    _configured = True
