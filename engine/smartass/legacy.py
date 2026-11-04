"""The one seam to v1 (smartview_mcp). Everything v2 borrows from v1 is imported here or from v1 modules named
in this docstring: excel.jobs, excel.target, excel.sheets, styling, sheetorder, answer_write._Cells, svhttp, grid,
atlas, server (host, guard, _call_user, lifespan). When v1 is deleted these are lifted into smartass/."""
from __future__ import annotations

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
    _configured = True
