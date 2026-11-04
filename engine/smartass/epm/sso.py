"""Smart View EPM Cloud login handling.

The clicking itself is done by the global PowerShell watcher `engine\\sso_watcher.ps1` (.NET UI Automation), started at
logon from the Startup folder. It clicks "Telenor SSO" in the "Login" dialog of ANY Excel (this server's, excel-mcp's,
the user's own) and never types anything. A Python 'uiautomation' implementation was tried first and removed: it
blocked on the dialog in interactive Excel (2026-10-13).

This module only makes sure that watcher is alive (process + fresh heartbeat) and reports the Smart View version.
"""
from __future__ import annotations

import logging
import subprocess
import time

import psutil

from smartass.core import config

log = logging.getLogger("smartview.sso")
SV_DLL = config.SMARTVIEW_DLL
WATCHER_PS1 = config.ENGINE_DIR / "sso_watcher.ps1"
POWERSHELL = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
HEARTBEAT = config.SVKIT / "logs" / "sso_watcher.heartbeat"
STALE_S = 90


def smartview_version() -> str | None:
    try:
        import win32api
        info = win32api.GetFileVersionInfo(SV_DLL, "\\")
        ms, ls = info["FileVersionMS"], info["FileVersionLS"]
        return f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"
    except Exception:   # noqa: BLE001
        return None


def global_watcher_pids() -> list[int]:
    out = []
    for p in psutil.process_iter(["name", "cmdline"]):
        try:
            if (p.info["name"] or "").lower() == "powershell.exe" and "sso_watcher.ps1" in " ".join(p.info["cmdline"] or []):
                out.append(p.pid)
        except psutil.Error:
            pass
    return out


def _heartbeat_age() -> float | None:
    try:
        return time.time() - HEARTBEAT.stat().st_mtime
    except OSError:
        return None


def ensure_global_watcher() -> bool:
    """Start the watcher if it is not running, and restart it if its heartbeat is stale (a hung UIA call)."""
    pids = global_watcher_pids()
    age = _heartbeat_age()
    if pids and (age is None or age < STALE_S):
        return True
    for pid in pids:                              # running but stuck: replace it
        try:
            psutil.Process(pid).kill()
            log.warning("killed stuck SSO watcher pid=%s (heartbeat %.0fs old)", pid, age or -1)
        except psutil.Error:
            pass
    subprocess.Popen([POWERSHELL, "-NoProfile", "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass",
                      "-File", str(WATCHER_PS1), "-LogDir", str(config.HOME / "logs")], creationflags=0x08000000)   # CREATE_NO_WINDOW
    time.sleep(2)
    log.info("started global SSO watcher: %s", global_watcher_pids())
    return bool(global_watcher_pids())
