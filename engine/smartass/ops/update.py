"""Updates from the shared release folder (SharePoint, synced by OneDrive). The check is a file read; applying runs
the release's own installer. A dev checkout (repo with .git) never updates itself."""
from __future__ import annotations

import json
import subprocess
from pathlib import Path

from .. import __version__
from ..core import config

CREATE_NO_WINDOW = 0x08000000


def parse(v: str) -> tuple[int, ...]:
    return tuple(int(p) for p in v.strip().lstrip("v").split("."))


def check(folder: str | None = None, current: str | None = None, pin: str | None = None) -> dict:
    folder = config.RELEASE_FOLDER if folder is None else folder
    current = __version__ if current is None else current
    pin = config.RELEASE_PIN if pin is None else pin
    if not folder:
        return {"status": "not configured", "current": current, "update_available": False}
    try:
        latest = json.loads((Path(folder) / "latest.json").read_text(encoding="utf-8"))["version"]
        target = pin or latest
        avail = parse(target) != parse(current) if pin else parse(latest) > parse(current)
    except (OSError, ValueError, KeyError) as e:
        return {"status": "error", "current": current, "update_available": False,
                "error": f"cannot read {folder}\\latest.json: {e}"}
    installer = Path(folder) / target / "installer" / "install.ps1"
    if avail and not installer.exists():
        return {"status": "error", "current": current, "update_available": False,
                "error": f"release {target} has no installer at {installer}"}
    return {"status": "ok", "current": current, "latest": latest, "target": target,
            "update_available": avail, "installer": str(installer)}


def is_dev_checkout() -> bool:
    return (config.REPO_DIR / ".git").exists()


def _cmd(installer: str) -> list[str]:
    return ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", installer, "-NoDoctor"]


def apply(force: bool = False) -> str:
    if is_dev_checkout() and not force:
        return "dev checkout: update with git pull (or --force to install the release anyway)"
    r = check()
    if not r["update_available"]:
        return f"no update ({r.get('error') or r['status']}, current {r['current']})"
    p = subprocess.run(_cmd(r["installer"]), capture_output=True, text=True, creationflags=CREATE_NO_WINDOW)
    return f"installed {r['target']}" if p.returncode == 0 else f"installer failed ({p.returncode}): {p.stderr[-2000:]}"
