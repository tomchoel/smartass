"""smartass doctor: cheap health checks (starts neither Excel nor a browser). Each check gives OK / WARN / FAIL,
a detail and the fix. Used by the CLI, sv_status and the HTTP /health endpoint."""
from __future__ import annotations

import contextlib
import json
import os
import platform
import sys
import time
import urllib.error
import urllib.request
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path

from smartass import __version__
from smartass.core import config, logsetup

BUNDLE_FIX = "Send the file from `smartass doctor --bundle` to the maintainer."


@dataclass
class Check:
    name: str
    status: str            # OK | WARN | FAIL
    detail: str
    fix: str = ""


def _no_proxy_open(url_or_req, timeout: float):
    return urllib.request.build_opener(urllib.request.ProxyHandler({})).open(url_or_req, timeout=timeout)


def _config() -> Check:
    if config.CONFIG_ERROR:
        return Check("config", "FAIL", config.CONFIG_ERROR,
                     "Fix the TOML syntax, or delete config.toml and re-run installer\\install.ps1.")
    if not config.USER_CONFIG.exists():
        return Check("config", "WARN", f"no {config.USER_CONFIG}", "Run installer\\install.ps1.")
    return Check("config", "OK", str(config.USER_CONFIG))


def _excel() -> Check:
    if config.EXCEL_EXE and os.path.exists(config.EXCEL_EXE):
        return Check("excel", "OK", config.EXCEL_EXE)
    return Check("excel", "FAIL", f"not found: {config.EXCEL_EXE}",
                 "Install desktop Excel (Microsoft 365), or set paths.excel_exe in config.toml.")


def _smartview() -> Check:
    from smartass.epm import sso
    v = sso.smartview_version()
    if v:
        return Check("smartview", "OK", f"{v} ({config.SMARTVIEW_DLL})")
    return Check("smartview", "FAIL", f"HsAddin.dll not found at {config.SMARTVIEW_DLL}",
                 "Install Oracle Smart View (Company Portal), or set paths.smartview_dll in config.toml.")


def _epm(timeout: float = 5) -> Check:
    try:
        urllib.request.urlopen(config.EPM_BASE, timeout=timeout)
    except urllib.error.HTTPError:
        pass                                            # any HTTP answer means the server is reachable
    except (urllib.error.URLError, OSError) as e:
        return Check("epm", "FAIL", f"{config.EPM_BASE} unreachable: {e}",
                     "Connect to the Telenor network (or VPN) and check the proxy.")
    return Check("epm", "OK", f"{config.EPM_BASE} reachable (the login itself is checked on the first query)")


def _drill_browser() -> Check:
    from smartass.epm.drill_browser import browser_health
    h = browser_health()
    detail = "; ".join(f"{k}: {v}" for k, v in h["channels"].items())
    if h["usable"]:
        return Check("drill_browser", "OK", detail)
    return Check("drill_browser", "FAIL", detail,
                 "Install Microsoft Edge, or ask IT why the Edge UserDataDir policy is set.")


def _atlas(max_age_days: float = 35) -> Check:
    p = config.ATLAS_DB
    if not p.exists():
        return Check("atlas", "FAIL", f"missing: {p}",
                     "Re-run installer\\install.ps1 (copies the monthly atlas), or call sv_atlas_build.")
    age = (time.time() - p.stat().st_mtime) / 86400
    status = "OK" if age <= max_age_days else "WARN"
    return Check("atlas", status, f"{p} ({age:.0f} days old)",
                 "" if status == "OK" else "Re-run the installer after the monthly release, or call sv_atlas_build.")


def _catalog() -> Check:
    from smartass.knowledge import atlas
    from . import catalog
    try:
        cat = catalog.load()
    except catalog.CatalogError as e:
        return Check("catalog", "FAIL", str(e), "Fix knowledge\\catalog.toml (see the error), or update smartass.")
    if not config.ATLAS_DB.exists():
        return Check("catalog", "WARN", "loaded; members not checked (no atlas)",
                     "Build the atlas with sv_atlas_build.")
    with contextlib.closing(atlas.connect(config.ATLAS_DB)) as db:
        problems, unchecked, notes = catalog.validate(cat, db)
    if problems:
        return Check("catalog", "WARN", f"{len(problems)} member(s) not in the cubes: " + "; ".join(problems[:5]),
                     "Update knowledge\\catalog.toml to the current members (a reorg or rename?).")
    detail = (f"{sum(len(v) for v in cat.pnl.values())} P&L lines and {len(cat.entities)} business areas match the "
              "atlas" + (f"; {'; '.join(notes)}" if notes else ""))
    if unchecked:
        return Check("catalog", "WARN", f"{detail}; not checked: {', '.join(unchecked)}",
                     "Build the missing dimensions with sv_atlas_build.")
    return Check("catalog", "OK", detail)


def _output() -> Check:
    d = config.ANALYSER_DIR
    try:
        d.mkdir(parents=True, exist_ok=True)
        t = d / f".smartass_write_test_{os.getpid()}"
        t.write_text("ok", encoding="utf-8")
        t.unlink()
    except OSError as e:
        return Check("output", "FAIL", f"cannot write {d}: {e}",
                     "Set paths.analyser in config.toml to a folder you can write.")
    return Check("output", "OK", str(d))


def _http(timeout: float = 2) -> Check:
    if not config.HTTP_SECRET:
        return Check("http_api", "WARN", "no http.secret in config.toml", "Run installer\\install.ps1.")
    req = urllib.request.Request(f"http://127.0.0.1:{config.HTTP_PORT}/health",
                                 headers={"X-Smartass-Secret": config.HTTP_SECRET})
    try:
        _no_proxy_open(req, timeout).read()
    except (urllib.error.URLError, OSError) as e:
        return Check("http_api", "WARN", f"not answering on 127.0.0.1:{config.HTTP_PORT} ({e})",
                     "Log off and on (Startup entry), or run `smartass serve`.")
    return Check("http_api", "OK", f"127.0.0.1:{config.HTTP_PORT}")


def _update() -> Check:
    from . import update
    r = update.check()
    if r["status"] == "not configured":
        return Check("update", "OK", "release folder not configured (release.folder)")
    if r["status"] == "error":
        return Check("update", "WARN", r["error"], "Check that the release folder is synced (OneDrive).")
    if r["update_available"]:
        return Check("update", "WARN", f"{r['current']} -> {r['target']}", "Run `smartass update` (or log off and on).")
    return Check("update", "OK", f"{r['current']} is current")


def _read_hkcu(path: str, name: str):
    """An HKCU registry value, or None (missing, or not Windows)."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as k:
            return winreg.QueryValueEx(k, name)[0]
    except (ImportError, OSError):
        return None


def _smartview_startup(read=_read_hkcu) -> Check:
    """Smart View settings that slow every Excel start (found 2026-11-03): the extension update check at Office
    start signs in to EPM ~20 s after Excel opens, and that login window blocks all automation of that Excel;
    diagnostics capture slows Smart View and keeps sign-in cookies in plain text."""
    problems, fixes = [], []
    if str(read(r"Software\Oracle\SmartView\extensions", "CheckOnStartup")) == "1":
        problems.append("Smart View checks for extension updates at every Excel start (an EPM login window "
                        "~20 s after Excel opens, which blocks smartass)")
        fixes.append("Smart View > Options > Extensions: untick 'Check for updates when Microsoft Office starts'")
    if str(read(r"Software\Hyperion Solutions\HyperionSmartView\Options", "Profile")) == "1":
        problems.append("Smart View diagnostics capture is on (slower; sign-in cookies stored in plain text)")
        fixes.append("Smart View > Options > Advanced: turn diagnostics off, then delete "
                     "%APPDATA%\\Oracle\\SmartView\\Diagnostics")
    if problems:
        return Check("smartview_startup", "WARN", "; ".join(problems), "; ".join(fixes) + " (close Excel first).")
    return Check("smartview_startup", "OK", "no login at Excel start-up; diagnostics capture off")


CHECKS = (_config, _excel, _smartview, _smartview_startup, _epm, _drill_browser, _atlas, _catalog, _output, _http,
          _update)


def run(include_http: bool = True) -> list[Check]:
    out = []
    for f in CHECKS:
        if f is _http and not include_http:
            continue
        try:
            out.append(f())
        except Exception as e:   # noqa: BLE001 - a check never breaks the doctor
            out.append(Check(getattr(f, "__name__", "check").lstrip("_"), "FAIL", f"{type(e).__name__}: {e}",
                             BUNDLE_FIX))
    return out


def overall(results: list[Check]) -> str:
    statuses = {c.status for c in results}
    return "FAIL" if "FAIL" in statuses else ("WARN" if "WARN" in statuses else "OK")


def as_dict(results: list[Check]) -> dict:
    return {"status": overall(results), "version": __version__, "checks": [asdict(c) for c in results]}


def to_text(results: list[Check]) -> str:
    lines = [f"smartass {__version__}: {overall(results)}"]
    for c in results:
        lines.append(f"  {c.status:<4}  {c.name:<14} {c.detail}")
        if c.fix and c.status != "OK":
            lines.append(f"        fix: {c.fix}")
    return "\n".join(lines)


def _versions() -> dict:
    from importlib.metadata import PackageNotFoundError, version
    pk = {}
    for name in ("mcp", "pydantic", "pywin32", "psutil", "openpyxl", "playwright"):
        try:
            pk[name] = version(name)
        except PackageNotFoundError:
            pk[name] = None
    from smartass.epm import sso
    return {"smartass": __version__, "python": sys.version, "os": platform.platform(),
            "smartview": sso.smartview_version(), "packages": pk}


def _tail(p: Path, max_bytes: int = 2_000_000) -> str:
    with p.open("rb") as f:
        f.seek(max(0, p.stat().st_size - max_bytes))
        return f.read().decode("utf-8", errors="replace")


def bundle(dest_dir: Path | None = None) -> Path:
    """Zip doctor result, versions, redacted config.toml and the tail of every log, for the maintainer."""
    dest_dir = Path(dest_dir or config.HOME / "bundles")
    dest_dir.mkdir(parents=True, exist_ok=True)
    path = dest_dir / f"smartass-doctor-{time.strftime('%Y%m%d-%H%M%S')}.zip"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("doctor.json", json.dumps(as_dict(run()), indent=2))
        z.writestr("versions.json", json.dumps(_versions(), indent=2))
        if config.USER_CONFIG.exists():
            z.writestr("config.toml", logsetup.redact(config.USER_CONFIG.read_text(encoding="utf-8")))
        for f in sorted((config.HOME / "logs").glob("*.log*")):
            z.writestr(f"logs/{f.name}", logsetup.redact(_tail(f)))
    return path
