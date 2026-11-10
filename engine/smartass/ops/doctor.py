"""smartass doctor: cheap health checks (starts neither Excel nor a browser). Each check gives OK / WARN / FAIL,
a detail and the fix. Used by the CLI, fin_status and the HTTP /health endpoint."""
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

from .. import __version__
from ..core import config, logsetup

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
    from ..epm import sso
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
    from ..epm.drill_browser import browser_health
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
                     "Re-run installer\\install.ps1 (copies the monthly atlas), or run `smartass atlas build`.")
    age = (time.time() - p.stat().st_mtime) / 86400
    status = "OK" if age <= max_age_days else "WARN"
    return Check("atlas", status, f"{p} ({age:.0f} days old)",
                 "" if status == "OK" else "Re-run the installer after the monthly release, or run `smartass atlas build`.")


def _knowledge() -> Check:
    """The semantic layer (knowledge/semantic) against the current atlas, with the build checks of `smartass
    knowledge build`: FAIL on a build error (a metric member not in the atlas, an alias clash); WARN when knowledge.db
    is missing, older than the atlas or the semantic files, or when drafts await approval."""
    import os

    from ..knowledge.build import BuildReport, _atlas, semantic_hash, check, summary
    from ..knowledge.semantic import SemanticError, load_all
    try:
        sem = load_all(config.SEMANTIC_DIR)
    except (OSError, SemanticError) as e:
        return Check("knowledge", "FAIL", f"{type(e).__name__}: {e}",
                     "Fix the semantic file named in the error (knowledge\\semantic), or update smartass.")
    detail = (f"{len(sem.metrics)} metrics, {len(sem.facts)} facts, {len(sem.concepts)} concepts, "
              f"{len(sem.rules)} rules, {len(sem.entities)} entities, {len(sem.kpis)} kpis, "
              f"{len(sem.vocab)} vocab, {len(sem.notes)} notes")
    if not config.ATLAS_DB.exists():
        return Check("knowledge", "WARN", f"{detail}; members not checked (no atlas)",
                     "Re-run installer\\install.ps1 (copies the monthly atlas), or run `smartass atlas build`.")
    report = BuildReport(out=config.KNOWLEDGE_DB)
    with contextlib.closing(_atlas(config.ATLAS_DB)) as db:
        check(sem.metrics, sem.facts, db, report, sem.concepts, sem)
    if report.errors:
        return Check("knowledge", "FAIL", f"{len(report.errors)} build error(s): " + "; ".join(report.errors[:3]),
                     "Fix the semantic files (a reorg or rename?), then run `smartass knowledge build`.")
    rebuild = "Run `smartass knowledge build` (or `smartass knowledge refresh` after a month-end)."
    k = config.KNOWLEDGE_DB
    if not k.exists():
        return Check("knowledge", "WARN", f"{detail}; knowledge.db not built yet (built on first use)", rebuild)
    problems = []
    if os.path.getmtime(k) < os.path.getmtime(config.ATLAS_DB):
        problems.append("knowledge.db is older than the atlas")
    if summary(k).get("semantic_hash") != semantic_hash(config.SEMANTIC_DIR):
        problems.append("knowledge.db is older than the semantic files")
    if report.drafts:
        problems.append(f"{report.drafts} draft(s) await approval")
        rebuild = "Review the drafts, then `smartass knowledge approve <id|all>`."
    if problems:
        return Check("knowledge", "WARN", f"{detail}; " + "; ".join(problems), rebuild)
    return Check("knowledge", "OK", f"{detail} match the atlas ({k})")


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
        return Check("http_api", "WARN", "no HTTP API secret (secret.bin)", "Run `smartass setup` (or installer\\install.ps1).")
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


_VBOM_KEY = r"Software\Microsoft\Office\16.0\Excel\Security"
_VBOM_POLICY = r"Software\Policies\Microsoft\Office\16.0\Excel\Security"


def _vba_access(read=_read_hkcu) -> Check:
    """Our hidden Excel loads its VBA bridge through wb.VBProject, which Excel refuses unless 'Trust access to the
    VBA project object model' is on (HKCU ...\\Excel\\Security AccessVBOM = 1; a policy value wins)."""
    policy = read(_VBOM_POLICY, "AccessVBOM")
    value = policy if policy is not None else read(_VBOM_KEY, "AccessVBOM")
    if str(value) == "1":
        return Check("vba_access", "OK", "Excel trusts access to the VBA project object model")
    where = "set by policy" if policy is not None else f"HKCU\\{_VBOM_KEY} AccessVBOM = {value}"
    return Check("vba_access", "FAIL", f"Excel blocks access to the VBA project object model ({where}); the "
                 "Smart View bridge cannot load",
                 "Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings: tick 'Trust "
                 "access to the VBA project object model' (close Excel first; a policy value needs IT).")


CHECKS = (_config, _excel, _smartview, _smartview_startup, _vba_access, _epm, _drill_browser, _atlas, _knowledge,
          _output, _http, _update)


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
    from ..epm import sso
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
