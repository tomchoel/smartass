"""Paths and constants. Settings are read in order: engine\\defaults.toml (org defaults, in git) ->
%LOCALAPPDATA%\\smartass\\config.toml (per user) -> environment variables (SV_*, SVKIT_DIR, SMARTASS_*).
Machine-local state lives in HOME (outside OneDrive). No user-specific path is written in code."""
from __future__ import annotations

import os
import tomllib
from pathlib import Path

PKG = Path(__file__).resolve().parent
ENGINE_DIR = PKG.parent
REPO_DIR = ENGINE_DIR.parent
MCP_DIR = ENGINE_DIR                                       # old name, kept for callers
DEFAULTS_FILE = ENGINE_DIR / "defaults.toml"
PACKS_DIR = REPO_DIR / "packs"
BAS_FILES = [ENGINE_DIR / "vba" / "smartview_body.bas", ENGINE_DIR / "vba" / "modBridge.bas"]


def merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in over.items():
        out[k] = merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def load_settings(defaults_file: Path, user_file: Path) -> tuple[dict, str | None]:
    """Merged settings, plus the error text if a file could not be read. The server still starts; doctor reports it."""
    settings: dict = {}
    error = None
    for f in (defaults_file, user_file):
        try:
            settings = merge(settings, tomllib.loads(Path(f).read_text(encoding="utf-8")))
        except FileNotFoundError:
            continue
        except (OSError, tomllib.TOMLDecodeError) as e:
            error = f"{f}: {e}"
    return settings, error


def get(settings: dict, dotted: str, default=None):
    cur = settings
    for part in dotted.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return default
        cur = cur[part]
    return default if cur is None or cur == "" else cur


def resolve(env, settings: dict, env_name: str | None, dotted: str, default=None):
    """Env var (if set) > user config.toml > defaults.toml > default."""
    if env_name and env.get(env_name):
        return env[env_name]
    return get(settings, dotted, default)


def as_int(value, name: str, fallback: int) -> tuple[int, str | None]:
    """value as int, else the fallback plus a readable error (a wrong-typed user value must not stop the server)."""
    try:
        return int(value), None
    except (TypeError, ValueError):
        return fallback, f"{name} must be an integer, got {value!r} - using {fallback}"


def as_list(value, name: str, fallback: list) -> tuple[list, str | None]:
    """value as a list of strings (a single string is a one-element list), else the fallback plus a readable error."""
    if isinstance(value, str):
        return [value], None
    if isinstance(value, (list, tuple)) and all(isinstance(v, str) for v in value):
        return list(value), None
    return list(fallback), f"{name} must be a list of strings, got {value!r} - using {list(fallback)}"


def _app_path(exe: str) -> str | None:
    import winreg
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(hive, rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe}") as k:
                v = str(winreg.QueryValueEx(k, "")[0]).strip('"')
                if v and os.path.exists(v):
                    return v
        except OSError:
            pass
    return None


def find_first(configured: str | None, app_exe: str | None, candidates: list[str]) -> str | None:
    """Configured path, else the registry App Paths entry, else the first candidate that exists."""
    if configured:
        return os.path.expandvars(configured)
    if app_exe:
        found = _app_path(app_exe)
        if found:
            return found
    for c in candidates:
        p = os.path.expandvars(c)
        if os.path.exists(p):
            return p
    return None


HOME = Path(os.environ.get("SMARTASS_HOME") or os.environ.get("SVKIT_DIR")
            or Path(os.environ["LOCALAPPDATA"]) / "smartass")
SVKIT = HOME                                               # old name, kept for callers
USER_CONFIG = HOME / "config.toml"
SETTINGS, CONFIG_ERROR = load_settings(DEFAULTS_FILE, USER_CONFIG)
ORG_SETTINGS, _ = load_settings(DEFAULTS_FILE, DEFAULTS_FILE.with_name("absent.toml"))   # fallbacks for bad user values


def _pick(env_name: str | None, dotted: str, default=None):
    return resolve(os.environ, SETTINGS, env_name, dotted, default)


def _checked(result, source: str = "config.toml"):
    """Unwrap an as_int/as_list result; its error (if any) is appended to CONFIG_ERROR."""
    global CONFIG_ERROR
    value, err = result
    if err:
        CONFIG_ERROR = f"{CONFIG_ERROR}; {source}: {err}" if CONFIG_ERROR else f"{source}: {err}"
    return value


def _path(v) -> Path:
    return Path(os.path.expandvars(str(v))).expanduser()


ONEDRIVE = Path(os.environ.get("OneDriveCommercial") or os.environ.get("OneDrive") or Path.home())

WORK_DIR = HOME / "work"                                   # copies of workbooks being refreshed
PID_FILE = HOME / "pids.json"                              # PIDs this server started (never kill others)
ATLAS_DB = HOME / "atlas.db"
SERVER_LOG = HOME / "logs" / "server.log"
BRIDGE_XLAM = HOME / "svbridge.xlam"                       # bridge add-in loaded into the user's Excel

ANALYSER_DIR = _path(_pick("SV_ANALYSER_DIR", "paths.analyser", ONEDRIVE / "smartass" / "analyser"))
OUT_DIR = _path(_pick("SV_OUT_DIR", "paths.output", ONEDRIVE / "smartass" / "output"))
QUERY_LOG = _path(_pick("SV_QUERY_LOG", "paths.query_log", HOME / "logs" / "queries.jsonl"))
# Business map ("forretningskart"): core maps in the repo; packs add their own knowledge folders
KNOWLEDGE_DIR = _path(_pick("SV_KNOWLEDGE_DIR", "paths.knowledge", REPO_DIR / "knowledge"))
PACKS = _checked(as_list(get(SETTINGS, "packs.enabled", []), "packs.enabled",
                         as_list(get(ORG_SETTINGS, "packs.enabled", []), "", [])[0]))
# Regulated audit record: never written unless the caller explicitly allows it for that request.
ARCHIVE_ROOT = _path(_pick("SV_ARCHIVE_ROOT", "paths.archive_root",
                           ONEDRIVE / str(get(SETTINGS, "archive.subdir",
                                              "Analyse (Telenor Norge AS) - Produktregnskapet"))))


def knowledge_file(name: str) -> Path:
    """A knowledge file from the core maps, else from an enabled pack. The core path is returned if neither has it."""
    for d in [KNOWLEDGE_DIR, *(PACKS_DIR / p / "knowledge" for p in PACKS)]:
        if (d / name).exists():
            return d / name
    return KNOWLEDGE_DIR / name


_EXCEL = [r"%ProgramFiles%\Microsoft Office\Root\Office16\EXCEL.EXE",
          r"%ProgramFiles(x86)%\Microsoft Office\Root\Office16\EXCEL.EXE",
          r"%ProgramFiles%\Microsoft Office\Office16\EXCEL.EXE"]
EXCEL_EXE = find_first(_pick("SV_EXCEL_EXE", "paths.excel_exe"), "excel.exe", _EXCEL) or os.path.expandvars(_EXCEL[0])
_SV_DLL = [r"%ProgramFiles(x86)%\Oracle\SmartView\bin\HsAddin.dll", r"%ProgramFiles%\Oracle\SmartView\bin\HsAddin.dll"]
SMARTVIEW_DLL = find_first(_pick("SV_SMARTVIEW_DLL", "paths.smartview_dll"), None, _SV_DLL) or os.path.expandvars(_SV_DLL[0])

USER_CALL_TIMEOUT_S = float(os.environ.get("SV_USER_CALL_TIMEOUT", "600"))   # whole analysis incl. several pulls
QUEUE_TIMEOUT_S = float(os.environ.get("SV_QUEUE_TIMEOUT", "900"))

# Telenor sensitivity label applied before any Excel SaveAs (unlabelled saves are silently cancelled).
LABEL_ID = _pick("SV_LABEL_ID", "label.id")
LABEL_SITE = _pick("SV_LABEL_SITE", "label.site")

RELEASE_FOLDER = _pick("SMARTASS_RELEASE", "release.folder")
RELEASE_PIN = get(SETTINGS, "release.version")
HTTP_PORT = _checked(as_int(_pick("SMARTASS_HTTP_PORT", "http.port", 47811), "http.port",
                            as_int(get(ORG_SETTINGS, "http.port", 47811), "", 47811)[0]),
                     "SMARTASS_HTTP_PORT" if os.environ.get("SMARTASS_HTTP_PORT") else "config.toml")
HTTP_SECRET = _pick("SMARTASS_HTTP_SECRET", "http.secret")

CONNECTIONS = {
    "FinMax": "Actuals incl. FTE (app FinMax, cube FinMax, pod plan1)",
    "FinLite": "Actuals without FTE, lighter/faster (app FinMax, cube FinLite, pod plan1)",
    "FinPlan": "Plan/forecast (app FinPlan, pod stratmodel). Different dimensions: Scenario/Version/Years/Period",
}

# Global Smart View options set while the server runs and restored on shutdown (they are user prefs).
GLOBAL_OPTIONS = {104: 3, 109: False, 110: False}         # messages: none; no home panel; no comment dialog
# Sheet options for our own grids.
# Suppression OFF (6/7/9 rows, 41/42/43 columns, 44 blocks): results are mapped by position, and the user's own
# Smart View defaults would otherwise drop rows and shift every label onto the wrong member.
SHEET_OPTIONS = {5: 0, 6: False, 7: False, 9: False, 41: False, 42: False, 43: False, 44: False,
                 13: "#Missing", 14: "#NoAccess", 17: "#Invalid"}

CALL_TIMEOUT_S = float(os.environ.get("SV_CALL_TIMEOUT", "180"))
START_TIMEOUT_S = float(os.environ.get("SV_START_TIMEOUT", "120"))
MAX_CELLS = int(os.environ.get("SV_MAX_CELLS", "50000"))
PREVIEW_ROWS = 60

# ---------------------------------------------------------------- drill-through (sv_drill)
EPM_BASE = _pick("SV_EPM_BASE", "epm.base")
# Browsers tried in order. Edge first: an HKCU Chrome policy (UserDataDir) forces every Chrome launch onto the user's
# own profile, so a hidden Chrome hands off to the user's running Chrome and exits (code 21). A channel whose
# UserDataDir policy is set is skipped up front; a channel that fails to launch is retried once, then the next one.
_ch = os.environ.get("SV_DRILL_CHANNELS")
DRILL_CHANNELS = tuple(c.strip() for c in (_ch.split(",") if _ch else _checked(as_list(
    get(SETTINGS, "drill.channels", ["msedge"]), "drill.channels",
    as_list(get(ORG_SETTINGS, "drill.channels", ["msedge"]), "", ["msedge"])[0]))) if c.strip())
DRILL_PROFILE = SVKIT / "drill_edge"              # Edge profile + lock anchor; other channels: drill_edge_<channel>
DRILL_LAUNCH_TRIES = 2                             # per channel
DRILL_DB = SVKIT / "drill.db"
DRILL_PAGE = 500                                   # rows per getData call (the page's own size)
DRILL_SPLIT_AT = int(os.environ.get("SV_DRILL_SPLIT_AT", "20000"))
DRILL_MAX_LINES = int(os.environ.get("SV_DRILL_MAX_LINES", "300000"))
DRILL_XL_FULL = int(os.environ.get("SV_DRILL_XL_FULL", "50000"))   # above this: key columns only in Excel
DRILL_KEEP_DAYS = 90
DRILL_LOGIN_TIMEOUT_S = 60                         # per attempt; one retry in a fresh tab (spec §6)
DRILL_CAPTURE_TIMEOUT_S = 45
DRILL_IDLE_S = int(os.environ.get("SV_DRILL_IDLE_S", "300"))   # close the hidden browser (and free the profile) when idle
TEXT_FIXES = KNOWLEDGE_DIR / "text_fixes.json"              # shipped; replaced by every update
USER_TEXT_FIXES = HOME / "text_fixes.json"                 # the user's own corrections; win over the shipped list
