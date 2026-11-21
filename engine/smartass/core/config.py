"""Paths and constants. Settings are read in order: engine\\defaults.toml (org defaults, in git) ->
%LOCALAPPDATA%\\smartass\\config.toml (per user) -> environment variables (SV_*, SVKIT_DIR, SMARTASS_*).
Machine-local state lives in HOME (outside OneDrive). No user-specific path is written in code."""
from __future__ import annotations

import os
import tomllib
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]                 # engine/smartass (this file is smartass/core/config.py)
ENGINE_DIR = PKG.parent
REPO_DIR = ENGINE_DIR.parent
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


def oauth_scopes(value) -> dict[str, str]:
    """[oauth.scopes] as {'https://host': scope}: keys are pod base URLs (any case, port or path); an entry that is
    not a URL with a host, or has an empty scope, is left out."""
    from urllib.parse import urlsplit
    out = {}
    for k, v in (value.items() if isinstance(value, dict) else ()):
        host = (urlsplit(str(k).strip()).hostname or "").lower()
        if host and isinstance(v, str) and v.strip():
            out[f"https://{host}"] = v.strip()
    return out


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
_lad = os.environ.get("LOCALAPPDATA")
LEGACY_HOME = Path(_lad) / "svkit" if _lad else None       # old svkit engine's state (read only, never changed)
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

PID_FILE = HOME / "pids.json"                              # PIDs this server started (never kill others)
ATLAS_DB = HOME / "atlas.db"
NEW_WORKBOOK_SEED = HOME / "new_workbook_seed.xlsx"        # a labelled empty workbook; new analysis files are copies
SERVER_LOG = HOME / "logs" / "server.log"

ANALYSER_DIR = _path(_pick("SV_ANALYSER_DIR", "paths.analyser", ONEDRIVE / "smartass" / "analyser"))
OUT_DIR = _path(_pick("SV_OUT_DIR", "paths.output", ONEDRIVE / "smartass" / "output"))
QUERY_LOG = _path(_pick("SV_QUERY_LOG", "paths.query_log", HOME / "logs" / "queries.jsonl"))
RUN_LOG = _path(_pick("SV_RUN_LOG", "paths.run_log", HOME / "logs" / "runs.jsonl"))   # one line per answer (core)
PREFERENCES = _path(_pick("SV_PREFERENCES", "paths.preferences", HOME / "preferences.toml"))  # learned defaults (core)
# Business map ("forretningskart"): core maps in the repo; packs add their own knowledge folders
KNOWLEDGE_DIR = _path(_pick("SV_KNOWLEDGE_DIR", "paths.knowledge", REPO_DIR / "knowledge"))
SEMANTIC_DIR = KNOWLEDGE_DIR / "semantic"                  # metrics/*.toml + facts/*.toml (sub-spec 2), in git
KNOWLEDGE_DB = HOME / "knowledge.db"                       # compiled from SEMANTIC_DIR + the atlas; never in git
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
LABEL_NAME = _pick("SV_LABEL_NAME", "label.name", "Internal")      # written into a built file's label properties
LABEL_CONTENT_BITS = _pick("SV_LABEL_CONTENT_BITS", "label.content_bits", 2)   # ... with the marking bits Excel records

_release = _pick("SMARTASS_RELEASE", "release.folder")
RELEASE_FOLDER = str(_path(_release)) if _release else _release   # %VAR% and ~ expand; empty = not configured
RELEASE_PIN = get(SETTINGS, "release.version")
HTTP_PORT = _checked(as_int(_pick("SMARTASS_HTTP_PORT", "http.port", 47811), "http.port",
                            as_int(get(ORG_SETTINGS, "http.port", 47811), "", 47811)[0]),
                     "SMARTASS_HTTP_PORT" if os.environ.get("SMARTASS_HTTP_PORT") else "config.toml")
SECRET_FILE = HOME / "secret.bin"                          # HTTP API secret, DPAPI-protected (this Windows user)


def _dpapi():
    import win32crypt
    return win32crypt


def protect_secret(text: str) -> bytes:
    return _dpapi().CryptProtectData(text.encode("utf-8"), "smartass http secret", None, None, None, 0)


def read_secret(path: Path = SECRET_FILE) -> str | None:
    """The secret from a DPAPI blob; None when the file is missing or cannot be unprotected (other user/machine)."""
    try:
        blob = path.read_bytes()
        return _dpapi().CryptUnprotectData(blob, None, None, None, 0)[1].decode("utf-8") or None
    except Exception:   # noqa: BLE001 - OSError, pywintypes.error, bad bytes: all mean "no usable secret"
        return None


def load_http_secret(env=os.environ, path: Path = SECRET_FILE, settings: dict | None = None) -> str | None:
    """SMARTASS_HTTP_SECRET, else secret.bin, else a not yet migrated plain http.secret in config.toml
    (`smartass setup` migrates it into secret.bin and removes it from the file)."""
    return (env.get("SMARTASS_HTTP_SECRET") or read_secret(path)
            or get(SETTINGS if settings is None else settings, "http.secret") or None)


HTTP_SECRET = load_http_secret()

# Global Smart View options set while the server runs and restored on shutdown (they are user prefs).
GLOBAL_OPTIONS = {104: 3, 109: False, 110: False}         # messages: none; no home panel; no comment dialog
# Sheet options for our own grids.
# Suppression OFF (6/7/9 rows, 41/42/43 columns, 44 blocks): results are mapped by position, and the user's own
# Smart View defaults would otherwise drop rows and shift every label onto the wrong member.
SHEET_OPTIONS = {5: 0, 6: False, 7: False, 9: False, 41: False, 42: False, 43: False, 44: False,
                 13: "#Missing", 14: "#NoAccess", 17: "#Invalid"}

CALL_TIMEOUT_S = float(os.environ.get("SV_CALL_TIMEOUT", "180"))
# While a hidden-Excel call runs, look for Smart View dialogs (dialogs.py) after DIALOG_CHECK_AFTER_S, every
# DIALOG_POLL_S. A WebView2 Login window gets LOGIN_GRACE_S for the SSO watcher: SSO once needed 18 s; a hung window
# is stuck at once.
DIALOG_CHECK_AFTER_S = 5.0
DIALOG_POLL_S = 2.0
LOGIN_GRACE_S = float(os.environ.get("SV_LOGIN_GRACE", "30"))
TRANSIENT_RETRY_WAIT_S = 3.0                        # before the one retry after rc=752 / rc=57704
RESET_WAIT_S = 10.0                                 # after killing our Excel, wait this long for the COM call to return

# ---------------------------------------------------------------- drill-through (fin_drill)
EPM_BASE = _pick("SV_EPM_BASE", "epm.base")
DRILL_DB = SVKIT / "drill.db"
DRILL_PAGE = 5000                                  # rows per getData call: EPM's offset pages have no stable order
                                                   # (a 623-line cell paged by 500 gave duplicates and a different sum
                                                   # each time, 02.11.2026), so a piece must fit in ONE call
DRILL_SPLIT_AT = min(DRILL_PAGE, int(os.environ.get("SV_DRILL_SPLIT_AT", str(DRILL_PAGE))))
DRILL_MAX_LINES = int(os.environ.get("SV_DRILL_MAX_LINES", "300000"))
DRILL_XL_FULL = int(os.environ.get("SV_DRILL_XL_FULL", "50000"))   # above this: key columns only in Excel
FASTCOM = os.environ.get("SV_FASTCOM", "1") != "0"   # 0: pywin32's standard COM wrapper instead of xl/fastcom.py
INPROC = os.environ.get("SV_INPROC", "1") != "0"     # 0: batches run over COM, never through the add-in (xl/inproc)
INPROC_DIR = HOME / "inproc"                         # the add-in's builds (one folder per hash) and the job files
PAINT_PAUSE_S = float(os.environ.get("SV_PAINT_PAUSE", "0.15"))   # the job's pause for Excel to paint the new answer
                                                     # (only in the active workbook of a visible Excel; 0: none)
SAVE = "sync" if os.environ.get("SV_SAVE", "").strip().lower() == "sync" else "later"   # sync: our file saved in the job
SAVE_DELAY_S = float(os.environ.get("SV_SAVE_DELAY", "0.3"))   # the save after the reply waits this long for the paint
DRILL_KEEP_DAYS = 90
SVHTTP_TIMEOUT_S = float(os.environ.get("SV_HTTP_TIMEOUT", "120"))   # one Smart View HTTP request incl. a sign-in
SVHTTP_SESSIONS = int(os.environ.get("SV_HTTP_SESSIONS", "8"))   # Smart View requests at once per cube, each on its own session
CLOSED_MONTH_TTL_S = float(os.environ.get("SV_CLOSED_MONTH_TTL", "1800"))   # reuse an HTTP-read closed month this long
# Smart View HTTP route: the data source is chosen by Server NAME (several servers share app FinMax, one of them the
# test pod). A pod whose host contains "-test" is refused unless allow_test is true.
SVHTTP_FINMAX_SERVER = str(_pick("SV_HTTP_FINMAX_SERVER", "svhttp.finmax_server", "FinMax"))
SVHTTP_FINPLAN_SERVER = str(_pick("SV_HTTP_FINPLAN_SERVER", "svhttp.finplan_server", "FinPlan_for_DataLoad"))
_allow_test = _pick("SV_HTTP_ALLOW_TEST", "svhttp.allow_test", False)
SVHTTP_ALLOW_TEST = _allow_test is True or str(_allow_test).strip().lower() in ("1", "true", "yes")
# OAuth bearer route (epm/oauth.py): grids and drill lines need all three set (without them a tool call answers
# with the set-up text, oauth.NOT_SET_UP). The scope names ONE environment's service instance,
# so there is one per pod base URL: [oauth.scopes] "https://<pod host>" = "urn:opc:serviceInstanceID=...::all".
OAUTH_IDCS = str(_pick("SMARTASS_OAUTH_IDCS", "oauth.idcs", "") or "").strip().rstrip("/")
OAUTH_CLIENT_ID = str(_pick("SMARTASS_OAUTH_CLIENT_ID", "oauth.client_id", "") or "").strip()
OAUTH_SCOPES = oauth_scopes(get(SETTINGS, "oauth.scopes"))
OAUTH_DIR = HOME / "oauth"                         # one DPAPI-protected sign-in per pod (this Windows user)
DRILL_PARALLEL = int(os.environ.get("SV_DRILL_PARALLEL", "4"))   # getData calls in flight at once (drill_token)
ANSWER_PREFETCH = os.environ.get("SV_ANSWER_PREFETCH", "1") != "0"   # fetch an answer's pulls at once, while Excel opens
DRILL_DERIVE = os.environ.get("SV_DRILL_DERIVE", "1") != "0"   # split siblings' drill queries built from one of EPM's (drill_derive)
TEXT_FIXES = KNOWLEDGE_DIR / "text_fixes.json"              # shipped; replaced by every update
USER_TEXT_FIXES = HOME / "text_fixes.json"                 # the user's own corrections; win over the shipped list
