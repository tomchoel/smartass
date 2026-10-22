"""Paths and constants. Everything machine-local lives in %LOCALAPPDATA%\\svkit (outside OneDrive)."""
import os
from pathlib import Path

PKG = Path(__file__).resolve().parent
MCP_DIR = PKG.parent
SV_DIR = MCP_DIR.parent                                   # ...\prodreg2026\smartview
BAS_FILES = [SV_DIR / "smartview_body.bas", MCP_DIR / "vba" / "modBridge.bas"]

SVKIT = Path(os.environ.get("SVKIT_DIR", Path(os.environ["LOCALAPPDATA"]) / "svkit"))
WORK_DIR = SVKIT / "work"                                 # copies of workbooks being refreshed
PID_FILE = SVKIT / "pids.json"                            # PIDs this server started (never kill others)
ATLAS_DB = SVKIT / "atlas.db"
SERVER_LOG = SVKIT / "logs" / "server.log"

OUT_DIR = Path(os.environ.get("SV_OUT_DIR", SV_DIR / "output"))
# Business map ("forretningskart"): master Excel + generated JSON, team-readable on OneDrive (concepts, accounts ...)
KNOWLEDGE_DIR = Path(os.environ.get("SV_KNOWLEDGE_DIR", SV_DIR / "knowledge"))
ANALYSER_DIR = Path(os.environ.get("SV_ANALYSER_DIR", SV_DIR.parent / "analyser"))       # new analysis workbooks
BRIDGE_XLAM = SVKIT / "svbridge.xlam"                     # bridge add-in loaded into the user's Excel
# Regulated audit record: never written unless the caller explicitly allows it for that request.
ARCHIVE_ROOT = Path(os.environ.get(
    "SV_ARCHIVE_ROOT", Path.home() / "OneDrive - Telenor" / "Analyse (Telenor Norge AS) - Produktregnskapet"))
EXCEL_EXE = r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"
USER_CALL_TIMEOUT_S = float(os.environ.get("SV_USER_CALL_TIMEOUT", "600"))   # whole analysis incl. several pulls
QUEUE_TIMEOUT_S = float(os.environ.get("SV_QUEUE_TIMEOUT", "900"))
QUERY_LOG = Path(os.environ.get("SV_QUERY_LOG", SV_DIR / "logs" / "queries.jsonl"))

CONNECTIONS = {
    "FinMax": "Actuals incl. FTE (app FinMax, cube FinMax, pod plan1)",
    "FinLite": "Actuals without FTE, lighter/faster (app FinMax, cube FinLite, pod plan1)",
    "FinPlan": "Plan/forecast (app FinPlan, pod stratmodel). Different dimensions: Scenario/Version/Years/Period",
}

# Telenor sensitivity label applied before any Excel SaveAs (unlabelled saves are silently cancelled).
LABEL_ID = os.environ.get("SV_LABEL_ID", "f604d2c9-1577-460e-b668-57374a0216c3")   # "Internal"
LABEL_SITE = os.environ.get("SV_LABEL_SITE", "1676489c-5c72-46b7-ba63-9ab90c4aad44")

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
EPM_BASE = os.environ.get("SV_EPM_BASE", "https://plan1-telenorepm.epm.eu-frankfurt-1.ocs.oraclecloud.com")
# Browsers tried in order. Edge first: an HKCU Chrome policy (UserDataDir) forces every Chrome launch onto the user's
# own profile, so a hidden Chrome hands off to the user's running Chrome and exits (code 21). A channel whose
# UserDataDir policy is set is skipped up front; a channel that fails to launch is retried once, then the next one.
DRILL_CHANNELS = tuple(c.strip() for c in os.environ.get("SV_DRILL_CHANNELS", "msedge,chrome").split(",") if c.strip())
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
TEXT_FIXES = KNOWLEDGE_DIR / "text_fixes.json"
