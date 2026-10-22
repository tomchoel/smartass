# smartass sub-project 1 (repo + hardened engine): implementation plan

> **For agentic workers: REQUIRED SUB-SKILL:** use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to carry out this plan task by task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move the smartview MCP into the git repo `C:\dev\smartass`. Give it two-level config, a doctor, a localhost
`/health` API, an installer, an update check, CI and docs. Then switch the owner over without losing anything.

**Architecture:**
- The package stays `smartview_mcp`, now under `engine\`.
- Settings come from three layers, read by `config.py` in this order:
  1. `engine\defaults.toml`;
  2. `%LOCALAPPDATA%\smartass\config.toml`;
  3. environment variables.
- New small modules: `logsetup`, `doctor`, `httpapi`, `update`, `setup_user`, `cli`.
- `server.py` is split into `tools_*` modules. Tool names and behaviour do not change.
- The installer is a thin `install.ps1` that runs uv, then `smartass setup`, then `smartass doctor`.

**Tech stack:**
- Python 3.12 with uv and a lockfile; hatchling build; pytest; ruff.
- stdlib `tomllib`, `http.server` and `zipfile`.
- pywin32 (COM shortcuts); PowerShell 5.1 (installer only).
- GitHub Actions on `windows-latest`.

**Spec:** `docs\specs\2026-10-22-smartass-repo-design.md`.

## Ground rules for every task

**Paths**
- `OLD` = `C:\Users\t992115\OneDrive - Telenor\Projects\prodreg2026\smartview`
- `REPO` = `C:\dev\smartass`

**Commands**
- Run every command from `REPO\engine` unless a step says otherwise.
- Tests: `uv run pytest -q`.
- Lint: `uv run ruff check .`.

**Editing and commits**
- Never edit UTF-8 files with `Get-Content | Set-Content`, because PowerShell 5.1 mangles æ/ø/å. Use the editor tool.
- Commit after every task with `git -C C:\dev\smartass commit`. Local identity and `commit.gpgsign=false` are already
  set in the repo config. End each commit message with
  `Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>`.

**The live MCP**
- The running MCP keeps using `OLD` until Task 16, so nothing breaks for the owner while you work.

## File structure (after the plan)

```
C:\dev\smartass\
├─ .github\workflows\ci.yml          CI: uv sync, ruff, pytest, secret scan
├─ .githooks\pre-commit              runs scripts\check_staged.py
├─ .gitignore  .gitattributes  README.md  CHANGELOG.md
├─ addin\README.md                   placeholder (sub-project 2)
├─ docs\DESIGN.md  docs\troubleshooting.md  docs\release.md  docs\specs\  docs\plans\
├─ installer\install.ps1             per user, no admin, idempotent
├─ scripts\check_staged.py           pre-commit/CI guard (stdlib only)
├─ scripts\publish_release.ps1       release snapshot + latest.json (+ atlas)
├─ knowledge\                        core maps (JSON + review XLSX + build scripts)
├─ packs\produktregnskap\knowledge\  source_*_2025.json (inputs only Produktregnskap needs)
├─ skills\smartview\                 master copy of the skill
└─ engine\
   ├─ pyproject.toml  uv.lock  .python-version  defaults.toml
   ├─ run_server.py  child_server.py  sso_watcher.ps1  gen_atlas_md.py
   ├─ scripts\uia_dump.ps1
   ├─ vba\modBridge.bas  vba\smartview_body.bas
   ├─ smartview_mcp\
   │  ├─ config.py        CHANGED: layered settings, auto-detect, no user paths
   │  ├─ logsetup.py      NEW: rotating log + redaction
   │  ├─ server.py        CHANGED: core only (mcp object, tool decorator, host, lifespan)
   │  ├─ tools_status.py  tools_query.py  tools_ws.py  tools_drill.py  tools_xl.py   NEW (moved code)
   │  ├─ doctor.py  httpapi.py  update.py  setup_user.py  cli.py                     NEW
   │  └─ sso.py, concepts.py, userexcel.py: small edits where noted
   └─ tests\
      ├─ test_*.py (existing) + test_config.py test_paths.py test_logsetup.py test_errors.py test_doctor.py
      │  test_httpapi.py test_update.py test_setup_user.py test_check_staged.py test_smoke_stdio.py
      ├─ stdio_util.py    helper: talk MCP over stdio
      ├─ fixtures\
      └─ live\            manual scripts (moved from tests\), README.md
```

## Decisions made while planning (flagged for the owner)

1. **State folder.** Machine-local state moves from `%LOCALAPPDATA%\svkit` to `%LOCALAPPDATA%\smartass`, as the spec
   says. `SVKIT_DIR` still overrides it. Task 16 copies the owner's atlas, drill cache and browser profile across.
2. **Connection descriptions.** The `CONNECTIONS` text dictionary stays in `config.py`. It describes the cubes and is
   the same for every deployment. `defaults.toml` holds the values that can differ: the EPM URL, label IDs, drill
   channels, the archive folder, the release folder and the HTTP port.
3. **Release format.** A release is an unpacked snapshot folder `<release>\<version>\` plus
   `<release>\latest.json` and `<release>\atlas\atlas.db`. This is simpler than a zip: the installer just copies it.
4. **When updates run.** The update runs at login: `smartass serve` checks `latest.json` first, and if a newer version
   exists it starts the release's installer and exits. That is the spec's "update at start-up when no tool call is
   running" without any locking. A running MCP picks up the new code through the relay's reload.
5. **SSO watcher entry.** The installer also creates the Startup entry for the SSO watcher. Today that entry is a
   hand-made `svkit-sso-watcher.lnk`; colleagues need the same.
6. **Files not carried over.** `modSVKit.bas`, `sso_click.ps1`, `spikes\` and the stray xlsx stay in `OLD`. They are
   unused by the engine.

---

### Task 1: Import the code into the repo layout

**Files**
- Create: `.gitignore`, `.gitattributes`, `addin\README.md`, `engine\`, `knowledge\`,
  `packs\produktregnskap\knowledge\`, `skills\smartview\`, `docs\DESIGN.md`.
- Docs to copy: the old `specs\` and `plans\`.

- [ ] **Step 1: Copy the files** (PowerShell, from any folder)

```powershell
$old = "C:\Users\t992115\OneDrive - Telenor\Projects\prodreg2026\smartview"; $new = "C:\dev\smartass"
$x = @('/XD','__pycache__','.pytest_cache','.venv','/XF','*.pyc','/NFL','/NDL','/NJH','/NJS','/NP')
robocopy "$old\mcp" "$new\engine" /E @x
robocopy "$old\knowledge" "$new\knowledge" /E @x
robocopy "$env:USERPROFILE\.claude\skills\smartview" "$new\skills\smartview" /E @x
robocopy "$old\specs" "$new\docs\specs" /E @x
robocopy "$old\plans" "$new\docs\plans" /E @x
New-Item -ItemType Directory -Force "$new\engine\scripts", "$new\packs\produktregnskap\knowledge", "$new\addin" | Out-Null
Copy-Item "$old\smartview_body.bas" "$new\engine\vba\"
Copy-Item "$old\uia_dump.ps1" "$new\engine\scripts\"
Copy-Item "$old\DESIGN.md" "$new\docs\DESIGN.md"
foreach ($f in 'source_pd_to_prodreg_code_2025.json','source_pd_unresolved_2025.json','source_product_list_2025.json') {
  Move-Item "$new\knowledge\$f" "$new\packs\produktregnskap\knowledge\$f" }
Get-ChildItem $new -Recurse -Include *.db,*.log,*.jsonl -File | Select-Object FullName
```

Expected: the last command prints nothing, because no databases or logs were copied.

- [ ] **Step 2: Point the knowledge build scripts at the pack folder and at `config`**

In `knowledge\build_products_map.py`, replace the `load` function with:

```python
PACK = os.path.join(os.path.dirname(HERE), "packs", "produktregnskap", "knowledge")


def load(name):
    folder = PACK if name.startswith("source_") else HERE        # Produktregnskap inputs live in the pack
    with open(os.path.join(folder, name), encoding="utf-8") as f:
        return json.load(f)
```

In `knowledge\build_accounts_map.py`, replace the `OUT_DIR = ...` line (line 18) and the `ATLAS = ...` line (line 21)
with:

```python
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "engine"))
from smartview_mcp import config as _cfg  # noqa: E402

OUT_DIR = str(_cfg.OUT_DIR)
ATLAS = str(_cfg.ATLAS_DB)
```

Add `import sys` there if it is missing. In `knowledge\build_finplan_map.py`, line 154, replace
`out = os.path.join(os.path.dirname(HERE), "output")` with:

```python
    from build_accounts_map import OUT_DIR as out
```

In all four `build_*_map.py` files, change the text `prodreg2026\\smartview\\knowledge\\` to `smartass\\knowledge\\`.

- [ ] **Step 3: Write `.gitignore`**

```gitignore
.venv/
__pycache__/
*.pyc
.pytest_cache/
.ruff_cache/
*.db
*.db-journal
*.db-wal
*.db-shm
*.log
*.jsonl
logs/
output/
~$*
```

- [ ] **Step 4: Write `.gitattributes`**

```gitattributes
* text=auto
*.bas text eol=crlf
*.ps1 text eol=crlf
*.xlsx binary
```

- [ ] **Step 5: Write `addin\README.md`**

```markdown
# addin

Placeholder for sub-project 2: the Office.js Excel task pane (AI chat + right-click "drill this cell"). It talks to
the local engine over the HTTP API (`engine\smartview_mcp\httpapi.py`). Nothing here is built yet.
```

- [ ] **Step 6: Commit**

```powershell
cd C:\dev\smartass; git add -A; git status --short | Select-Object -First 5; git commit -q -m "Import smartview MCP, knowledge and skill into repo layout"
```

Expected: the commit succeeds. `git ls-files | Select-String '\.db$'` prints nothing.

---

### Task 2: pyproject, uv lock, ruff and the test layout

**Files**
- Modify: `engine\pyproject.toml`, `engine\smartview_mcp\__init__.py`.
- Create: `engine\.python-version`, `engine\tests\live\` (move manual scripts into it).
- Edit: `sso.py` (one comment only).

- [ ] **Step 1: Replace `engine\pyproject.toml`**

```toml
[project]
name = "smartass"
version = "0.2.0"
description = "smartass engine: read-only Smart View (FinMax / FinLite / FinPlan) for AI agents and Excel"
requires-python = ">=3.12"
dependencies = [
    "mcp==2.3.0",
    "pydantic==2.14.1",
    "pywin32==312",
    "psutil==7.2.2",
    "openpyxl",
    "playwright==1.64.0",
]

[dependency-groups]
dev = ["pytest", "anyio", "ruff"]

[project.scripts]
smartview-mcp = "smartview_mcp.server:main"
smartass = "smartview_mcp.cli:main"

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.hatch.build.targets.wheel]
packages = ["smartview_mcp"]

[tool.pytest.ini_options]
testpaths = ["tests"]
norecursedirs = ["live", "fixtures"]

[tool.ruff]
line-length = 120
target-version = "py312"
extend-exclude = ["tests/live"]

[tool.ruff.lint]
select = ["E9", "F"]
```

`uiautomation` is gone; `sso.py` only names it in a comment. `playwright` is now declared. The `cli` entry point is
added in Task 9; until then `smartass` fails at runtime, but nothing calls it before Task 9.

- [ ] **Step 2: Pin Python and add the version**

Write `engine\.python-version` containing `3.12`. Replace `engine\smartview_mcp\__init__.py` with:

```python
"""smartass engine: read-only Smart View MCP server (package name kept from smartview-mcp)."""
__version__ = "0.2.0"
```

- [ ] **Step 3: Move the manual scripts into `tests\live`**

```powershell
cd C:\dev\smartass\engine\tests; New-Item -ItemType Directory -Force live | Out-Null
'build_atlas','close_test_books','compare_numbers','diag_alias','diag_dims','diag_restart','diff_sheet','inspect_answer',
'inspect_rows','live_analysis','live_build','live_drill','live_drill_report','live_refresh','live_relay','live_sessions',
'live_smoke','live_xl','probe_finplan','render_sheet','test_styling_live','verify_gl' | ForEach-Object { git mv "$_.py" "live\$_.py" }
Select-String live\*.py -Pattern 'parents\[1\]|MCP\s*=' | ForEach-Object { "$($_.Filename):$($_.LineNumber): $($_.Line.Trim())" }
```

In every line printed, change `parents[1]` to `parents[2]`, because the scripts are now one folder deeper. Scripts that
use `__file__.rsplit("\\tests", 1)[0]` still resolve to `engine` and need no change.

In `live\diag_restart.py`, line 13, replace the hardcoded path with:

```python
DUMP = str(Path(__file__).resolve().parents[2] / "scripts" / "uia_dump.ps1")
```

Add `from pathlib import Path` there if it is missing.

- [ ] **Step 4: Lock, sync and run the tests on Python 3.12**

```powershell
cd C:\dev\smartass\engine; uv lock; uv sync; $env:PYTHONUTF8="1"; uv run python -c "import sys; print(sys.version)"; uv run pytest -q
```

Expected:
- The printed version starts with `3.12`.
- `138 passed` or more, `0 failed`. Today's count is 139 passed and 1 skipped. The skip was `test_styling_live`, which
  now lives in `live\`.
- If a test fails only on 3.12 because the code uses a 3.13+ API, fix the code. Do not raise the Python version.

- [ ] **Step 5: Ruff**

Run `uv run ruff check .`, then `uv run ruff check . --fix` for the fixes ruff marks as safe. Review the rest by hand.

Expected: `All checks passed!`. Run the tests again afterwards; they must still pass.

- [ ] **Step 6: Commit**

```powershell
cd C:\dev\smartass; git add -A; git commit -q -m "uv project: lockfile, Python 3.12, playwright declared, ruff, live tests split out"
```

---

### Task 3: Stdio smoke test (the safety net for the refactor)

**Files**
- Create: `engine\tests\stdio_util.py`, `engine\tests\test_smoke_stdio.py`.

- [ ] **Step 1: Write the helper `engine\tests\stdio_util.py`**

```python
"""Talk to the MCP server over stdio, as a client does. No Excel or browser is started: listing tools touches neither."""
from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
from pathlib import Path

ENGINE = Path(__file__).resolve().parents[1]


def list_tools(home: Path, timeout: float = 60) -> list[dict]:
    env = {**os.environ, "SMARTASS_HOME": str(home), "PYTHONUTF8": "1", "PYTHONUNBUFFERED": "1"}
    p = subprocess.Popen([sys.executable, "-B", str(ENGINE / "child_server.py")], stdin=subprocess.PIPE,
                         stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, env=env, cwd=str(ENGINE))
    lines: queue.Queue = queue.Queue()
    threading.Thread(target=lambda: [lines.put(raw) for raw in p.stdout], daemon=True).start()

    def send(msg: dict):
        p.stdin.write((json.dumps(msg) + "\n").encode())
        p.stdin.flush()

    def answer(mid: int) -> dict:
        while True:
            msg = json.loads(lines.get(timeout=timeout))
            if msg.get("id") == mid:
                return msg

    try:
        send({"jsonrpc": "2.0", "id": 1, "method": "initialize",
              "params": {"protocolVersion": "2025-06-18", "capabilities": {},
                         "clientInfo": {"name": "smoke", "version": "0"}}})
        assert "result" in answer(1)
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        send({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
        return answer(2)["result"]["tools"]
    finally:
        p.stdin.close()
        try:
            p.wait(30)
        except subprocess.TimeoutExpired:
            p.kill()
```

- [ ] **Step 2: Write the test `engine\tests\test_smoke_stdio.py`**

```python
from stdio_util import list_tools


def test_stdio_server_starts_and_lists_every_tool(tmp_path):
    names = {t["name"] for t in list_tools(tmp_path)}
    assert len(names) >= 36
    assert {"sv_status", "sv_query", "sv_ws_build", "sv_drill", "xl_read", "xl_revert"} <= names
```

- [ ] **Step 3: Run it**

`uv run pytest tests\test_smoke_stdio.py -q`

Expected: it passes. `SMARTASS_HOME` is not read yet, so the server uses `%LOCALAPPDATA%\svkit`. Task 4 makes the
variable count.

Notes:
- If `import stdio_util` fails, add an empty `engine\tests\conftest.py`. pytest's default rootdir import mode then puts
  `tests\` on `sys.path`.
- If the server rejects protocol version `2025-06-18`, use the `protocolVersion` that the server sends back in its
  error.

- [ ] **Step 4: Save the tool snapshot for Task 7**

```powershell
cd C:\dev\smartass\engine; uv run python -c "import json,sys,tempfile,pathlib; sys.path.insert(0,'tests'); from stdio_util import list_tools; t=sorted(list_tools(pathlib.Path(tempfile.mkdtemp())), key=lambda x:x['name']); json.dump(t, open(r'C:\dev\smartass\engine\tests\tools_before.json','w',encoding='utf-8'), indent=1, sort_keys=True); print(len(t))"
```

Expected: `36`. Do not commit `tools_before.json`. Task 7 deletes it.

- [ ] **Step 5: Commit** the two test files: `git add engine/tests/stdio_util.py engine/tests/test_smoke_stdio.py`
  (plus `conftest.py` if you created one), then `git commit -m "Stdio smoke test: server starts and lists all tools"`.

---

### Task 4: Layered config (`defaults.toml` → user `config.toml` → env)

**Files**
- Create: `engine\defaults.toml`, `engine\tests\test_config.py`.
- Rewrite: `engine\smartview_mcp\config.py`.
- Modify: `engine\smartview_mcp\concepts.py:22`, `engine\smartview_mcp\sso.py:22-23,69-70`,
  `engine\sso_watcher.ps1:4,10`, `engine\run_server.py:2,23,42`.

- [ ] **Step 1: Write the failing tests `engine\tests\test_config.py`**

```python
from smartview_mcp import config


def _defaults(tmp_path):
    d = tmp_path / "defaults.toml"
    d.write_text('[epm]\nbase = "https://org"\n[drill]\nchannels = ["msedge", "chrome"]\n', encoding="utf-8")
    return d


def test_user_file_overrides_defaults_and_env_overrides_both(tmp_path):
    u = tmp_path / "config.toml"
    u.write_text('[epm]\nbase = "https://user"\n', encoding="utf-8")
    s, err = config.load_settings(_defaults(tmp_path), u)
    assert err is None
    assert s["epm"]["base"] == "https://user" and s["drill"]["channels"] == ["msedge", "chrome"]
    assert config.resolve({"SV_EPM_BASE": "https://env"}, s, "SV_EPM_BASE", "epm.base") == "https://env"
    assert config.resolve({}, s, "SV_EPM_BASE", "epm.base") == "https://user"
    assert config.resolve({}, s, None, "nope.x", "dflt") == "dflt"


def test_missing_user_file_is_fine_and_bad_toml_is_reported(tmp_path):
    s, err = config.load_settings(_defaults(tmp_path), tmp_path / "absent.toml")
    assert err is None and s["epm"]["base"] == "https://org"
    bad = tmp_path / "bad.toml"
    bad.write_text("[paths\nx = 1\n", encoding="utf-8")
    s, err = config.load_settings(_defaults(tmp_path), bad)
    assert "bad.toml" in err and s["epm"]["base"] == "https://org"


def test_find_first_prefers_configured_then_candidates(tmp_path):
    exe = tmp_path / "EXCEL.EXE"
    exe.write_text("")
    assert config.find_first(str(exe), None, []) == str(exe)
    assert config.find_first(None, None, [str(tmp_path / "nope.exe"), str(exe)]) == str(exe)
    assert config.find_first(None, None, [str(tmp_path / "nope.exe")]) is None


def test_knowledge_file_falls_back_to_enabled_pack(tmp_path, monkeypatch):
    core, packs = tmp_path / "knowledge", tmp_path / "packs"
    (packs / "p1" / "knowledge").mkdir(parents=True)
    core.mkdir()
    (core / "a.json").write_text("{}")
    (packs / "p1" / "knowledge" / "b.json").write_text("{}")
    monkeypatch.setattr(config, "KNOWLEDGE_DIR", core)
    monkeypatch.setattr(config, "PACKS_DIR", packs)
    monkeypatch.setattr(config, "PACKS", ["p1"])
    assert config.knowledge_file("a.json") == core / "a.json"
    assert config.knowledge_file("b.json") == packs / "p1" / "knowledge" / "b.json"
    monkeypatch.setattr(config, "PACKS", [])
    assert config.knowledge_file("b.json") == core / "b.json"


def test_shipped_defaults_parse_and_hold_org_values():
    s, err = config.load_settings(config.DEFAULTS_FILE, config.DEFAULTS_FILE.with_name("absent.toml"))
    assert err is None
    assert s["epm"]["base"].startswith("https://") and s["drill"]["channels"][0] == "msedge"
    assert s["label"]["id"] and s["http"]["port"] > 1024
```

Run `uv run pytest tests\test_config.py -q`. Expected: FAIL with `AttributeError: ... has no attribute 'load_settings'`.

- [ ] **Step 2: Write `engine\defaults.toml`**

```toml
# Org defaults for every smartass user (in git). Per-user overrides: %LOCALAPPDATA%\smartass\config.toml.
# Environment variables (SV_*, SMARTASS_*) override both.

[epm]
base = "https://plan1-telenorepm.epm.eu-frankfurt-1.ocs.oraclecloud.com"

[label]          # Telenor sensitivity label "Internal", applied before any Excel SaveAs
id = "f604d2c9-1577-460e-b668-57374a0216c3"
site = "1676489c-5c72-46b7-ba63-9ab90c4aad44"

[drill]          # browsers tried in order for sv_drill (Edge first: Chrome's UserDataDir policy blocks a hidden Chrome)
channels = ["msedge", "chrome"]

[archive]        # under the user's OneDrive for Business root; never written unless explicitly allowed
subdir = "Analyse (Telenor Norge AS) - Produktregnskapet"

[release]        # SharePoint/Teams release folder (synced local path); empty until it exists
folder = ""

[http]           # localhost-only API (the secret is per user, in config.toml)
port = 47811
```

- [ ] **Step 3: Rewrite `engine\smartview_mcp\config.py`**

The block from `CONNECTIONS = {` down to `PREVIEW_ROWS = 60` stays exactly as it is today. The drill constants below
`EPM_BASE` stay as they are, except the two lines shown.

```python
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


def _pick(env_name: str | None, dotted: str, default=None):
    return resolve(os.environ, SETTINGS, env_name, dotted, default)


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
PACKS = list(get(SETTINGS, "packs.enabled", []) or [])
# Regulated audit record: never written unless the caller explicitly allows it for that request.
ARCHIVE_ROOT = _path(_pick("SV_ARCHIVE_ROOT", "paths.archive_root",
                           ONEDRIVE / get(SETTINGS, "archive.subdir", "Analyse (Telenor Norge AS) - Produktregnskapet")))


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
HTTP_PORT = int(_pick("SMARTASS_HTTP_PORT", "http.port", 47811))
HTTP_SECRET = _pick("SMARTASS_HTTP_SECRET", "http.secret")

# >>> keep unchanged from today: CONNECTIONS, GLOBAL_OPTIONS, SHEET_OPTIONS, CALL_TIMEOUT_S, START_TIMEOUT_S,
#     MAX_CELLS, PREVIEW_ROWS and the "drill-through" comment header

EPM_BASE = _pick("SV_EPM_BASE", "epm.base")
# Browsers tried in order. Edge first: an HKCU Chrome policy (UserDataDir) forces every Chrome launch onto the user's
# own profile, so a hidden Chrome hands off to the user's running Chrome and exits (code 21). A channel whose
# UserDataDir policy is set is skipped up front; a channel that fails to launch is retried once, then the next one.
_ch = os.environ.get("SV_DRILL_CHANNELS")
DRILL_CHANNELS = tuple(c.strip() for c in (_ch.split(",") if _ch else get(SETTINGS, "drill.channels", ["msedge"]))
                       if c.strip())
# >>> keep unchanged from today: DRILL_PROFILE ... DRILL_IDLE_S, TEXT_FIXES
```

When you paste the file:
- Replace the two `# >>> keep unchanged` markers with the original lines. Remove the markers.
- Remove the old definitions of the names above (the old `SV_DIR`, `EXCEL_EXE`, `LABEL_*` and so on) so that each name
  is defined only once.

- [ ] **Step 4: Use the new names in the other modules**

- `concepts.py:22`: `p = config.KNOWLEDGE_DIR / name` → `p = config.knowledge_file(name)`.
- `sso.py:22-23`:
  ```python
  SV_DLL = config.SMARTVIEW_DLL
  WATCHER_PS1 = config.ENGINE_DIR / "sso_watcher.ps1"
  ```
- `sso.py:69-70`: change the `-File` argument list to:
  ```python
  "-File", str(WATCHER_PS1), "-LogDir", str(config.HOME / "logs")], creationflags=0x08000000)   # CREATE_NO_WINDOW
  ```
- `sso.py` docstring line 3: `mcp\\sso_watcher.ps1` → `engine\\sso_watcher.ps1`.
- `sso_watcher.ps1` line 4:
  ```powershell
  param([int]$IntervalMs = 1000, [string]$LogDir = (Join-Path $env:LOCALAPPDATA 'smartass\logs'))
  ```
- `sso_watcher.ps1` line 10: `$logDir = $LogDir`.
- `run_server.py:2`: change the docstring to
  `Registered by installer\install.ps1 as: <engine>\.venv\Scripts\python.exe <this file>`.
- `run_server.py:23`:
  ```python
  LOG = Path(os.environ.get("SMARTASS_HOME") or os.environ.get("SVKIT_DIR")
             or Path(os.environ.get("LOCALAPPDATA", str(HERE))) / "smartass") / "logs" / "proxy.log"
  ```
- `run_server.py:42`, in `fingerprint()`:
  ```python
  files = [CHILD, HERE / "defaults.toml", *sorted((HERE / "smartview_mcp").glob("*.py")), *sorted((HERE / "vba").glob("*.bas"))]
  ```
- `gen_atlas_md.py:8`:
  ```python
  sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
  from smartview_mcp import config  # noqa: E402
  DB = str(config.ATLAS_DB)
  ```
  Add `import sys` if it is missing. In line 31, change `%LOCALAPPDATA%\\svkit\\atlas.db` to
  `%LOCALAPPDATA%\\smartass\\atlas.db`.

- [ ] **Step 5: Run everything**

`uv run pytest -q; uv run ruff check .`

Expected:
- All tests pass, including the 5 new ones and the smoke test. The smoke test now writes into `tmp_path`, because
  `SMARTASS_HOME` is honoured.
- `uv run python -c "from smartview_mcp import config as c; print(c.EXCEL_EXE, c.SMARTVIEW_DLL, c.ANALYSER_DIR)"`
  prints real paths.

- [ ] **Step 6: Commit** with the message `Layered config: defaults.toml, user config.toml, env; auto-detect Excel and Smart View`.

---

### Task 5: No user-specific paths left (spec "Done when" #4)

**Files**
- Create: `engine\tests\test_paths.py`.
- Modify: `engine\tests\test_logic.py:45-55`, `engine\smartview_mcp\server.py` (docstrings at today's lines 79, 618
  and 1355), and any `skills\smartview\**\*.md` line the test reports.

- [ ] **Step 1: Write the failing test `engine\tests\test_paths.py`**

```python
import re
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
# Built from pieces so this file does not match itself.
BANNED = re.compile("|".join(["t99" + "2115", "prodreg" + "2026", "tom_" + "chamkrai", r"OneDrive - Telenor\\Pro" + "jects"]),
                    re.I)
GLOBS = ["engine/**/*.py", "engine/**/*.ps1", "engine/**/*.toml", "engine/**/*.bas", "installer/**/*",
         "skills/**/*.md", "scripts/**/*"]


def test_no_user_specific_paths_in_code_or_skills():
    hits = []
    for g in GLOBS:
        for f in REPO.glob(g):
            if not f.is_file() or ".venv" in f.parts or f.name == "test_paths.py":
                continue
            for n, line in enumerate(f.read_text(encoding="utf-8", errors="replace").splitlines(), 1):
                if BANNED.search(line):
                    hits.append(f"{f.relative_to(REPO)}:{n}: {line.strip()[:100]}")
    assert not hits, "\n".join(hits)
```

Run `uv run pytest tests\test_paths.py -q`. Expected: FAIL. It lists the `test_logic.py`, `server.py` and skill hits.

- [ ] **Step 2: Replace the two path tests at the end of `test_logic.py`**

```python
def test_archive_guard_local_url_and_unsaved(monkeypatch, tmp_path):
    from smartview_mcp import config
    root = tmp_path / "Analyse (Telenor Norge AS) - Produktregnskapet"
    monkeypatch.setattr(config, "ARCHIVE_ROOT", root)
    assert is_archive(str(root / "2025" / "x.xlsx"))
    assert is_archive("https://telenorgroup.sharepoint.com/sites/AnalyseTelenorNorgeAS/Shared Documents/2025/x.xlsx")
    assert not is_archive(str(tmp_path / "analyser" / "a.xlsx"))
    assert not is_archive("Book1")


def test_local_path_maps_personal_onedrive(monkeypatch):
    from smartview_mcp import userexcel
    monkeypatch.setattr(userexcel, "_onedrive_mounts",
                        lambda: [("https://contoso-my.sharepoint.com/personal/a_b/Documents/", r"C:\OD")])
    url = "https://contoso-my.sharepoint.com/personal/a_b/Documents/smartass/analyser/a%20b.xlsx"
    assert local_path(url) == r"C:\OD\smartass\analyser\a b.xlsx"
```

- [ ] **Step 3: Fix the remaining hits**

- In `server.py`, replace every `prodreg2026\\analyser` in docstrings and in `INSTRUCTIONS` with `the analyser folder`
  (`config.ANALYSER_DIR`, set in config.toml).
- In the skill's `.md` files, replace owner paths with the generic wording: "the analyser folder (`paths.analyser` in
  `%LOCALAPPDATA%\smartass\config.toml`)". Replace `%LOCALAPPDATA%\svkit` with `%LOCALAPPDATA%\smartass`.
- Fix any other reported line the same way. Never weaken the test.

- [ ] **Step 4: Run** `uv run pytest -q`. Expected: all tests pass.

- [ ] **Step 5: Commit** with the message `Remove user-specific paths from code, tests and skill; guard with a test`.

---

### Task 6: Rotating, redacted logs

**Files**
- Create: `engine\smartview_mcp\logsetup.py`, `engine\tests\test_logsetup.py`.
- Modify: `engine\smartview_mcp\server.py:26-29`.

- [ ] **Step 1: Write the failing test `engine\tests\test_logsetup.py`**

```python
import logging

from smartview_mcp import logsetup


def test_redact_hides_tokens_and_secrets():
    logsetup.add_secret("s3cr3t-value-123")
    text = ("GET /x?sso_token=abcDEF123456789xyz&a=1 Authorization: Bearer eyJhbGciOi.payload.sig "  # allow-secret
            'secret = "zzzzzzzzzzzzzzzzzzzz" and s3cr3t-value-123')  # allow-secret
    out = logsetup.redact(text)
    assert "abcDEF" not in out and "eyJhbG" not in out and "zzzz" not in out and "s3cr3t" not in out
    assert "sso_token=***" in out


def test_configure_writes_redacted_rotating_log(tmp_path):
    h = logsetup.configure(tmp_path / "logs" / "t.log", max_bytes=200, backups=2)
    try:
        lg = logging.getLogger("smartview.test")
        for i in range(20):
            lg.info("line %s sso_token=abcdefghijklmnopqrstu", i)  # allow-secret
        h.flush()
        files = sorted(p.name for p in (tmp_path / "logs").iterdir())
        assert files[0] == "t.log" and len(files) <= 3
        assert "abcdefghij" not in (tmp_path / "logs" / "t.log").read_text(encoding="utf-8")
    finally:
        logging.getLogger().removeHandler(h)
        h.close()
```

Run it. Expected: FAIL (`ModuleNotFoundError: smartview_mcp.logsetup`).

- [ ] **Step 2: Write `engine\smartview_mcp\logsetup.py`**

```python
"""Engine logging: one rotating file per process kind, with sso tokens, bearer tokens and the HTTP secret redacted."""
from __future__ import annotations

import logging
import re
from logging.handlers import RotatingFileHandler
from pathlib import Path

_PATTERNS = [re.compile(r"(sso_token=)[^&\s'\"]+"),
             re.compile(r"(Bearer\s+)[A-Za-z0-9._~+/=-]+"),
             re.compile(r"((?:X-Smartass-Secret|secret)\s*[:=]\s*['\"]?)[^\s'\",}]+", re.I)]
_secrets: list[str] = []


def add_secret(value: str | None) -> None:
    if value and len(value) >= 8 and value not in _secrets:
        _secrets.append(value)


def redact(text) -> str:
    out = str(text)
    for p in _PATTERNS:
        out = p.sub(r"\1***", out)
    for s in _secrets:
        out = out.replace(s, "***")
    return out


class _RedactingFormatter(logging.Formatter):
    def format(self, record):
        return redact(super().format(record))       # covers tracebacks too


def configure(path: Path, level: int = logging.INFO, max_bytes: int = 5_000_000, backups: int = 5) -> logging.Handler:
    """Attach one rotating handler to the root logger (idempotent per path)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    root = logging.getLogger()
    for h in root.handlers:
        if getattr(h, "_smartass_path", None) == str(path):
            return h
    logging.raiseExceptions = False                  # two clients' servers may rotate the same file; never crash
    h = RotatingFileHandler(path, maxBytes=max_bytes, backupCount=backups, encoding="utf-8", delay=True)
    h._smartass_path = str(path)
    h.setFormatter(_RedactingFormatter("%(asctime)s %(levelname)s %(name)s %(message)s"))
    root.addHandler(h)
    root.setLevel(level)
    return h
```

- [ ] **Step 3: Use it in `server.py`**

Replace the `config.SERVER_LOG.parent.mkdir(...)` and `logging.basicConfig(...)` lines (today's lines 26-28) with:

```python
logsetup.configure(config.SERVER_LOG)
logsetup.add_secret(config.HTTP_SECRET)
```

Then add `logsetup` to the import line: `from . import atlas, config, logsetup, ops, qlog`.

- [ ] **Step 4: Run** `uv run pytest -q; uv run ruff check .`. Expected: all tests pass and ruff is clean.

- [ ] **Step 5: Commit** with the message `Rotating redacted logs`.

---

### Task 7: Split `server.py` into tool modules (no behaviour change)

**Files**
- Create: `engine\smartview_mcp\tools_status.py`, `tools_query.py`, `tools_ws.py`, `tools_drill.py`, `tools_xl.py`.
- Modify: `engine\smartview_mcp\server.py` (keeps only the core) and `engine\tests\test_drill.py` (the monkeypatched
  tests).

**What goes where.** Move each definition with its decorator, comments and blank lines intact. Copy code; do not
rewrite it.

| Module | Definitions (today's names, in today's order) |
|---|---|
| `server.py` (core) | module imports, logging setup, `log`, `Connection`, `READ` `LOCAL` `WRITE` `ADMIN`, `_host`, `_STARTED`, `host`, `lifespan`, `INSTRUCTIONS`, `mcp`, `AxisIn`, `_err`, `tool`, `_call`, `_safe_name`, `main`, plus every helper used by two or more tool modules (see Step 2) |
| `tools_status.py` | `sv_status`, `sv_atlas_build`, `sv_atlas_status`, `sv_show`, `sv_restart`, `sv_atlas_changes` |
| `tools_query.py` | `sv_dimensions`, `sv_query`, `sv_find_member`, `sv_members`, `sv_refresh_workbook`, the "advanced: full ad hoc API" block (`_sheet`, `_grid_preview`, `sv_adhoc_open` … `sv_adhoc_option`, `sv_member_info`), `sv_scope` |
| `tools_ws.py` | the "analysis in the user's workbook" block (`Target` … `sv_ws_build`, incl. `fin`, `SPEC_DOC`, `PullIn`) **except** `_has_drill_cells` and `_drill_rows`, plus `sv_concept` |
| `tools_drill.py` | `_has_drill_cells`, `_drill_rows`, then the "sv_drill" block: `_browser` … `_sv_drill` |
| `tools_xl.py` | `_is_our_analysis_file`, `XL_RULES`, `_xl`, all `xl_*` tools |

- [ ] **Step 1: Find the shared helpers**

```powershell
cd C:\dev\smartass\engine\smartview_mcp
'_call_user','_session','_finish','_tidy','_result','_load_session','fin','_timed','_is_our_analysis_file','_safe_name','_sheet' |
 ForEach-Object { $n=$_; "$n -> " + ((Select-String server.py -Pattern "\b$n\(" | ForEach-Object LineNumber) -join ',') }
```

A helper whose callers fall into two or more of the modules in the table moves to the `server.py` core. For example,
`_call_user` is used by both `tools_ws` and `tools_xl`, so it moves to the core. Write down the final list before you
move anything.

- [ ] **Step 2: Create each tool module with this header** (adjust the imports to what the module uses; ruff `F401`
  and `F821` tell you exactly what is missing or unused):

```python
"""<one line: which tools>. Registered on the shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

from typing import Literal

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from . import atlas, config, ops, qlog
from . import server as core                      # live module ref: read core._host at call time
from .server import ADMIN, LOCAL, READ, WRITE, AxisIn, Connection, _call, _safe_name, log, tool
```

Rules for the moved code:
- **Mutable globals stay with their owner.** `_host` lives in `server`, so tools read `core._host`. `_browser`,
  `_drill_lock` and `_too_wide` live in `tools_drill`.
- **Cross-module helpers.** `tools_drill` uses session/answer helpers from `tools_ws` through
  `from . import tools_ws as ws` and `ws._session(...)`. `tools_ws` never imports `tools_drill` at module level.
  `_answer_job` needs `_has_drill_cells` and `_drill_rows`, so it imports them inside the function:
  `from .tools_drill import _drill_rows, _has_drill_cells`.
- **No `from .tools_x import name` at module level in `server.py`.** Any test may import a tool module first. That
  module imports `server`, which imports the tool modules back while the first one is still half-loaded. Plain module
  imports (`from . import tools_drill`) survive this cycle; name imports do not.
- **Moved functions** keep their names and signatures.

- [ ] **Step 3: Make `server.py` the core**

Delete the moved definitions. In `lifespan`, replace the `_browser` lines with:

```python
        finally:
            from . import tools_drill
            if tools_drill._browser is not None:
                tools_drill._browser.shutdown()
```

At the very end of `server.py`, replace the old `main` block with:

```python
# Tool modules register their tools on `mcp` when imported (module imports only - see the rule above).
from . import tools_drill, tools_query, tools_status, tools_ws, tools_xl  # noqa: E402,F401


def main():
    mcp.run()


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Point the tests at the module that owns each name**

- In `tests\test_drill.py`, replace every `from smartview_mcp import server` with
  `from smartview_mcp import tools_drill as server`. All drill names used by the tests (`_lines_of`, `_drill_cell`,
  `_ALIAS`, `_has_drill_cells`, `_grid_key`, `_prev_drill`, `_restore_lines`, `_gap_warning`, `_drill_rows`,
  `_drill_fetch`, `_too_wide`, `_sv_drill`, `sv_drill`, `_call`, `qlog`, `ToolError`) are then defined or imported in
  `tools_drill`, so the monkeypatches hit the names that the code actually uses.
- In `tests\test_logic.py`, change `from smartview_mcp.server import fin` to `from smartview_mcp.tools_ws import fin`.
- Find any other test reference with `Select-String tests\test_*.py -Pattern "smartview_mcp.server|server\."` and
  repoint it the same way.

- [ ] **Step 5: Verify that nothing changed for clients**

```powershell
cd C:\dev\smartass\engine; uv run ruff check .; uv run pytest -q
uv run python -c "import json,sys,tempfile,pathlib; sys.path.insert(0,'tests'); from stdio_util import list_tools; a=json.load(open('tests/tools_before.json',encoding='utf-8')); b=sorted(list_tools(pathlib.Path(tempfile.mkdtemp())), key=lambda x:x['name']); print('IDENTICAL' if json.dumps(a,sort_keys=True)==json.dumps(b,sort_keys=True) else 'DIFF')"
(Get-Content smartview_mcp\server.py).Count
```

Expected:
- ruff is clean and all tests pass.
- The comparison prints `IDENTICAL`: same names, descriptions and input schemas.
- `server.py` is about 150 lines.

If the comparison prints `DIFF`, diff the two JSON files and fix the move. Do not change the snapshot.

- [ ] **Step 6: Clean up and commit**

```powershell
Remove-Item tests\tools_before.json; cd C:\dev\smartass; git add -A; git commit -q -m "Split server.py into tools_status/query/ws/drill/xl (tool list identical)"
```

---

### Task 8: Every tool error says what failed, why, and the next step

**Files**
- Modify: `engine\smartview_mcp\server.py` (the `tool` decorator).
- Create: `engine\tests\test_errors.py`.

- [ ] **Step 1: Write the failing test `engine\tests\test_errors.py`**

```python
import pytest
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import server


class HostError(Exception):
    pass


def test_explain_names_tool_cause_and_next_step():
    msg = server.explain("sv_query", HostError("Excel did not answer"))
    assert msg.startswith("sv_query failed - HostError: Excel did not answer")
    assert "Next step:" in msg and "sv_restart" in msg
    assert "Next step:" in server.explain("sv_query", ValueError("bad"))


def test_with_next_step_keeps_existing_advice():
    assert server.with_next_step("Bad member X. Next step: use sv_find_member.") == \
        "Bad member X. Next step: use sv_find_member."
    assert "Next step:" in server.with_next_step("Bad member X")


async def test_guard_turns_any_exception_into_a_readable_tool_error():
    async def sv_fake():
        raise RuntimeError("sso_token=abcdefghijklmnopqrstuvwxyz leaked")  # allow-secret
    guarded = server.guard(sv_fake)
    with pytest.raises(ToolError) as e:
        await guarded()
    assert "sv_fake failed - RuntimeError" in str(e.value) and "abcdefghij" not in str(e.value)
```

The async test needs `anyio`'s pytest plugin, which is in the dev dependencies. If pytest collects the async test
without running it, put `@pytest.mark.anyio` on the test and add this fixture to `tests\conftest.py`:

```python
import pytest


@pytest.fixture
def anyio_backend():
    return "asyncio"
```

Run it. Expected: FAIL (`AttributeError: ... 'explain'`).

- [ ] **Step 2: Replace the `tool` decorator in `server.py`**

```python
NEXT_STEP = {
    "HostError": "Excel/Smart View did not respond: call sv_restart, then retry.",
    "DrillError": "The hidden drill browser failed: retry once.",
    "CellTooWide": "Narrow the cell (e.g. child entities) and drill again.",
    "TimeoutError": "Retry with a smaller request (fewer members or periods).",
}
GENERIC_NEXT = ("If it repeats, call sv_status (in a terminal: `smartass doctor`) and send the file from "
                "`smartass doctor --bundle` to the maintainer.")


def with_next_step(msg: str) -> str:
    return msg if "next step" in msg.lower() else f"{msg.rstrip(' .')}. Next step: {GENERIC_NEXT}"


def explain(tool_name: str, e: Exception) -> str:
    base = f"{tool_name} failed - {type(e).__name__}: {logsetup.redact(e)}"
    hint = NEXT_STEP.get(type(e).__name__)
    return f"{base.rstrip(' .')}. Next step: {hint} {GENERIC_NEXT}" if hint else with_next_step(base)


def guard(fn):
    """ANY exception reaches the model as readable text with a next step (the SDK otherwise reports only
    'Error executing tool X' for non-ToolError exceptions)."""
    import functools

    @functools.wraps(fn)
    async def guarded(*args, **kwargs):
        try:
            return await fn(*args, **kwargs)
        except ToolError as e:
            raise ToolError(with_next_step(logsetup.redact(e))) from None
        except Exception as e:   # noqa: BLE001
            log.exception("tool %s failed", fn.__name__)
            raise ToolError(explain(fn.__name__, e)) from None
    return guarded


def tool(annotations: ToolAnnotations = READ):
    def deco(fn):
        return mcp.tool(annotations=annotations)(guard(fn))
    return deco
```

- [ ] **Step 3: Run** `uv run pytest -q`.

Expected: all tests pass. If an existing test compares a full error message with `==`, change it to `in`, so the test
checks the cause and not the added advice. Never remove the assertion.

- [ ] **Step 4: Commit** with the message `Tool errors state cause and next step; secrets redacted`.

---

### Task 9: `smartass doctor` (+ `--bundle`) and the CLI

**Files**
- Create: `engine\smartview_mcp\doctor.py`, `engine\smartview_mcp\cli.py`, `engine\tests\test_doctor.py`.
- Modify: `engine\smartview_mcp\tools_status.py` (`sv_status`).

- [ ] **Step 1: Write the failing tests `engine\tests\test_doctor.py`**

```python
import json
import zipfile

from smartview_mcp import config, doctor


def test_overall_and_exceptions_become_fail(monkeypatch):
    def ok():
        return doctor.Check("a", "OK", "fine")

    def boom():
        raise RuntimeError("kaput")
    monkeypatch.setattr(doctor, "CHECKS", (ok, boom))
    rs = doctor.run()
    assert [c.status for c in rs] == ["OK", "FAIL"] and "kaput" in rs[1].detail and rs[1].fix
    assert doctor.overall(rs) == "FAIL"
    assert doctor.overall([doctor.Check("a", "OK", ""), doctor.Check("b", "WARN", "")]) == "WARN"


def test_config_check(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "CONFIG_ERROR", "config.toml: Expected ']'")
    assert doctor._config().status == "FAIL"
    monkeypatch.setattr(config, "CONFIG_ERROR", None)
    monkeypatch.setattr(config, "USER_CONFIG", tmp_path / "config.toml")
    assert doctor._config().status == "WARN"
    (tmp_path / "config.toml").write_text("", encoding="utf-8")
    assert doctor._config().status == "OK"


def test_output_check_writes_and_cleans_up(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "ANALYSER_DIR", tmp_path / "a")
    assert doctor._output().status == "OK" and list((tmp_path / "a").iterdir()) == []


def test_bundle_has_doctor_versions_logs_and_redacted_config(monkeypatch, tmp_path):
    (tmp_path / "logs").mkdir()
    (tmp_path / "logs" / "server.log").write_text("x sso_token=abcdefghijklmnopqrstu\n", encoding="utf-8")  # allow-secret
    (tmp_path / "config.toml").write_text('[http]\nsecret = "verysecretvalue123456"\n', encoding="utf-8")  # allow-secret
    monkeypatch.setattr(config, "HOME", tmp_path)
    monkeypatch.setattr(config, "USER_CONFIG", tmp_path / "config.toml")
    monkeypatch.setattr(doctor, "CHECKS", (lambda: doctor.Check("a", "OK", "fine"),))
    path = doctor.bundle()
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        assert {"doctor.json", "versions.json", "config.toml", "logs/server.log"} <= names
        assert "verysecret" not in z.read("config.toml").decode()
        assert "abcdefghij" not in z.read("logs/server.log").decode()
        assert json.loads(z.read("doctor.json"))["status"] == "OK"
```

Run it. Expected: FAIL (`ModuleNotFoundError: smartview_mcp.doctor`).

- [ ] **Step 2: Write `engine\smartview_mcp\doctor.py`**

```python
"""smartass doctor: cheap health checks (starts neither Excel nor a browser). Each check gives OK / WARN / FAIL,
a detail and the fix. Used by the CLI, sv_status and the HTTP /health endpoint."""
from __future__ import annotations

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

from . import __version__, config, logsetup

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
    from . import sso
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
    from .drill_browser import browser_health
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


CHECKS = (_config, _excel, _smartview, _epm, _drill_browser, _atlas, _output, _http, _update)


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
    from . import sso
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
```

`doctor` imports `update`, which Task 11 creates. Until then the `_update` check reports FAIL inside `run()` and the
tests still pass, because they patch `CHECKS`. Task 11 makes the check real.

- [ ] **Step 3: Write `engine\smartview_mcp\cli.py`**

```python
"""smartass command line: doctor, serve (HTTP API + update at login), setup (installer), update."""
from __future__ import annotations

import argparse
import json
import sys


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="smartass")
    sub = ap.add_subparsers(dest="cmd", required=True)
    d = sub.add_parser("doctor", help="health checks (cheap)")
    d.add_argument("--bundle", action="store_true", help="zip logs, redacted config and versions for the maintainer")
    d.add_argument("--json", action="store_true")
    sub.add_parser("serve", help="localhost HTTP API (started at login)")
    sub.add_parser("setup", help="per-user setup (run by installer\\install.ps1)")
    u = sub.add_parser("update", help="install the release from the release folder")
    u.add_argument("--force", action="store_true")
    a = ap.parse_args(argv)

    if a.cmd == "doctor":
        from . import doctor
        if a.bundle:
            print(doctor.bundle())
            return 0
        rs = doctor.run()
        print(json.dumps(doctor.as_dict(rs), indent=2) if a.json else doctor.to_text(rs))
        return 1 if doctor.overall(rs) == "FAIL" else 0
    if a.cmd == "serve":
        from . import config, httpapi, logsetup, update
        logsetup.configure(config.HOME / "logs" / "engine.log")
        logsetup.add_secret(config.HTTP_SECRET)
        if update.start_if_available():
            return 0                                  # the installer restarts `serve` when it is done
        return httpapi.serve()
    if a.cmd == "setup":
        from . import setup_user
        for line in setup_user.run():
            print(line)
        return 0
    if a.cmd == "update":
        from . import update
        print(update.apply(force=a.force))
        return 0
    return 2


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Show the doctor in `sv_status`** (`tools_status.py`)

Add `from . import doctor` to the imports. In the dict that `sv_status` returns, add this entry after
`"drill_browser": drill,`:

```python
            "doctor": doctor.as_dict(doctor.run()),
```

Also update the docstring to "... atlas age, which drill browsers can start, and the doctor checks (config, Excel,
Smart View, EPM, output folder, HTTP API, update)."

- [ ] **Step 5: Run**

`uv run pytest -q; uv run ruff check .; uv run smartass doctor`

Expected:
- All tests pass.
- The doctor prints one line per check. `update` shows FAIL until Task 11. `http_api` shows WARN until Task 10.
  `atlas` is FAIL, because `%LOCALAPPDATA%\smartass\atlas.db` does not exist until Task 16.

- [ ] **Step 6: Commit** with the message `smartass doctor (+ --bundle) and CLI; doctor in sv_status`.

---

### Task 10: Localhost HTTP API skeleton (`/health` only)

**Files**
- Create: `engine\smartview_mcp\httpapi.py`, `engine\tests\test_httpapi.py`.

- [ ] **Step 1: Write the failing test `engine\tests\test_httpapi.py`**

```python
import json
import threading
import urllib.error
import urllib.request

import pytest

from smartview_mcp import httpapi

SECRET = "s" * 32


@pytest.fixture
def srv():
    s = httpapi.make_server(0, SECRET, health=lambda: {"status": "OK"})
    threading.Thread(target=s.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{s.server_address[1]}"
    s.shutdown()
    s.server_close()


def _get(url, secret=None):
    req = urllib.request.Request(url, headers={"X-Smartass-Secret": secret} if secret else {})
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    return opener.open(req, timeout=5)


def test_health_needs_the_secret(srv):
    for bad in (None, "wrong"):
        with pytest.raises(urllib.error.HTTPError) as e:
            _get(srv + "/health", bad)
        assert e.value.code == 401
    assert json.loads(_get(srv + "/health", SECRET).read()) == {"status": "OK"}


def test_unknown_path_is_404(srv):
    with pytest.raises(urllib.error.HTTPError) as e:
        _get(srv + "/query", SECRET)
    assert e.value.code == 404


def test_binds_loopback_only():
    s = httpapi.make_server(0, SECRET)
    try:
        assert s.server_address[0] == "127.0.0.1"
    finally:
        s.server_close()


def test_serve_refuses_without_secret(monkeypatch):
    monkeypatch.setattr(httpapi.config, "HTTP_SECRET", "")
    assert httpapi.serve() == 2
```

Run it. Expected: FAIL (`ModuleNotFoundError`).

- [ ] **Step 2: Write `engine\smartview_mcp\httpapi.py`**

```python
"""Localhost-only HTTP API for the future Excel add-in. Sub-project 1 ships only GET /health, behind a per-user
shared secret (header X-Smartass-Secret). Never binds anything but 127.0.0.1."""
from __future__ import annotations

import errno
import hmac
import json
import logging
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from . import config

log = logging.getLogger("smartass.http")


def _default_health() -> dict:
    from . import doctor
    return doctor.as_dict(doctor.run(include_http=False))


def make_server(port: int, secret: str, health=None) -> ThreadingHTTPServer:
    health = health or _default_health

    class Handler(BaseHTTPRequestHandler):
        server_version = "smartass"

        def _send(self, code: int, body: dict):
            data = json.dumps(body).encode()
            self.send_response(code)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def do_GET(self):  # noqa: N802
            given = self.headers.get("X-Smartass-Secret", "")
            if not hmac.compare_digest(given.encode(), secret.encode()):
                return self._send(401, {"error": "missing or wrong X-Smartass-Secret"})
            if self.path.split("?")[0] != "/health":
                return self._send(404, {"error": "unknown path"})
            return self._send(200, health())

        def log_message(self, fmt, *args):
            log.info("%s %s", self.address_string(), fmt % args)

    return ThreadingHTTPServer(("127.0.0.1", port), Handler)


def serve() -> int:
    if not config.HTTP_SECRET:
        log.error("no http.secret in config.toml - run the installer")
        return 2
    try:
        srv = make_server(config.HTTP_PORT, config.HTTP_SECRET)
    except OSError as e:
        if e.errno in (errno.EADDRINUSE, 10048):
            log.info("port %s already in use - another serve is running", config.HTTP_PORT)
            return 0
        raise
    log.info("serving on 127.0.0.1:%s", config.HTTP_PORT)
    srv.serve_forever()
    return 0
```

- [ ] **Step 3: Run** `uv run pytest tests\test_httpapi.py -q; uv run ruff check .`. Expected: 4 passed, ruff is
  clean.
- [ ] **Step 4: Commit** with the message `Localhost HTTP API skeleton: /health behind per-user secret`.

---

### Task 11: Update check and apply from the release folder

**Files**
- Create: `engine\smartview_mcp\update.py`, `engine\tests\test_update.py`.

Release folder layout (written by `scripts\publish_release.ps1` in Task 15):

```
<release>\latest.json            {"version": "0.2.1", "published": "2026-11-02"}
<release>\0.2.1\                 unpacked `git archive` of tag v0.2.1 (has installer\install.ps1)
<release>\atlas\atlas.db         monthly atlas snapshot
```

- [ ] **Step 1: Write the failing test `engine\tests\test_update.py`**

```python
import json

from smartview_mcp import update


def _release(tmp_path, latest, versions):
    (tmp_path / "latest.json").write_text(json.dumps({"version": latest}), encoding="utf-8")
    for v in versions:
        (tmp_path / v / "installer").mkdir(parents=True)
        (tmp_path / v / "installer" / "install.ps1").write_text("", encoding="utf-8")
    return tmp_path


def test_parse_orders_numerically():
    assert update.parse("0.10.0") > update.parse("0.9.3")


def test_not_configured_and_missing_folder(tmp_path):
    assert update.check("", "0.2.0", "")["status"] == "not configured"
    assert update.check(str(tmp_path / "nope"), "0.2.0", "")["status"] == "error"


def test_latest_newer_is_available(tmp_path):
    r = update.check(str(_release(tmp_path, "0.2.1", ["0.2.1"])), "0.2.0", "")
    assert r["status"] == "ok" and r["update_available"] and r["target"] == "0.2.1"
    assert r["installer"].endswith(r"0.2.1\installer\install.ps1")
    assert not update.check(str(tmp_path), "0.2.1", "")["update_available"]
    assert not update.check(str(tmp_path), "0.3.0", "")["update_available"]   # never downgrade by itself


def test_pin_wins_also_downwards(tmp_path):
    rel = _release(tmp_path, "0.3.0", ["0.2.0", "0.3.0"])
    r = update.check(str(rel), "0.3.0", "0.2.0")
    assert r["update_available"] and r["target"] == "0.2.0"


def test_missing_target_folder_is_error(tmp_path):
    r = update.check(str(_release(tmp_path, "0.2.1", [])), "0.2.0", "")
    assert r["status"] == "error" and "0.2.1" in r["error"]


def test_dev_checkout_never_auto_updates(monkeypatch):
    monkeypatch.setattr(update, "is_dev_checkout", lambda: True)
    assert update.start_if_available() is False
```

Run it. Expected: FAIL.

- [ ] **Step 2: Write `engine\smartview_mcp\update.py`**

```python
"""Updates from the shared release folder (SharePoint, synced by OneDrive). The check is a file read; applying runs
the release's own installer. A dev checkout (repo with .git) never updates itself."""
from __future__ import annotations

import json
import subprocess
from pathlib import Path

from . import __version__, config

CREATE_NO_WINDOW = 0x08000000
DETACHED = 0x00000008


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


def start_if_available() -> bool:
    """Called by `smartass serve` at login: start the installer detached and let it restart serve."""
    if is_dev_checkout():
        return False
    r = check()
    if not r["update_available"]:
        return False
    subprocess.Popen(_cmd(r["installer"]), creationflags=DETACHED | CREATE_NO_WINDOW, close_fds=True)
    return True
```

- [ ] **Step 3: Run** `uv run pytest -q; uv run smartass doctor`. Expected: all tests pass, and `update` now shows
  `OK release folder not configured`.
- [ ] **Step 4: Commit** with the message `Update check/apply from release folder (pin supported, no auto-downgrade)`.

---

### Task 12: Per-user setup (`smartass setup`) and `installer\install.ps1`

**Files**
- Create: `engine\smartview_mcp\setup_user.py`, `engine\tests\test_setup_user.py`, `installer\install.ps1`.

The MCP server name stays `smartview`. Agents therefore keep seeing the same tool names (`smartview-sv_query` …),
and the skill text stays valid.

- [ ] **Step 1: Write the failing test `engine\tests\test_setup_user.py`**

```python
import json
import subprocess

from smartview_mcp import setup_user


def test_user_config_is_written_once_with_a_secret(tmp_path):
    p = tmp_path / "config.toml"
    assert setup_user.write_user_config(p) == "created"
    first = p.read_text(encoding="utf-8")
    assert "[http]" in first and "secret = " in first and "[packs]" in first
    assert setup_user.write_user_config(p) == "kept"
    assert p.read_text(encoding="utf-8") == first


def test_register_mcp_creates_updates_and_keeps_env(tmp_path):
    p = tmp_path / "mcp-config.json"
    p.write_text(json.dumps({"mcpServers": {"other": {"command": "x"},
                                            "smartview": {"command": "old", "env": {"K": "1"}}}}), encoding="utf-8")
    entry = {"type": "local", "command": "py.exe", "args": ["run_server.py"], "env": {}, "tools": ["*"]}
    assert setup_user.register_mcp(p, entry) == "updated"
    d = json.loads(p.read_text(encoding="utf-8"))
    assert d["mcpServers"]["other"] == {"command": "x"}
    assert d["mcpServers"]["smartview"]["command"] == "py.exe" and d["mcpServers"]["smartview"]["env"] == {"K": "1"}
    assert (tmp_path / "mcp-config.json.bak-smartass").exists()
    assert setup_user.register_mcp(p, entry) == "unchanged"
    assert setup_user.register_mcp(tmp_path / "absent.json", entry) == "missing"
    assert setup_user.register_mcp(tmp_path / "new.json", entry, create=True) == "updated"
    assert setup_user.register_mcp(tmp_path / "nodir" / "x.json", entry, create=True) == "missing"


def test_link_skill_replaces_dir_and_old_junction(tmp_path):
    src, dst, backup = tmp_path / "src", tmp_path / "skills" / "smartview", tmp_path / "backup"
    src.mkdir()
    (src / "SKILL.md").write_text("new", encoding="utf-8")
    dst.mkdir(parents=True)
    (dst / "SKILL.md").write_text("old", encoding="utf-8")
    assert setup_user.link_skill(src, dst, backup) == "linked (old copy moved to backup)"
    assert (dst / "SKILL.md").read_text(encoding="utf-8") == "new"
    assert any(backup.iterdir())
    assert setup_user.link_skill(src, dst, backup) == "unchanged"
    other = tmp_path / "other"
    other.mkdir()
    subprocess.run(["cmd", "/c", "rmdir", str(dst)], check=True)
    subprocess.run(["cmd", "/c", "mklink", "/J", str(dst), str(other)], check=True, capture_output=True)
    assert setup_user.link_skill(src, dst, backup) == "linked"
    assert (dst / "SKILL.md").exists()
```

Run it. Expected: FAIL.

- [ ] **Step 2: Write `engine\smartview_mcp\setup_user.py`**

```python
"""Per-user setup, run by installer\\install.ps1 (`smartass setup`). Idempotent: every step reports
created/updated/unchanged. Never needs admin rights."""
from __future__ import annotations

import json
import os
import secrets
import shutil
import subprocess
import time
from pathlib import Path

from . import config

SERVER_NAME = "smartview"
USER = Path(os.environ.get("USERPROFILE", str(Path.home())))
STARTUP = Path(os.environ.get("APPDATA", str(USER / "AppData" / "Roaming"))) / \
    "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"

CONFIG_TEMPLATE = """# smartass user settings. Defaults live in engine/defaults.toml - only override what you need.

[paths]
# analyser = 'C:\\Users\\you\\OneDrive - Telenor\\analyser'   # where answer workbooks go
# excel_exe = 'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE'

[packs]
enabled = []            # e.g. ["produktregnskap"]

[release]
# version = "0.2.0"     # pin a version (also downwards); empty = latest

[http]
secret = "{secret}"
"""


def write_user_config(path: Path = config.USER_CONFIG) -> str:
    if path.exists():
        return "kept"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(CONFIG_TEMPLATE.format(secret=secrets.token_urlsafe(32)), encoding="utf-8")
    return "created"


def mcp_entries(python: Path, relay: Path) -> dict:
    return {
        "copilot": {"type": "local", "command": str(python), "args": [str(relay)], "env": {}, "tools": ["*"]},
        "claude": {"type": "stdio", "command": str(python), "args": [str(relay)], "env": {}},
    }


def register_mcp(path: Path, entry: dict, create: bool = False) -> str:
    if not path.exists():
        if not (create and path.parent.exists()):
            return "missing"
        path.write_text("{}", encoding="utf-8")
    data = json.loads(path.read_text(encoding="utf-8") or "{}")
    servers = data.setdefault("mcpServers", {})
    old = servers.get(SERVER_NAME) or {}
    new = {**entry, "env": old.get("env") or entry.get("env", {})}
    if old == new:
        return "unchanged"
    bak = path.with_name(path.name + ".bak-smartass")
    if not bak.exists():
        shutil.copy2(path, bak)
    servers[SERVER_NAME] = new
    tmp = path.with_name(path.name + ".tmp-smartass")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, path)
    return "updated"


def _is_junction(p: Path) -> bool:
    try:
        return p.is_junction() or p.is_symlink()
    except OSError:
        return False


def link_skill(src: Path, dst: Path, backup_root: Path) -> str:
    """dst becomes a junction to src. A real folder is moved to backup_root (NOT next to the skills, or agents
    would load it as a second skill)."""
    if _is_junction(dst) and os.path.realpath(dst) == os.path.realpath(src):
        return "unchanged"
    note = ""
    if _is_junction(dst):
        os.rmdir(dst)
    elif dst.exists():
        backup_root.mkdir(parents=True, exist_ok=True)
        shutil.move(str(dst), str(backup_root / f"{dst.name}-{time.strftime('%Y%m%d-%H%M%S')}"))
        note = " (old copy moved to backup)"
    dst.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(["cmd", "/c", "mklink", "/J", str(dst), str(src)], check=True, capture_output=True)
    return "linked" + note


def make_shortcut(lnk: Path, target: str, args: str, workdir: str) -> None:
    import win32com.client
    sc = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(lnk))
    sc.TargetPath, sc.Arguments, sc.WorkingDirectory, sc.WindowStyle = target, args, workdir, 7
    sc.Save()


def startup_entries() -> list[str]:
    venv = config.ENGINE_DIR / ".venv" / "Scripts"
    logs = config.HOME / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    make_shortcut(STARTUP / "smartass-engine.lnk", str(venv / "pythonw.exe"), "-m smartview_mcp.cli serve",
                  str(config.ENGINE_DIR))
    make_shortcut(STARTUP / "smartass-sso-watcher.lnk", "powershell.exe",
                  f'-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "{config.ENGINE_DIR / "sso_watcher.ps1"}"'
                  f' -LogDir "{logs}"', str(config.ENGINE_DIR))
    out = ["startup: smartass-engine.lnk, smartass-sso-watcher.lnk"]
    legacy = STARTUP / "svkit-sso-watcher.lnk"
    if legacy.exists():
        legacy.unlink()
        out.append("startup: removed legacy svkit-sso-watcher.lnk")
    return out


def copy_atlas() -> str:
    if not config.RELEASE_FOLDER:
        return "atlas: no release folder configured"
    src = Path(config.RELEASE_FOLDER) / "atlas" / "atlas.db"
    if not src.exists():
        return f"atlas: none in {src.parent}"
    dst = config.ATLAS_DB
    if dst.exists() and dst.stat().st_mtime >= src.stat().st_mtime:
        return "atlas: local copy is current"
    tmp = dst.with_suffix(".db.tmp")
    shutil.copy2(src, tmp)
    os.replace(tmp, dst)
    return f"atlas: copied from {src}"


def run() -> list[str]:
    config.HOME.mkdir(parents=True, exist_ok=True)
    out = [f"config: {write_user_config()} ({config.USER_CONFIG})"]
    e = mcp_entries(config.ENGINE_DIR / ".venv" / "Scripts" / "python.exe", config.ENGINE_DIR / "run_server.py")
    out.append(f"copilot mcp: {register_mcp(USER / '.copilot' / 'mcp-config.json', e['copilot'], create=True)}")
    out.append(f"claude mcp: {register_mcp(USER / '.claude.json', e['claude'])}")
    skill_src = config.REPO_DIR / "skills" / "smartview"
    backup = config.HOME / "backup"
    out.append(f"skill (claude): {link_skill(skill_src, USER / '.claude' / 'skills' / 'smartview', backup)}")
    out.append(f"skill (agents): {link_skill(skill_src, USER / '.agents' / 'skills' / 'smartview', backup)}")
    out += startup_entries()
    out.append(copy_atlas())
    return out
```

On this machine, `~\.agents\skills\smartview` is a junction to `~\.claude\skills\smartview` (see the skills-layout
note). Pointing both junctions straight at the repo gives the same result, without a chain of junctions.

- [ ] **Step 3: Write `installer\install.ps1`**

```powershell
<#  smartass installer - per user, no admin, safe to re-run.
    From a release folder: copies the release to %LOCALAPPDATA%\smartass\app and installs there.
    From a git checkout (dev mode): installs in place and enables the pre-commit hook.  #>
param([switch]$NoDoctor)
$ErrorActionPreference = 'Stop'
$src = Split-Path $PSScriptRoot -Parent
$dev = Test-Path (Join-Path $src '.git')
if ($dev) { $app = $src } else {
    $app = Join-Path $env:LOCALAPPDATA 'smartass\app'
    Get-CimInstance Win32_Process -Filter "Name='pythonw.exe' or Name='python.exe'" |
        Where-Object { $_.CommandLine -like "*$app*" -and $_.CommandLine -like '*cli serve*' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
    robocopy $src $app /MIR /XD .venv __pycache__ .git /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy failed ($LASTEXITCODE)" }
}
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Invoke-RestMethod https://astral.sh/uv/install.ps1 | Invoke-Expression
    $env:Path = "$env:USERPROFILE\.local\bin;$env:Path"
}
$engine = Join-Path $app 'engine'
Push-Location $engine
try {
    if ($dev) { uv sync --frozen } else { uv sync --frozen --no-dev }
    if ($LASTEXITCODE) { throw "uv sync failed ($LASTEXITCODE)" }
    $py = Join-Path $engine '.venv\Scripts\python.exe'
    & $py -m smartview_mcp.cli setup
    if ($LASTEXITCODE) { throw "smartass setup failed ($LASTEXITCODE)" }
    if ($dev) { git -C $app config core.hooksPath .githooks }
    Start-Process (Join-Path $engine '.venv\Scripts\pythonw.exe') -ArgumentList '-m','smartview_mcp.cli','serve' `
        -WorkingDirectory $engine -WindowStyle Hidden
    if (-not $NoDoctor) { Start-Sleep 2; & $py -m smartview_mcp.cli doctor }
} finally { Pop-Location }
Write-Host "smartass installed ($(if ($dev) {'dev checkout'} else {$app})). Restart Copilot CLI / Claude Code to load it."
```

- [ ] **Step 4: Run**

`uv run pytest -q; uv run ruff check .`

Expected: all tests pass. **Do not run `install.ps1` on the owner's profile yet.** That happens in Task 16, after the
state migration.

- [ ] **Step 5: Commit** with the message `Per-user setup (config, MCP registration, skill junctions, Startup) and installer`.

---

### Task 13: Pre-commit guard against data and secrets

**Files**
- Create: `scripts\check_staged.py`, `.githooks\pre-commit`, `engine\tests\test_check_staged.py`.

- [ ] **Step 1: Write the failing test `engine\tests\test_check_staged.py`**

```python
import importlib.util
from pathlib import Path

_p = Path(__file__).resolve().parents[2] / "scripts" / "check_staged.py"
_spec = importlib.util.spec_from_file_location("check_staged", _p)
cs = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cs)


def test_blocks_databases_and_stray_workbooks():
    assert cs.problems("engine/atlas.db", b"") and cs.problems("x/drill.sqlite", b"")
    assert cs.problems("docs/analyse.xlsx", b"")
    assert not cs.problems("packs/produktregnskap/knowledge/source_x.xlsx", b"")
    assert not cs.problems("engine/smartview_mcp/knowledge/map.xlsx", b"")


def test_blocks_tokens_unless_marked():
    tok = "sso_token=" + "a" * 30
    assert cs.problems("a.py", tok.encode())
    assert cs.problems("a.py", ("x = 'eyJ" + "a" * 20 + "." + "b" * 20 + "'").encode())
    assert cs.problems("a.toml", ('secret = "' + "z" * 24 + '"').encode())
    assert not cs.problems("a.py", (tok + "  # allow-secret").encode())
    assert not cs.problems("a.toml", b'secret = "{secret}"')
```

Run it. Expected: FAIL (the file is not found).

- [ ] **Step 2: Write `scripts\check_staged.py`** (stdlib only, so the hook runs even without the venv)

```python
"""Blocks commits that contain data or secrets. Default: staged files (git index). --all: every tracked file (CI).
A line containing `allow-secret` is skipped (test fixtures)."""
from __future__ import annotations

import re
import subprocess
import sys

DB = re.compile(r"\.(db|sqlite3?|db-wal|db-shm)$", re.I)
WORKBOOK = re.compile(r"\.(xlsx|xlsm|xlsb|xls)$", re.I)
WORKBOOK_OK = ("knowledge/", "packs/")
TOKENS = [
    re.compile(r"sso_token=[A-Za-z0-9._~%-]{16,}"),
    re.compile(r"eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}"),
    re.compile(r"(?i)\b(secret|password|passwd|api_key|apikey|client_secret)\s*[=:]\s*['\"][^'\"{}\s]{16,}['\"]"),
]


def problems(path: str, data: bytes) -> list[str]:
    p = path.replace("\\", "/")
    if DB.search(p):
        return [f"{p}: database files never go in git"]
    if WORKBOOK.search(p):
        return [] if any(m in p for m in WORKBOOK_OK) else [f"{p}: workbooks only under knowledge/ or packs/"]
    out = []
    for n, line in enumerate(data.decode("utf-8", errors="ignore").splitlines(), 1):
        if "allow-secret" in line:
            continue
        if any(t.search(line) for t in TOKENS):
            out.append(f"{p}:{n}: looks like a token or secret")
    return out


def _git(*a: str) -> bytes:
    return subprocess.run(["git", *a], capture_output=True, check=True).stdout


def main(argv: list[str]) -> int:
    if "--all" in argv:
        files = _git("ls-files", "-z").decode().split("\0")
        read = lambda f: open(f, "rb").read()  # noqa: E731
    else:
        files = _git("diff", "--cached", "--name-only", "--diff-filter=ACMR", "-z").decode().split("\0")
        read = lambda f: _git("show", f":{f}")  # noqa: E731
    found = [m for f in files if f for m in problems(f, read(f))]
    for m in found:
        print("BLOCKED", m)
    if found:
        print("Remove the data/secret, or mark a deliberate test fixture line with `allow-secret`.")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

- [ ] **Step 3: Write `.githooks\pre-commit`** with LF line endings. `.gitattributes` from Task 1 must say
  `.githooks/* text eol=lf`; if it does not, add that line.

```sh
#!/bin/sh
py=engine/.venv/Scripts/python.exe
[ -x "$py" ] || py=python
exec "$py" scripts/check_staged.py
```

- [ ] **Step 4: Run**

```powershell
cd C:\dev\smartass\engine; uv run pytest tests\test_check_staged.py -q
cd ..; git config core.hooksPath .githooks; engine\.venv\Scripts\python.exe scripts\check_staged.py --all
```

Expected:
- 2 tests pass.
- `--all` prints nothing and exits 0. If it reports a tracked file, fix the file; do not loosen the rules.

- [ ] **Step 5: Prove that the hook blocks a commit**

```powershell
Set-Content -Path probe.db -Value x; git add probe.db; git commit -m probe; "exit=$LASTEXITCODE"
git reset -q probe.db; Remove-Item probe.db
```

Expected: `BLOCKED probe.db ...` and a non-zero exit code.

- [ ] **Step 6: Commit** with the message `Pre-commit guard: no databases, stray workbooks or tokens`.

---

### Task 14: CI on every push

**Files**
- Create: `.github\workflows\ci.yml`.

- [ ] **Step 1: Write `.github\workflows\ci.yml`**

```yaml
name: ci
on: [push, pull_request]
permissions:
  contents: read
jobs:
  test:
    runs-on: windows-latest
    defaults:
      run:
        working-directory: engine
    env:
      PYTHONUTF8: "1"
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.12"
      - run: python -m pip install uv
      - run: uv sync --frozen
      - run: uv run ruff check .
      - run: uv run pytest -q
        env:
          SMARTASS_HOME: ${{ runner.temp }}\smartass
      - run: uv run python ..\scripts\check_staged.py --all
```

Only GitHub-owned actions are used, because orgs often block third-party actions.

- [ ] **Step 2: Check locally what CI will see**

```powershell
cd C:\dev\smartass\engine; $env:SMARTASS_HOME = "$env:TEMP\smartass-ci"; uv run pytest -q; Remove-Item Env:SMARTASS_HOME
```

Expected: all tests pass with an empty home.

A test that needs a real Excel, Smart View or OneDrive fails here, or on the first CI run. Give it
`@pytest.mark.skipif(not shutil.which("excel") and not os.environ.get("SMARTASS_LIVE"), reason="needs desktop Excel")`.
Better still, make it independent of the machine with `tmp_path` or monkeypatch.

- [ ] **Step 3: Commit** with the message `CI: ruff, pytest, data/secret guard on windows-latest`.

The first CI run happens after the push in Task 16. If it is red, fix it before calling the task done.

---

### Task 15: README, troubleshooting, changelog and release script

**Files**
- Create: `README.md`, `docs\troubleshooting.md`, `docs\release.md`, `CHANGELOG.md`, `scripts\publish_release.ps1`,
  `tests\live\README.md` (that is, `engine\tests\live\README.md`).

- [ ] **Step 1: Write `README.md`.** Keep it to about one screen, with these sections:
  1. **What it is.** One paragraph: an agent works read-only against FinMax, FinLite and FinPlan through Smart View,
     and the answers are built in your own Excel.
  2. **Install (colleagues).** `powershell -ExecutionPolicy Bypass -File "<release folder>\<version>\installer\install.ps1"`.
     Then restart Copilot CLI or Claude Code, and run `smartass doctor`.
  3. **Data safety.** Read-only against EPM. Nothing leaves the PC except the EPM calls themselves. No data, databases
     or tokens in git (the pre-commit hook and CI enforce this). Answers are written only to your `analyser` folder.
  4. **Layout.** A copy of the file map from the top of this plan, shortened.
  5. **Develop.** `cd engine; uv sync; uv run pytest -q; uv run ruff check .`; the live checks are in
     `engine\tests\live\README.md`; the release steps are in `docs\release.md`.

- [ ] **Step 2: Write `docs\troubleshooting.md`.** Add one section per doctor check (`config`, `excel`, `smartview`,
  `epm`, `drill_browser`, `atlas`, `output`, `http_api`, `update`), each with:
  - the symptom, meaning the line `smartass doctor` prints;
  - the cause;
  - the fix, in the same words as the `fix` text in `doctor.py`.

  Then add two more sections:
  - "Smart View asks me to log in all the time": the SSO watcher, `logs\sso_watcher.log`, and how to restart it from
    the Startup entry.
  - "Still stuck": `smartass doctor --bundle`, then send the zip to the maintainer. The zip contains no tokens or
    passwords (they are redacted), but it does contain file paths.

- [ ] **Step 3: Write `scripts\publish_release.ps1`**

```powershell
<# Publish the committed HEAD as a release: <Folder>\<version>\ + latest.json (+ atlas snapshot), tag v<version>. #>
param([Parameter(Mandatory)][string]$Folder, [switch]$WithAtlas)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
if (git -C $repo status --porcelain) { throw 'working tree not clean - commit first' }
$ver = (Select-String -Path "$repo\engine\pyproject.toml" -Pattern '^version = "(.+)"').Matches[0].Groups[1].Value
$dest = Join-Path $Folder $ver
if (Test-Path $dest) { throw "release $ver already exists in $Folder - bump the version" }
$zip = Join-Path $env:TEMP "smartass-$ver.zip"
git -C $repo archive --format zip -o $zip HEAD
Expand-Archive $zip $dest; Remove-Item $zip
if ($WithAtlas) {
    New-Item -ItemType Directory -Force (Join-Path $Folder 'atlas') | Out-Null
    $py = "$repo\engine\.venv\Scripts\python.exe"
    & $py -c "import sqlite3,sys; s=sqlite3.connect(sys.argv[1]); d=sqlite3.connect(sys.argv[2]); s.backup(d); d.close(); s.close()" `
        "$env:LOCALAPPDATA\smartass\atlas.db" (Join-Path $Folder 'atlas\atlas.db')
}
@{ version = $ver; published = (Get-Date -Format 'yyyy-MM-dd') } | ConvertTo-Json |
    Set-Content -Encoding utf8 (Join-Path $Folder 'latest.json')
git -C $repo tag "v$ver"; git -C $repo push origin "v$ver"
"published $ver to $Folder"
```

`latest.json` is written last, so clients never see a half-copied release.

- [ ] **Step 4: Write `docs\release.md`** as a checklist:
  1. Bump `version` in `engine\pyproject.toml` and `__version__`.
  2. Add a CHANGELOG entry.
  3. Run `uv lock` and the tests.
  4. Run the live checks in `engine\tests\live\README.md` on this PC.
  5. Commit and push; CI must be green.
  6. Run `scripts\publish_release.ps1 -Folder <release> [-WithAtlas]`.
  7. Monthly: re-run with `-WithAtlas` after `sv_atlas_build`. This bumps the patch version.
  8. To roll back: users pin `release.version`, or you republish the older version under a new number.

- [ ] **Step 5: Write `CHANGELOG.md`** with a `## 0.2.0` entry: repo layout, `uv`, config layering, `doctor` and
  bundle, redacted rotating logs, error messages that name the next step, HTTP `/health`, the updater, the installer,
  the pre-commit guard, and CI.

- [ ] **Step 6: Write `engine\tests\live\README.md`.** It lists the release live checks: the five checks from Task 16
  Step 5, each with the exact tool call and its expected result.

- [ ] **Step 7: Commit** with the message `README, troubleshooting, release process and script`.

---

### Task 16: Migrate the owner, prove a fresh install, go live

The live MCP switches from `OLD` to the repo in this task. Until Step 3, nothing changes on the owner's setup.

- [ ] **Step 1: Copy the state from `%LOCALAPPDATA%\svkit` to `%LOCALAPPDATA%\smartass`.** Do this while no MCP tool
  call is running. Databases are copied with the SQLite backup API, so a live writer cannot leave a torn copy.

```powershell
$o = "$env:LOCALAPPDATA\svkit"; $n = "$env:LOCALAPPDATA\smartass"
New-Item -ItemType Directory -Force "$n\logs" | Out-Null
$py = 'C:\dev\smartass\engine\.venv\Scripts\python.exe'
foreach ($db in 'atlas.db','drill.db') {
  & $py -c "import sqlite3,sys; s=sqlite3.connect(sys.argv[1]); d=sqlite3.connect(sys.argv[2]); s.backup(d); d.close(); s.close()" "$o\$db" "$n\$db"
}
'sessions.json','smartview_options_original.json','closed_month.json','dims_cache.json' |
  ForEach-Object { if (Test-Path "$o\$_") { Copy-Item "$o\$_" "$n\" } }
robocopy "$o\drill_edge" "$n\drill_edge" /E /NFL /NDL /NJH /NJS /NP | Out-Null
if (Test-Path "$o\logs\queries.jsonl") { Copy-Item "$o\logs\queries.jsonl" "$n\logs\" }
& $py -c "import sqlite3,sys; [print(p, sqlite3.connect(p).execute('pragma integrity_check').fetchone()[0]) for p in sys.argv[1:]]" "$n\atlas.db" "$n\drill.db"
```

Expected: `ok` for both databases. The Edge profile keeps the FinLite login, so the first drill needs no new login.

- [ ] **Step 2: Fresh-user simulation (spec "Done when" #1).** Run it before touching the owner's profile, so a broken
  installer cannot break the live setup.

```powershell
$t = "$env:TEMP\smartass-fresh"; Remove-Item $t -Recurse -Force -ErrorAction SilentlyContinue
'local','roaming','profile\.copilot','onedrive','release' | ForEach-Object { New-Item -ItemType Directory -Force "$t\$_" | Out-Null }
'{}' | Set-Content "$t\profile\.copilot\mcp-config.json"
@{ version = '0.2.0' } | ConvertTo-Json | Set-Content "$t\release\latest.json"
git -C C:\dev\smartass archive --format zip -o "$t\rel.zip" HEAD; Expand-Archive "$t\rel.zip" "$t\release\0.2.0"
$env:UV_CACHE_DIR = (uv cache dir); $env:UV_PYTHON_INSTALL_DIR = (uv python dir)   # read BEFORE the overrides: reuse real cache/Python
$env:LOCALAPPDATA = "$t\local"; $env:APPDATA = "$t\roaming"; $env:USERPROFILE = "$t\profile"
$env:OneDriveCommercial = "$t\onedrive"; $env:SMARTASS_RELEASE = "$t\release"; $env:SMARTASS_HTTP_PORT = '47899'
powershell -NoProfile -ExecutionPolicy Bypass -File "$t\release\0.2.0\installer\install.ps1"
```

Check:
- `smartass doctor` prints `OK` for config, excel, smartview, epm, drill_browser, output, http_api and update.
- `atlas` is FAIL, because the fake release has no `atlas\atlas.db`. Copy `%LOCALAPPDATA%\smartass\atlas.db` (the real
  one) into `$t\release\atlas\`, re-run the installer, and `atlas` must turn OK. This proves the repair path too.
- `$t\profile\.copilot\mcp-config.json` has `mcpServers.smartview` pointing to `$t\local\smartass\app\engine`.
- `$t\profile\.claude\skills\smartview` is a junction.
- `$t\roaming\...\Startup` contains both `.lnk` files.
- Re-running the installer prints `kept` / `unchanged` everywhere.

If a check fails, fix the code (TDD: first a failing unit test that shows the bug), commit, then rebuild the snapshot
and re-run the simulation.

Clean up: stop the test `serve` process by PID, then delete `$t`.

```powershell
Get-CimInstance Win32_Process -Filter "Name='pythonw.exe'" | Where-Object CommandLine -like "*smartass-fresh*" | ForEach-Object { Stop-Process -Id $_.ProcessId }
```

Run all of Step 2 in one separate PowerShell process, so the environment overrides die with it.

- [ ] **Step 3: Install for the owner in dev mode, then write the owner's settings**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File C:\dev\smartass\installer\install.ps1 -NoDoctor
```

Then add this to `%LOCALAPPDATA%\smartass\config.toml` with the edit tool, keeping the generated secret:

```toml
[paths]
analyser = 'C:\Users\t992115\OneDrive - Telenor\Projects\prodreg2026\analyser'

[packs]
enabled = ["produktregnskap"]
```

If `[paths]` and `[packs]` already exist from the template, put the values in those sections; do not add the headers
a second time. The other settings keep their defaults; these two are the only owner values that today are hard-coded
in `OLD`.

Run `C:\dev\smartass\engine\.venv\Scripts\smartass.exe doctor`. Expected: all OK except `update`, which shows OK
"not configured" until the release folder exists.

- [ ] **Step 4: Ask the owner to restart Copilot CLI** (and Claude Code if it is open), so both pick up the new MCP
  entry. Use `ask_user` and wait for the answer. After the restart, `sv_status` must show the engine running from
  `C:\dev\smartass\engine` and the doctor block all OK.

- [ ] **Step 5: Live checks (spec "Done when" #2)**
  1. `sv_query` on each of FinMax, FinLite and FinPlan: a small grid, numbers returned.
  2. `sv_ws_build`: one small analysis, written to the owner's `analyser` folder.
  3. `sv_drill` on ac.12881 × un.SY × Sep 26 in FinLite: **95 lines, sum equal to the cell, difference exactly 0.**
  4. Relay reload: touch `engine\smartview_mcp\config.py` (save it without changes), make one tool call, and check
     `%LOCALAPPDATA%\smartass\logs\proxy.log` for one restart and no error.
  5. `smartass doctor --bundle`: open the zip and check that it contains no `sso_token=` value and no HTTP secret.

- [ ] **Step 6: Retire `OLD`**

```powershell
$old = "C:\Users\t992115\OneDrive - Telenor\Projects\prodreg2026\smartview"
"Moved to C:\dev\smartass on $(Get-Date -f yyyy-MM-dd). Delete this folder after $((Get-Date).AddDays(14).ToString('yyyy-MM-dd'))." |
  Set-Content "$old\MOVED.txt"
attrib +R "$old\*" /S /D
```

Leave `%LOCALAPPDATA%\svkit` as it is for the same two weeks. It is the rollback: point the MCP entries back to
`OLD`'s python and `run_server.py`.

- [ ] **Step 7: Push (when IT has created the repo, D4)**

```powershell
cd C:\dev\smartass; git remote add origin https://github.com/TelenorNorgeInternal/s<serviceID>-smartass.git; git push -u origin main
```

Check CI with `gh run watch` or in the Actions tab; it must be green (spec "Done when" #3). If the repo does not exist
yet, mark this step blocked in the todo list and finish Task 17; the push is the only step that waits.

---

### Task 17: Audit

- [ ] **Step 1: Run the full suite from clean**

```powershell
cd C:\dev\smartass\engine; uv sync --frozen; uv run ruff check .; uv run pytest -q
python ..\scripts\check_staged.py --all
```

- [ ] **Step 2: Repeat the "Done when" #4 scan by hand,** in addition to `test_paths.py`:

```powershell
Select-String -Path C:\dev\smartass\engine\*,C:\dev\smartass\engine\smartview_mcp\*,C:\dev\smartass\installer\*,C:\dev\smartass\skills\* -Recurse -Pattern 't992115','prodreg2026','tom_chamkrai' -SimpleMatch
```

Expected: no hits outside `engine\tests\live\` and `packs\`.

- [ ] **Step 3: Code review.** Use superpowers:requesting-code-review on the range from `bebde5c` to `HEAD`. Fix every
  Critical and Important finding, re-run Step 1, and commit.
- [ ] **Step 4: Report** with superpowers:verification-before-completion. Give each spec "Done when" item, the
  command or check that proves it, and the actual output. Mark open items as open (push and CI if the repo is not
  there yet, and the release folder path).

---

## Spec coverage (self-review)

| Spec item | Task |
|---|---|
| D1 owner first, colleagues later without redesign | 4 (config layering), 11, 12 |
| D2 own login, read-only | unchanged engine; 15 (README data safety) |
| D3 approved models only | no model calls in this sub-project; README states the rule (15) |
| D4 repo, local git first, push later | 1, 16 step 7 |
| D5 no databases, logs or outputs in git | 1 (`.gitignore`), 13, 14 |
| D6 monorepo at `C:\dev\smartass` | 1 |
| D7 package and tool names unchanged | 7 (tool list identical), 12 (server name `smartview`) |
| D8 two-level config + env vars | 4 |
| D9 uv + lockfile, Python 3.12 | 2, 12 |
| D10 GitHub for source, release folder for colleagues | 11, 12, 15 |
| D11 CI unit tests, live tests manual | 14, 15 (`tests\live\README.md`), 16 step 5 |
| D12 localhost HTTP `/health` with secret | 10, 12 (secret, Startup entry) |
| D13 English | all |
| Config: auto-detect Excel, Smart View DLL; default output in OneDrive | 4 |
| Installer steps 1-8 | 12 |
| Doctor checks + `--bundle`, in `sv_status` and `/health` | 9, 10 |
| Update from `latest.json`, rollback by pin | 11 (timing: flagged decision 4) |
| Owner migration steps 1-5 | 1, 16 |
| Split `server.py` | 7 |
| Declare playwright, drop uiautomation | 2 (`drill_chrome` exists only as a test assertion, nothing to remove) |
| Errors state what, why, next step | 8 |
| Rotating, redacted logs | 6 |
| Pre-commit hook rules | 13 |
| CI: tests, ruff, stdio smoke test | 3, 14 |
| Live checks before release | 15, 16 step 5 |
| Done when #1 fresh user | 16 step 2 |
| Done when #2 drill 12881 reconciles | 16 step 5 |
| Done when #3 CI green on first push | 14, 16 step 7 |
| Done when #4 no user paths in code | 5, 17 step 2 |
