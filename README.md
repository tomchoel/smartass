# smartass

**What it is.** An engine (MCP server) that lets an AI agent (Copilot CLI, Claude Code) work read-only against
Telenor Norge's Oracle EPM cubes FinMax, FinLite and FinPlan through Smart View. Answers are built as live
analyses in your own Excel: Smart View data sheets plus a formatted answer sheet.

## Install (colleagues)

```powershell
powershell -ExecutionPolicy Bypass -File "<release folder>\<version>\installer\install.ps1"
```

Then restart Copilot CLI or Claude Code, open a new terminal (the installer adds `smartass` to your user PATH), and
run `smartass doctor`. If `smartass` is not found, use the full path
`%LOCALAPPDATA%\smartass\app\engine\.venv\Scripts\smartass.exe`. Per user, no admin, safe to re-run.
Problems: [docs\troubleshooting.md](docs/troubleshooting.md).

## Data safety

- **Read-only against EPM.** The engine only retrieves data; it never submits data to the cubes.
- **Your own login.** Every user connects with their own Smart View / Telenor SSO login and sees only what
  they are allowed to see.
- **Nothing leaves the PC** except the calls to EPM themselves. The HTTP API listens on `127.0.0.1` only and
  needs a per-user secret.
- **Your own workbooks are never saved** by the engine. Changes it makes in a workbook you opened are left unsaved
  for you to review (`xl_revert` undoes them); a workbook refresh saves a copy, never the source. Only the
  engine's own answer workbooks are saved, in your `analyser` folder (`paths.analyser`).
- **The Produktregnskapet archive is protected:** writes there are refused unless you explicitly ask for it.
- **No data, databases or tokens in git.** The pre-commit hook and CI enforce this (`scripts\check_staged.py`);
  logs and `smartass doctor --bundle` redact tokens and secrets.

## Layout

```
.github\workflows\ci.yml     CI: ruff, pytest, data/secret guard
.githooks\pre-commit         runs scripts\check_staged.py
docs\                        DESIGN, troubleshooting, release, specs, plans
installer\install.ps1        per-user installer
scripts\                     check_staged.py, publish_release.ps1
knowledge\  packs\           business maps; pack-only inputs (packs\produktregnskap)
skills\smartass\             master copy of the agent skill
engine\                      Python package smartview_mcp, defaults.toml, tests\ (tests\live = manual)
```

## Develop

```powershell
cd engine; uv sync; uv run pytest -q; uv run ruff check .
```

Live checks against EPM: [engine\tests\live\README.md](engine/tests/live/README.md).
Releasing: [docs\release.md](docs/release.md).
