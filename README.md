# smartass

**What it is.** An engine (MCP server) that lets an AI agent (Copilot CLI, Claude Code) work read-only against
Telenor Norge's Oracle EPM cubes FinMax, FinLite and FinPlan through Smart View. Every finance question runs through
one pipeline: `sv_gate` checks the question's scope (and asks when it is ambiguous), then `sv_answer` pulls the data
(over Smart View's HTTP protocol, Excel as the fallback) and builds the answer in your own Excel: a formatted
`Q<n> Answer` sheet with live formulas, one source sheet per pull, checks, review findings and the lines that need
commentary. Follow-ups change the same answer; recipes (`pnl`, `actual_vs_plan`, `monthly_trend`, `cost_base`,
`concept`) draft the common analyses.

## Tools (25)

| Group | Tools |
|---|---|
| Pipeline | `sv_gate`, `sv_remember`, `sv_answer` |
| Health | `sv_status`, `sv_restart` |
| Atlas | `sv_atlas_build`, `sv_atlas_status`, `sv_atlas_changes` |
| Members | `sv_dimensions`, `sv_find_member`, `sv_members`, `sv_member_info` |
| Drill (GL lines) | `sv_drill` |
| Workbench (any open workbook) | `xl_workbooks`, `xl_overview`, `xl_read`, `xl_write`, `xl_rows`, `xl_add_sheet`, `xl_format`, `xl_format_sheet`, `xl_trace`, `xl_check`, `xl_changes`, `xl_revert` |

How the agent uses them: [skills\smartass\SKILL.md](skills/smartass/SKILL.md). Architecture: [docs\DESIGN.md](docs/DESIGN.md).

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
  for you to review (`xl_revert` undoes them). Only the engine's own answer workbooks are saved, in your
  `analyser` folder (`paths.analyser`). The engine never closes or restarts your Excel.
- **The Produktregnskapet archive is protected:** writes there are refused unless you explicitly ask for it.
- **No data, databases or tokens in git.** The pre-commit hook and CI enforce this (`scripts\check_staged.py`);
  logs and `smartass doctor --bundle` redact tokens and secrets.

## HTTP API (`smartass serve`)

Started at login; listens on `127.0.0.1` only (`http.port`), every request with the header `X-Smartass-Secret`
(`http.secret` in `%LOCALAPPDATA%\smartass\config.toml`). One request at a time (a second gets 409).

| Endpoint | Body | Returns |
|---|---|---|
| `GET /health` | – | the `smartass doctor` checks |
| `POST /gate` | `{"question"}` | what `sv_gate` returns |
| `POST /answer` | `{"spec" \| "followup", "q", "target", "allow_archive"}` | what `sv_answer` returns |

Errors: 400 (bad JSON, unknown fields, invalid input with its message), 401 (secret), 413 (body over 1 MB),
503 (Excel/EPM error with its message).

## Layout

```
.github\workflows\ci.yml     CI: ruff, pytest, data/secret guard
.githooks\pre-commit         runs scripts\check_staged.py
docs\                        DESIGN, troubleshooting, release, specs, plans, audit (docs\archive: before 2 Nov 2026)
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
