# smartass

**What it is.** An engine (MCP server `smartass`) that lets an AI agent (Claude Code, Copilot CLI) work read-only
against Telenor Norge's Oracle EPM cubes FinMax, FinLite and FinPlan. The agent writes an **AnalysisSpec v2**
(plan, queries, typed blocks); `fin_analyze` fetches the data (Smart View's HTTP protocol, an ad hoc grid in Excel
as the fallback), computes every value in Python and draws the answer in Excel: a formatted `Q<n> Answer` sheet
whose numbers are live formulas on one data table per question (`Data_Q<n>`), with checks, findings, the material
lines and their drivers. Follow-ups patch the stored spec; monthly packs refresh to the latest closed month.

## Tools (9)

| Group | Tools |
|---|---|
| Analysis | `fin_analyze` (new spec or follow-up patch), `fin_refresh` (re-run every answer in a workbook) |
| Names | `fin_lookup` (metrics, atlas members and, with `kind="answer"`, past answers) |
| Critic | `fin_critic` (the brief and rubric for an answer; stores the score). Playbooks are MCP prompts |
| Ledger lines | `fin_drill` (FinLite lines behind one cell, ranked and tied to the cube) |
| Health | `fin_status` (version, atlas, doctor; `restart=True` for a fresh server) |
| Open workbooks (with Excel MCP) | `xl_snapshot`, `xl_revert`, `xl_tidy` (TN style + formula check) |

How the agent uses them: [skills\smartass\SKILL.md](skills/smartass/SKILL.md). Architecture: [docs\DESIGN.md](docs/DESIGN.md).
v1 (the `sv_*` tools and the `xl_*` workbench) was retired in November 2026; the rename table is in
[CHANGELOG.md](CHANGELOG.md).

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
  needs a per-user secret (DPAPI-protected).
- **Your own workbooks are never saved** by the engine. Changes made in a workbook you opened are left unsaved
  for you to review (`xl_snapshot` before an edit, `xl_revert` undoes it). Only the engine's own answer workbooks
  are saved, in your `analyser` folder (`paths.analyser`). The engine never closes or restarts your Excel.
- **The Produktregnskapet archive is protected:** answers are never written there.
- **No data, databases or tokens in git.** The pre-commit hook and CI enforce this (`scripts\check_staged.py`);
  logs and `smartass doctor --bundle` redact tokens and secrets.

## HTTP API (`smartass serve`)

Started at login; listens on `127.0.0.1` only (`http.port`), every request with the header `X-Smartass-Secret`
(stored DPAPI-protected in `%LOCALAPPDATA%\smartass\secret.bin` by `smartass setup`). One request at a time (a
second gets 409).

| Endpoint | Body | Returns |
|---|---|---|
| `GET /health` | – | the `smartass doctor` checks |
| `POST /analyze` | `{"spec": <AnalysisSpec v2>, "target": "session"}` | what `fin_analyze(spec=...)` returns |

Errors: 400 (bad JSON, unknown fields, invalid spec or input, with its message), 401 (secret), 413 (body over
1 MB), 503 (Excel/EPM error with its message).

## Command line

`smartass doctor [--json | --bundle]`, `smartass atlas status`, `smartass atlas build [--cube X] [--dimension D]`,
`smartass setup`, `smartass update`, `smartass serve`, `smartass login [--status | --force | --forget]` (EPM sign-in for
the bearer route, see `docs\troubleshooting.md`, "oauth").

## Layout

```
.github\workflows\ci.yml     CI: ruff, pytest, data/secret guard
.githooks\pre-commit         runs scripts\check_staged.py
docs\                        DESIGN, troubleshooting, release, specs, plans (docs\archive: earlier designs)
installer\install.ps1        per-user installer
scripts\                     check_staged.py, bump.ps1, publish_release.ps1, probe_sv_http.py
knowledge\  packs\           business maps and metric catalogue; pack-only inputs (packs\produktregnskap)
skills\smartass\             master copy of the agent skill
engine\                      run_server.py (relay), child_server.py, VERSION, defaults.toml, vba\, tests\
engine\smartass\             server.py, pipeline.py, followup.py, workbench.py and the packages:
  core\       config, logs, timing, session, serverkit (tool guard, Excel jobs, lifespan)
  model\      spec, periods, expressions, blocks, findings, validation (no Excel, no network)
  data\       providers: Smart View HTTP, ad hoc grid fallback, recorder
  render\     data table, answer sheet, charts, style
  epm\        Smart View HTTP/XML, grids, sign-in, drill browser and lines
  knowledge\  atlas, resolver, interim metric catalogue
  xl\         Excel host (hidden + user Excel), sheets, styling, formula check, snapshots
  ops\        cli, doctor, setup, update, httpapi, atlas build
  testset\    the B-question test set (golden numbers outside git)
```

## Develop

```powershell
cd engine; uv sync; uv run pytest -q; uv run ruff check .
```

Test set against EPM (golden numbers outside git): `uv run python -m smartass.testset run`.
Rolling a code change into running servers: `scripts\bump.ps1` (bumps `engine\VERSION`) or
`fin_status(restart=True)`. Releasing: [docs\release.md](docs/release.md).
