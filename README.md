# smartass

**What it is.** An engine (MCP server `smartass`) that lets an AI agent (Claude Code, Copilot CLI) work read-only
against Telenor Norge's Oracle EPM cubes FinMax, FinLite and FinPlan. The agent writes an **AnalysisSpec v2**
(plan, queries, typed blocks); `fin_analyze` fetches the data (Smart View's HTTP protocol, Oracle's Planning REST
as the second engine), computes every value in Python and draws the answer in Excel: a formatted `Q<n> Answer` sheet
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

Needs git, [uv](https://docs.astral.sh/uv/getting-started/installation/) and desktop Excel. Per user, no admin rights,
safe to re-run:

```powershell
git clone https://github.com/tomchoel/smartass.git; cd smartass\engine
uv sync --frozen
uv run --frozen smartass setup
# then add [oauth] to %LOCALAPPDATA%\smartass\config.toml (see below) before:
uv run --frozen smartass login
```

`smartass setup` writes your `%LOCALAPPDATA%\smartass\config.toml`, registers the MCP server for Claude Code and
Copilot CLI, links the skill and puts `smartass` on your user PATH. EPM access needs `[oauth]` in that config.toml
(identity domain, client id, a scope per pod; docs\troubleshooting.md, "oauth") before `smartass login`, which signs
you in (one approval in your browser; no browser is installed or started by smartass) and builds the member atlas
when there is none (about half a minute, then the behaviour cards; `--no-atlas` skips it, `smartass atlas build`
builds it later). Then restart Copilot CLI or Claude Code, open a new terminal and run `smartass doctor`.
Update: `git pull`, then `uv sync --frozen` in `engine` (and `smartass setup` when the CHANGELOG says so).
Working on smartass itself: `git config core.hooksPath .githooks` turns on the pre-commit data and secret guard.
Problems: [docs\troubleshooting.md](docs/troubleshooting.md).

## Data safety

- **Read-only against EPM.** The engine only retrieves data; it never submits data to the cubes.
- **Your own login.** Every user signs in to EPM once with their own account (`smartass login`, approved with Telenor
  SSO in their browser) and sees only what they are allowed to see. Every request to EPM with your sign-in passes one
  allowlist of read requests.
- **Nothing leaves the PC** except the calls to EPM themselves. smartass opens no port: the agent starts the MCP
  server and talks to it over stdio.
- **Your own workbooks are never saved** by the engine. Changes made in a workbook you opened are left unsaved for you
  to review (`xl_snapshot` before an edit, `xl_revert` undoes it). Only the engine's own answer workbooks are saved, in
  your `analyser` folder (`paths.analyser`). The engine never restarts your Excel, and closes one only when it started
  that Excel itself for a new workbook that could not be saved. A server that does not stop within a minute has its
  own processes ended, never an Excel. Open item: anything that ends the server's processes from outside also ends an
  Excel smartass started for it.
- **The Produktregnskapet archive is protected:** answers are never written there.
- **No data, databases or tokens in git.** The pre-commit hook and CI enforce this (`scripts\check_staged.py`);
  logs and `smartass doctor --bundle` redact tokens and secrets.

## Command line

`smartass doctor [--json | --bundle | --fix]`, `smartass atlas status`,
`smartass atlas build [--cube X] [--dimension D]` (over Planning REST, no Excel), `smartass setup`,
`smartass login [--status | --force | --forget | --no-atlas]` (EPM sign-in for the bearer route; builds the atlas on a
PC that has none; see
`docs\troubleshooting.md`, "oauth").

## Layout

```
.github\workflows\ci.yml     CI: ruff, pytest, data/secret guard
.githooks\pre-commit         runs scripts\check_staged.py
docs\                        DESIGN, troubleshooting, release, specs, plans (docs\archive: earlier designs)
scripts\                     check_staged.py, bump.ps1
knowledge\  packs\           business maps and metric catalogue; pack-only inputs (packs\produktregnskap)
skills\smartass\             master copy of the agent skill
engine\                      run_server.py (relay), child_server.py, VERSION, defaults.toml, tests\
engine\smartass\             server.py, pipeline.py, exceljob.py (the Excel job), followup.py, workbench.py and the packages:
  core\       config, logs, timing, session, serverkit (tool guard, Excel jobs, lifespan)
  model\      spec, periods, expressions, blocks, findings, validation (no Excel, no network)
  data\       providers: Smart View HTTP, Planning REST (second engine), recorder
  render\     data table, answer sheet, charts, style
  epm\        the HTTPS boundary, sign-in, Smart View HTTP/XML, Planning REST, grids, the drill over the token
  knowledge\  atlas, resolver, interim metric catalogue
  xl\         jobs in the user's Excel, sheets, styling, formula check, snapshots
  ops\        cli, doctor (and an older release's leftovers), setup, login, atlas build
  testset\    the B-question test set (golden numbers outside git)
```

## Develop

```powershell
cd engine; uv sync; uv run pytest -q; uv run ruff check .
```

Test set against EPM (golden numbers outside git): `uv run python -m smartass.testset run`.
Rolling a code change into running servers: `scripts\bump.ps1` (bumps `engine\VERSION`) or
`fin_status(restart=True)`. Releasing: [docs\release.md](docs/release.md).
