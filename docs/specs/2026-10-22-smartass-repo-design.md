# smartass — sub-project 1: repo + hardened engine (design)

Date: 2026-10-22 · Owner: Tom Christian Hoel (tomchoel) · Status: approved in brainstorming, awaiting spec review

## Context

The smartview MCP (`prodreg2026\smartview\mcp`, ~7 900 LOC Python, package `smartview_mcp`) lets AI agents answer
finance questions from Oracle EPM Cloud (FinMax, FinLite, FinPlan) through the Smart View Excel add-in (COM, per-user
SSO), and writes live, reconciled analysis sheets. It also drills FinLite cells to ledger lines through a headless
Edge. Today it runs only on the owner's PC: no git, no README, no installer, hardcoded user paths, Playwright not
declared, and a 132 MB member atlas built locally.

Goal of the product ("smartass"): the best tool in Telenor Norge for Smart View and the finance structure, first for
the owner, later for colleagues in the finance division. Not overengineered: one maintainer, part-time.

## Product roadmap (three sub-projects, each with its own spec)

1. **Repo + hardened engine** (this spec).
2. **Excel add-in for colleagues**: Office.js task pane + right-click, AI chat from day one. Colleagues have only
   Microsoft 365 Copilot, which cannot reach a local COM engine, so the add-in is our own. The model is Azure OpenAI
   in Telenor's tenant (EU data zone, via IT). The AI tool loop runs inside the local engine; the pane is a thin UI.
   Deployed with M365 Centralized Deployment. UI designed with Impeccable (impeccable.style).
3. **Features**, in this order: close pack, "drill this cell", workbook audit → KPI/FinPlan → variance commentary,
   template library, accounting Q&A → Produktregnskap/CostPerform pack.

## Decisions (sub-project 1)

| # | Decision |
|---|---|
| D1 | First user is the owner only; built so colleagues can be added without redesign. |
| D2 | Each user queries EPM with their own Smart View login; read-only; no service account. |
| D3 | EPM numbers may go to Telenor-approved enterprise models only. |
| D4 | Repo: `TelenorNorgeInternal/s<serviceID>-smartass`, private; IT creates it (service ID). Work starts in a local git repo now and is pushed when it exists. |
| D5 | Repo holds code, skills and knowledge maps (JSON + review XLSX). Never `atlas.db`, `drill.db`, logs, outputs or analyses. |
| D6 | Monorepo layout (below). Working copy at `C:\dev\smartass`, outside OneDrive. |
| D7 | Package name `smartview_mcp` and tool names `sv_*` / `xl_*` stay unchanged. |
| D8 | Two-level config: org defaults in the repo, per-user overrides in `%LOCALAPPDATA%\smartass\config.toml`; env vars override both. |
| D9 | uv + lockfile is the only install mechanism (Python 3.12 per user, no admin). |
| D10 | Source and CI on GitHub; released builds, installer and the monthly prebuilt atlas in a SharePoint/Teams folder, so colleagues need no GitHub account. |
| D11 | CI: GitHub Actions on `windows-latest`, unit tests only. Live tests stay manual on the owner's PC. |
| D12 | Engine keeps stdio MCP and adds a localhost-only HTTP API (per-install secret); in this sub-project it serves `/health` only. |
| D13 | Everything in English (code, docs, UI). |

## 1. Architecture and layout

```
C:\dev\smartass\
├─ engine\                 today's mcp\: pyproject.toml + uv.lock, smartview_mcp\, vba\, tests\
├─ skills\smartview\       master copy of the skill; ~\.agents\skills and ~\.claude\skills link here
├─ knowledge\              core maps: accounts, org, FinPlan, concepts (JSON + review XLSX)
├─ packs\produktregnskap\  products map, Produktregnskap concepts, archive rules; off by default
├─ installer\install.ps1   per-user, no admin, idempotent
├─ addin\                  placeholder for sub-project 2
├─ docs\                   README, troubleshooting, DESIGN, specs, plans
└─ .github\workflows\ci.yml
```

Runtime: one engine per PC, two front doors:

- **stdio MCP**: Copilot CLI and Claude, exactly as today (via the existing reload relay `run_server.py`).
- **HTTP API on 127.0.0.1**: requires the per-install secret from `config.toml`; `/health` returns the doctor result.
  It is started by a per-user login entry (Startup folder shortcut, no admin).

State outside git: `%LOCALAPPDATA%\smartass\` (atlas.db, drill.db, logs, config.toml, drill browser profile). The
analysis output folder stays on OneDrive.

## 2. Config, install, doctor, update

**Config.** `config.py` reads, in order: `engine\defaults.toml` (EPM URL, connections, drill channels, sensitivity
label IDs) → `%LOCALAPPDATA%\smartass\config.toml` (output folder, archive root, enabled packs, release folder,
HTTP port and secret) → environment variables (existing `SV_*` / `SVKIT_*` names keep working). Excel, the Smart
View DLL and Edge are auto-detected (registry, then standard paths). After this, no `t992115` or `prodreg2026`
path remains in code; the owner's paths live in his `config.toml`.

Default output folder for a new user: `%OneDrive%\smartass\analyser` (OneDrive for Business root from the
`OneDriveCommercial` environment variable).

**`installer\install.ps1`** (one line to run, no admin, safe to re-run = repair):

1. Install uv into the user profile if missing.
2. `uv sync --frozen` in `engine\` (Python 3.12 + locked dependencies incl. Playwright; no browser download).
3. Copy the prebuilt atlas from the release folder if the local one is missing or older.
4. Write `config.toml` if missing (random HTTP secret; never overwrite user values).
5. Register the MCP server in `~\.copilot\mcp-config.json` and `~\.claude.json` when those files exist.
6. Create skill junctions for `skills\*`.
7. Create the login entry for the HTTP engine.
8. Run `smartass doctor` and print the result.

**`smartass doctor`** (CLI, also part of `sv_status` and `/health`). One line per check, OK/WARN/FAIL plus the fix:
Excel found · Smart View installed + version · EPM SSO reachable · drill browser usable (UserDataDir policy check) ·
atlas present + age · output folder writable · config valid · HTTP API listening. Cheap: starts neither Excel nor a
browser. `smartass doctor --bundle` zips logs, redacted config and versions into one file to send to the maintainer.

**Update.** At start-up the engine compares its version with `latest.json` in the release folder. When newer, it
updates once no tool call is running (`uv sync` against the new release, then the relay reloads the child). Rollback =
set `release_version` in `config.toml`.

**Migration (owner).**
1. Copy `prodreg2026\smartview` into the new layout at `C:\dev\smartass`.
2. `git init` and the first commit, with the exclusions from D5.
3. Repoint the MCP configs and skill junctions.
4. Run the full unit tests and one live drill.
5. Keep the old folder read-only for two weeks, then delete it.

## 3. Hardening, errors, testing

**Targeted code changes** (no rewrites):
- Split `server.py` (~1 300 LOC) into `tools_query`, `tools_ws`, `tools_drill`, `tools_xl`, `tools_status`; tool
  names and behaviour unchanged.
- Declare `playwright` in `pyproject.toml`.
- Remove dead pieces: `drill_chrome` leftovers; drop `uiautomation` if no longer used.
- Every `ToolError` states what failed, why, and the next step.

**Logs.** `%LOCALAPPDATA%\smartass\logs`, rotating, with `sso_token` and the HTTP secret redacted.

**Data safety rules** (README + code):
- Read-only towards EPM.
- Each user's own login.
- The user's own workbooks are never saved.
- The archive is protected unless explicitly allowed.
- A pre-commit hook blocks `*.db`, `*.xlsx` outside `knowledge\` and `packs\`, and token-like strings.

**Testing.**
- **CI (`windows-latest`, every push):** the unit tests (139 today), `ruff`, and a smoke test that the stdio MCP
  server starts and lists its tools with Excel mocked.
- **`engine\tests\live\` (manual, owner's PC, before each release):** a query each to FinMax, FinLite and FinPlan;
  one `sv_ws_build`; `sv_drill` on ac.12881 × un.SY × Sep 26 (must reconcile exactly, 95 lines); a relay reload.
- **Release checklist:**
  1. CI green and live tests pass.
  2. Bump the version and update `CHANGELOG.md`.
  3. Build and publish to the release folder with `latest.json`.

## Done when

1. A fresh Windows user runs the one-line install and `smartass doctor` is all OK.
2. `sv_drill` on 12881 Sep 26 reconciles exactly from the new location.
3. CI is green on the first push.
4. No user-specific path or name remains in code.

## Out of scope (later sub-projects)

The Excel add-in UI, the Azure OpenAI loop, the HTTP API beyond `/health`, the new features, the Intune packaging.

## Open items owned by the user

- A service ID and repo creation in TelenorNorgeInternal (IT).
- The SharePoint/Teams release folder (owner creates it; the path goes into `defaults.toml`).
- An Azure OpenAI endpoint request to IT (needed for sub-project 2, so ask early).
