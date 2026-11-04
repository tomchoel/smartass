# Changelog

## Unreleased

One pipeline: every question goes `sv_gate` -> `sv_answer` (design `docs/specs/2026-11-02-pipeline-architecture-design.md`).

### Added

- v2 model core (`engine/smartass/model`): typed analysis spec, period algebra, fact frame, expression trees
  (value and Excel formula from one tree), seven computed blocks (table, kpi_tiles, trend, matrix, bridge, ranking,
  drivers), data guards, material lines and key values. Not wired to tools yet (Plan 1B).
- v2 first answer (`smartass2` server, beside v1): `fin_analyze`, `fin_lookup`, `fin_status`; Smart View HTTP
  provider, interim atlas resolver, Q<n> Answer as live formulas on table Data_Q<n>, "Sheet = model" readback;
  answer files in `<analyser>\v2`. Not yet: follow-ups, packs, drill, xl_* tools, ad hoc fallback (Plan 1C).
- `sv_gate`: the advisor gate. A fixed checklist (period, scenario, entity, measure, currency/sign, granularity)
  against the semantic catalogue, the capability map and the user's learned defaults; `clear`, `ambiguous` (1-3
  questions) or `infeasible` (reason + proxy). It never pulls. `sv_remember` saves a gate answer as a default.
- `sv_answer`: one question from an answer spec v1 (`core/contract.py`): pulls over Smart View's HTTP provider
  (Smart View in Excel as the fallback), `Q<n> Answer` shown first, live formulas, checks, a findings table
  (Critical / Warning / Info), material lines and commentary drivers (data with evidence, or a hypothesis).
  - Follow-ups patch the stored spec (`followup=`): unchanged pulls are reused, the answer is redrawn inside the
    tool-owned range, a hidden backup restores it if the redraw fails.
  - Recipes draft the spec server-side: `pnl` (area or unit P&L, month and YTD vs last year, budget, forecast,
    bridges, KPIs, landing), `actual_vs_plan`, `monthly_trend`, `cost_base`, `concept`.
  - Run log: one JSON line per answer in `runs.jsonl` and a row in the workbook's hidden `_runlog` sheet.
- Recipe `pnl` reads its facts live on every call into an in-memory store; its check "Live read = source sheet"
  compares them with the source sheets. Materiality: 10 % of |comparison| (counted from NOK 500k) or NOK 5 million;
  a month swing of NOK 5 million is flagged; revenue, personnel cost and EBITDA are always commented. Consumer and
  Infrastructure Growth (abolished 1 June 2026) are refused with their successors.
- Login hardening: a Smart View Login that Windows reports as not responding counts as stuck at once (hidden Excel:
  restarted once; your Excel: "not responding; nothing was closed" after 10 s, never closed); a note after 5 s of
  an open Login; "Empty response from provider" is retried with a progress note, and in the hidden Excel a second
  one restarts only that Excel and retries once. Progress notes during every long stage.
- HTTP API: `POST /gate` and `POST /answer` beside `GET /health` (same calls as the MCP tools; 400 / 401 / 409 busy
  / 413 / 503).
- `[svhttp]` in `engine\defaults.toml`: the data source names and `allow_test`; environment variables
  `SV_HTTP_FINMAX_SERVER`, `SV_HTTP_FINPLAN_SERVER`, `SV_HTTP_ALLOW_TEST`, `SV_HTTP_TIMEOUT`.
- Semantic catalogue `knowledge\catalog.toml` (P&L lines and checks per cube, business areas, scenarios, KPIs,
  materiality) with a `smartass doctor` check.
- Faster answer sheets: answer cells written as one block, check colours set once per run of cells.

### Changed

- The SSO grace for the hidden Excel's sign-in (`LOGIN_GRACE_S`, env `SV_LOGIN_GRACE`) defaults to 30 s.
- Engine modules reorganised: packages `excel/` (the user's Excel) and `workbench/` (xl_* tools), `answer` +
  `answer_write`, `host` + `hostproc`.
- Tool count 35 -> 25.

### Removed (breaking)

- Tools `sv_query`, `sv_ws_pull`, `sv_ws_build`, `sv_ws_answer`, `sv_analyse`, `sv_concept`, `fin_pnl`,
  `sv_scope`, `sv_warehouse`, `sv_refresh_workbook`: use `sv_gate` + `sv_answer` (recipes for P&L, plan, trend,
  cost base and concepts); refresh an old workbook in Excel with Smart View's Refresh.
- The DuckDB warehouse (`warehouse.duckdb`, `paths.warehouse`, the doctor's warehouse check) and the `duckdb`
  dependency.
- Norwegian sheet names: sheets are `Q<n> Answer`, `Q<n> <cube>`, `Q<n> Lines` and `Sources`; old `Svar` /
  `Kilder` / `Linjer` sheets are no longer recognised as answers.

## 0.2.0

- Repo layout: the smartview MCP moves to the git repo `smartass` (package `smartview_mcp` under `engine\`),
  with `knowledge\`, `packs\`, `skills\`, `installer\`, `scripts\` and `docs\`.
- `uv` with a lockfile; Python 3.12; ruff.
- Config layering: `engine\defaults.toml` -> `%LOCALAPPDATA%\smartass\config.toml` -> environment variables;
  no user-specific paths in code; Excel, Smart View and OneDrive are auto-detected.
- `smartass doctor`: cheap health checks with a fix per check; `--bundle` zips results, versions, redacted config
  and log tails for the maintainer.
- Rotating logs with redaction of tokens and secrets.
- Error messages that name the next step.
- Localhost HTTP API (`127.0.0.1`, per-user secret) with `/health`.
- Updater: checks the release folder's `latest.json` at login, supports pinning (`release.version`), never
  downgrades automatically; `smartass update`.
- Per-user installer `installer\install.ps1` (no admin, idempotent): uv, `smartass setup` (config, MCP
  registration, skill junctions, Startup entries, atlas copy), `smartass doctor`.
- Pre-commit guard `scripts\check_staged.py`: no databases, stray workbooks or tokens in git.
- CI on `windows-latest`: ruff, pytest, data/secret guard.
