# Changelog

## Unreleased

- `fin_pnl(area)`: a business area's P&L in one call - month and YTD against last year (FinMax), budget and the
  latest forecast (FinPlan), materiality flags, an EBITDA bridge by line and by sub-unit, KPIs and the year-end
  landing. The numbers come from a local DuckDB warehouse (`%LOCALAPPDATA%\smartass\warehouse.duckdb`,
  `paths.warehouse`), refreshed by itself when a month closes or the atlas shows a restatement; the open month and
  the latest forecast are read live and every answer has a freshness stamp. The workbook keeps live Smart View
  source sheets, formulas only, and a "warehouse = source sheet" check per sheet.
  - Materiality: a variance is material at 10 % of |comparison| (counted from NOK 500k) or at NOK 5 million; a
    month swing of NOK 5 million or more is flagged; a comparison under NOK 500k counts as zero; revenue,
    personnel cost and EBITDA are always commented.
  - Variance = current - comparison, positive = better on every line. Cost per FTE is shown positive; personnel
    cost per FTE is the headline KPI. FTE averages use the month before the window plus each month end (YTD n+1
    points).
  - Year-end landing = YTD actual + the remaining months of the latest forecast; the seasonal share is taken from
    EBITDA before other items when it lies within 0.5-1.5 x n/12, otherwise the run rate YTD / n x 12. A forecast
    counts as stale only when it is two quarters or more behind.
  - Consumer (into SSM and E&I) and Infrastructure Growth (into TEC, ADM, E&I, SSM and BUS) were abolished
    1 June 2026: asking for them gives an explanation and the successors, never empty figures.
- `sv_warehouse(action='status'|'refresh'|'rebuild', area=None)` (rebuild re-reads an area's 36 months, or
  recreates the whole file; the warehouse carries a schema version and an older file is never mixed in); `smartass doctor` has a `warehouse` check (last fill, cells that were
  missing, paused refresh jobs with their last error, age) and lists the duckdb version in `--bundle`.
- The closed month is read over Smart View's HTTP protocol first (EnumMembers), with Smart View in Excel as the
  fallback.
- `fin_pnl` also takes a unit below a business area (any FinMax entity the atlas knows, e.g. `un.SY.5OK_W` or
  `5OK_W`). A unit is always read live and never kept in the warehouse (its checks say "live read = source
  sheet"); with no FinPlan member (most units) it is compared with last year only, and says so.
  `sv_warehouse('rebuild')` still takes business areas only.
- The `sv_analyse` template `division_pnl` is removed: use `fin_pnl`, which covers areas and units.
- New dependency: `duckdb==2.0.0`.
- Faster answer sheets: the answer cells are written as one block (a fin_pnl answer: 11 s -> 0.2 s) and the check
  colours are set once per run of check cells (7 s -> 0.2 s). A block Excel refuses is written cell by cell, so a
  bad formula still costs only its own warning.
- `sv_query` reads over Smart View's HTTP provider protocol first, from the drill's hidden Edge (no Excel):
  `source='auto'` (default) falls back to Smart View in Excel and says why in a warning; `'http'` never falls
  back; `'com'` forces Excel. The answer's `route` says which was used.
- `[svhttp]` in `engine\defaults.toml`: the Smart View data source names (`finmax_server`, `finplan_server`) and
  `allow_test` (a test pod is refused by default); environment variables `SV_HTTP_FINMAX_SERVER`,
  `SV_HTTP_FINPLAN_SERVER`, `SV_HTTP_ALLOW_TEST`, `SV_HTTP_TIMEOUT`.
- Semantic catalogue `knowledge\catalog.toml`: P&L lines and checks per cube, business areas with aliases,
  scenarios, KPIs and materiality.
- `smartass doctor`: catalogue check (structure, and members against the atlas).
- The P&L templates read their lines and checks from the catalogue.

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
