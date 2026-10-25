# Changelog

## Unreleased

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
