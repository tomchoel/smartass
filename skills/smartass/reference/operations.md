# Operating the smartass MCP (v2)

## How it runs

- **One MCP entry, `smartass`.** Claude Code starts `engine\run_server.py`, a small relay that runs the real
  server `engine\child_server.py` (`smartass.server`) as a child process and passes JSON-RPC through.
- **Code changes roll only on request.** The relay replaces its child when `engine\VERSION` changes
  (`scripts\bump.ps1` bumps it together with `pyproject.toml`) or when a client calls `fin_status(restart=True)`.
  Editing a `.py` file alone does nothing to a running server. After a reload, Claude Code may keep an old tool
  schema: run `/mcp` and reconnect smartass. Only an edit to `run_server.py` itself always needs a reconnect.
- **Two Excels.** The engine's own hidden Excel (`DispatchEx`, one COM thread, the VBA bridge `vba\modBridge.bas`
  + Oracle's `smartview_body.bas` loaded into an unsaved workbook) does ad hoc fallback grids; the answer is drawn
  in the user's Excel (session workbook, or `target="active"`). Only the engine's own Excel is ever killed or
  restarted; the PIDs it started are in `%LOCALAPPDATA%\smartass\pids.json` with the owning server.
- **Fetch route.** Smart View HTTP first: with the user's OAuth token where a pod has one (`smartass login`, no
  browser), else through the shared hidden Edge (signs in silently). A grid HTTP could not fetch is pulled as an ad
  hoc grid in Excel instead (`routes` in the reply; a warning says why).
- **Read-only.** Submit, calculate and business rules do not exist in the bridge.
- **Smart View options.** Global options 104 (messages: none), 109 (home panel off), 110 (comment dialog off) are
  set while the engine's Excel runs. The user's originals are saved in `smartview_options_original.json` and
  restored on shutdown; after a crash they are restored at the next start.
- **Parallel sessions** are fine. Each session has its own workbook; writes into one Excel window take turns (a
  lock per Excel), different Excels run in parallel. Long stages send `still working: <stage> (n s)` every 10 s:
  wait, do not call again.

## Files

| What | Where (`%LOCALAPPDATA%\smartass\` unless noted) |
|---|---|
| User settings | `config.toml` (defaults: `engine\defaults.toml`) |
| HTTP API secret (DPAPI, this Windows user only) | `secret.bin` |
| Atlas (all dimension trees) | `atlas.db` |
| Server log, relay log | `logs\server.log`, `logs\proxy.log` |
| Run log (one line per answer) | `logs\runs.jsonl` |
| Query log (audit) | `logs\queries.jsonl` |
| SSO watcher log + heartbeat | `logs\sso_watcher_ps.log`, `logs\sso_watcher.heartbeat` |
| Sessions (which workbook belongs to which session) | `sessions.json` |
| Drill cache, drill browser profile, failed drill pages | `drill.db`, `drill_edge\`, `logs\drill_fail_*` |
| Text repairs of your own | `text_fixes.json` |
| Answer workbooks | `paths.analyser` (`<OneDrive>\smartass\analyser\<project>\`) |

The spec of every answer is stored in the workbook (hidden `_svmeta`); that is what follow-ups and `fin_refresh`
read.

## Command line

| Command | Does |
|---|---|
| `smartass doctor` (`--json`, `--bundle`) | Health checks with a fix per problem; `--bundle` zips logs, redacted config and versions for the maintainer |
| `smartass atlas status` / `smartass atlas build [--cube X] [--dimension D]` | Atlas age; rebuild from Smart View (own hidden Excel, read-only, a few minutes). Then `gen_atlas_md.py` refreshes `reference\cubes.md` |
| `smartass login` (`--status`, `--force`, `--forget`) | EPM sign-in for the bearer route: prints a page and a code (page opens, code on the clipboard), waits for the approval. Needed once, and again after 7 days without use |
| `smartass setup` | Per-user setup (config, secret, MCP entries, skill link, Startup entries, PATH, atlas copy) |
| `smartass update` | Install the release from the release folder |
| `smartass serve` | Localhost HTTP API (started at login): `GET /health`, `POST /analyze` with `{"spec": ..., "target": "session"}`; header `X-Smartass-Secret`; one request at a time (409 while busy) |

## Errors and what they mean

Every error names one next step. The common ones:

| Symptom | Meaning / action |
|---|---|
| `unknown entity 'X' in FinMax - find the member with fin_lookup` (or an unknown metric) | Hard guard: nothing was drawn. Look the name up and send the spec again. |
| `invalid spec: …` / `Extra inputs are not permitted` | The spec breaks the contract (unknown field, a block metric not in its query, budget on a FinMax query …). Fix the named field. |
| `query q1: Actual covers … but … covers …` | Actual and plan periods do not line up; set `plan_periods` or align `periods`. |
| Warning `Smart View rejected these members …, so they were left out` | Those lines are missing (not zero); tell the user and pick a more specific member. |
| Warning `HTTP route failed (…); the grid was read through Smart View in Excel instead` | The data WAS fetched, through Excel. No action needed for the answer. |
| `checks.sheet_equals_model` = `DIFF in n cells` (+ `mismatches`) | Excel calculated something else than the model: report it; do not quote those cells. |
| Login window / `Smart View's sign-in window stopped responding` / `stayed open for more than 30s` | Sign-in: the SSO watcher clicks 'Telenor SSO'. The engine restarts its own hidden Excel when its Login hangs; a Login in the user's Excel needs the user. If an app holding WebView2 is named (often PowerPoint), ask the user to close it. |
| Doctor `WARN oauth <pod>: not signed in` / `sign-in ended` | Answers still work (hidden browser), only slower to start. Run `smartass login` in the background, show the user the page and code it prints, and let it wait for their approval. |
| `EPM gave no response; restarting the engine's hidden Excel` / `Empty response from provider` | Retried once after a restart of the engine's Excel; if it persists, the user signs in once in their own Excel, then ask again. |
| `Your Excel is in edit mode…` / `Excel is busy` | A cell is being typed in or a dialog is open: the user presses Enter or Esc, then ask again. |
| `…is in the Produktregnskapet archive` | Archive guard: answers never go into the archive. Use the session workbook. |
| `there is no v2 answer in this workbook to follow up` / `Q<n> is not a v2 answer in this workbook` | Follow-ups and `fin_refresh` need an answer drawn by `fin_analyze` in that workbook (the error lists the ones there). Otherwise ask it as a new question. |
| `Another agent has been working in that Excel for 10 min` | Turn-taking timeout in one Excel window. Try again shortly. |
| Repeated `HostError` | `fin_status(restart=True)`, then retry; if it repeats, `smartass doctor --bundle` to the maintainer. |
| HTTP `409 busy` / `401` / `413` | `smartass serve`: one request at a time / wrong secret / body over 1 MB. |

Drill errors (sign-in, blank drill page, too wide, no drill report): [drill.md](drill.md).
Doctor checks one by one: `docs\troubleshooting.md` in the repo.
