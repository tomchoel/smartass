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
- **How an answer reaches Excel.** The answer's sheets are built as an .xlsx without Excel (under
  `%LOCALAPPDATA%\smartass\tmp`, deleted when the answer is in), inserted into the workbook in one call (Excel's
  own insert of a file's sheets: no second workbook is opened) and finished there (tables, formulas, chart labels),
  with Excel quiet. The Excel work goes in three batches and a save, each run in one go inside Excel by a small
  add-in (see **The add-in** below); between the batch that brings the answer in and the one that closes it out,
  the job pauses 0.15 s so Excel paints the new answer with its numbers and charts (`SV_PAINT_PAUSE` in the
  server's environment sets the seconds; 0: no pause). The Excel job takes about 2-2.5 s for an answer without a lines sheet (2.7 s at
  61 sheets), about 4 s with a lines sheet, 1.6-1.9 s for a follow-up, and about 7 s for the first answer of a new
  workbook (the workbook is created). A follow-up or `fin_refresh` replaces the question's `Q<n>`
  sheets and tables whole: anything typed on them goes with them, so keep your own notes on your own sheets.
  A formula elsewhere in the workbook that points at a `Q<n>` sheet or its table (`='Q3 Answer'!E12`,
  `=SUM(Data_Q3[Value])`, a cross-sheet SUM, an array formula, a chart series, a defined name; also on another
  answer's sheet) is re-pointed at the new sheets before the old ones are dropped, so it shows the new answer's
  value: the same cell address, so a cell the new layout no longer has shows what is there now, or #REF! when the
  new answer has no such sheet. The reply warns about a formula Excel refused (a locked cell on a protected sheet, a
  table column the new answer lacks) - that one shows #REF!. **Not re-pointed** (they show #REF!): conditional
  formats, data validation lists, pivot tables built on a `Data_Q<n>` table, chart titles and labels linked to a
  cell, and links from other workbooks. After a follow-up Excel's Find dialog has its default options (Look in
  Formulas, whole cell off, By Rows): the search for these formulas uses Find, and Excel cannot say what they were
  before. Nothing of the user's is saved and Excel's settings are put back; when the answer's workbook is the
  active one it ends on the new Answer, otherwise the user's active sheet is put back.
- **The add-in.** The first answer in an Excel builds a small Excel-DNA add-in (`smartass_run_<hash>.xll`, compiled
  with the .NET compiler that ships with Windows, under `%LOCALAPPDATA%\smartass\inproc`) and loads it into that
  Excel with `RegisterXLL`, silently, for the rest of that Excel's life; each batch of an answer's Excel work is
  then one `Application.Run` into it. Anything that keeps it from loading or answering (no compiler, a 32-bit
  Excel, a policy, a dialog: `RegisterXLL` is given 60 s) makes that Excel use the COM route instead - the same
  work, a little slower - with one warning in `logs\server.log` ("in-process executor not used in Excel pid ...").
  `SV_INPROC=0` in the server's environment forces the COM route. The add-in runs nothing but job files the engine
  writes to `inproc\jobs` (a formula calling it gets #N/A); the user's settings it turns off while it works are
  put back by each batch itself, also when a step fails.
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
| An answer's file while it is moved in (deleted after) | `tmp\` |
| The in-process add-in (one folder per build) and its job files (deleted after each batch) | `inproc\` |

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
| `Excel operation <n> failed: <the operation>: operation <n> (<kind> <name>): <Excel's message>` | One step of the answer's Excel work was refused by Excel: the answer was taken out again (a follow-up's old answer put back) and Excel's settings restored. The operation and Excel's own message say what; often a protected or shared workbook, or a name the user gave a sheet. |
| Warning in `server.log`: `in-process executor not used in Excel pid <n>: <why>` | The add-in did not load in that Excel; answers run over COM (slower, same result). The reason says what (no compiler, 32-bit Excel, `RegisterXLL` refused or blocked). It is tried again when Excel restarts. |
| Repeated `HostError` | `fin_status(restart=True)`, then retry; if it repeats, `smartass doctor --bundle` to the maintainer. |
| HTTP `409 busy` / `401` / `413` | `smartass serve`: one request at a time / wrong secret / body over 1 MB. |

Drill errors (sign-in, blank drill page, too wide, no drill report): [drill.md](drill.md).
Doctor checks one by one: `docs\troubleshooting.md` in the repo.
