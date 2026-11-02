# Operating the smartass MCP

Code: the smartass repo, `engine\` (`<mcp>` below). Design and test history: `docs\DESIGN.md`.
Registered with `claude mcp add --scope user smartass -- "%LOCALAPPDATA%\smartass\venv\Scripts\python.exe" "<mcp>\run_server.py"`.

## Working in the user's Excel (analyses)

- sv_query reads over Smart View's HTTP protocol by default (`source='auto'`: ~0.2 s per grid once signed in, no
  Excel; signs in through the drill's hidden Edge) and falls back to Smart View in Excel. The answer's `route` says
  which was used. The first HTTP call in a session takes ~10-12 s, because it signs in first. Workbooks (sv_ws_*)
  still build live Smart View sheets in Excel.
- fin_pnl(area) answers "P&L for <business area>" from the local warehouse (DuckDB,
  `%LOCALAPPDATA%\smartass\warehouse.duckdb`): the first call fills it (about 8 s over HTTP, measured 2026-10-25:
  12 grids, ~9,000 cells), later calls refresh only what changed (a new closed month, an atlas restatement, a new
  forecast) and read the open month and the latest forecast live. If the warehouse is locked or missing it answers
  from live pulls and says so. `sv_warehouse(action='status')` shows the last fills and their reconciliation;
  `sv_warehouse('rebuild', area=<area>)` re-reads an area's 36 months (after a 'Warehouse = source sheet' DIFF),
  `sv_warehouse('rebuild')` recreates the whole file (also the fix for 'warehouse schema is older'). A job
  that fails on content (#NoAccess, an unknown member) pauses (15 min, then daily); transport errors retry on the next
  call. `smartass doctor` names paused jobs with their last error, and the cells a fill could not read.
  - Rules in force (owner decisions): material = 10 % of |comparison| from NOK 500k, or NOK 5 million; a month swing
    of NOK 5 million or more is flagged; |comparison| under NOK 500k counts as zero; revenue, personnel cost and
    EBITDA are always commented. Variance = current - comparison, positive = better everywhere. Cost per FTE is
    shown positive; personnel cost per FTE is the headline. FTE average = the month before the window + each month
    end (YTD n+1 points).
  - Landing = YTD actual + the remaining months of the latest forecast; the seasonal share comes from EBITDA before
    other items when within 0.5-1.5 x n/12, otherwise YTD / n x 12. The forecast is called stale only when it is two
    quarters or more behind.
  - Abolished areas are explained, not shown: Consumer went to SSM and E&I; Infrastructure Growth to TEC, ADM, E&I,
    SSM and BUS (1 June 2026; history is restated to today's organisation).

- **Attach.** The MCP attaches through the Excel window (`AccessibleObjectFromWindow`). It never uses `GetObject(path)`, because that
  starts a hidden Excel when the file is not registered yet. It always waits until no EPM `Login` dialog is open in that Excel
  before any COM call: a COM call during the login deadlocks Excel.
- **Bridge.** A hidden add-in `%LOCALAPPDATA%\smartass\svbridge_<hash>.xlam` is loaded into the user's Excel. The file name is
  versioned, so a new build never collides with one still loaded. Old builds are pruned when free. Smart View acts on the
  ACTIVE sheet, so every bridge call activates our target sheet first, and the user's view jumps to it.
- **Data sheets.** Suppression is forced OFF and the missing label is set to `#NumericZero`. Excel formatting is kept on
  refresh (option 30). Every returned label is verified against the requested member, and a moved row fails the pull.
- **State.** Layouts are stored in the hidden sheet `_svmeta`, so `sv_ws_answer` works after a server restart.
- **Failures.** A failed analysis rolls back the sheets it added.
- **Timeouts.** A call stuck in the USER's Excel is never resolved by killing it. The COM thread is abandoned and that job
  stops at its next call. The user's workbook is never saved by the MCP. New analysis files are saved, with the Internal label.
- **Double-clicks.** The MCP's hidden Excel sets `IgnoreRemoteRequests`, so files the user double-clicks open in their own Excel. The
  setting is per instance and does not persist (verified).

## Speed (benchmarked 2026-10-14; details in DESIGN.md "Performance pass")

- **Typical times.** A pull takes 6-8 s, an answer 4-7 s, a one-call `sv_ws_build` analysis 20-30 s, and a `spec_patch` follow-up about 4 s.
- **Outliers outside our code.**
  - Smart View asks for an SSO re-login about hourly per Excel, and the grid waits up to about 70 s until EPM shows
    the dialog. The watcher clicks within 1 s.
  - The once-a-day "Current Reporting Month" lookup starts the hidden Excel (about 13 s). It is then cached in `smartass\closed_month.json`.
- **What keeps it fast.** Do not undo these:
  - Dimension labels are cached in `smartass\dims_cache.json`, so user-Excel jobs never start the hidden Excel.
  - TN styles are installed once per workbook (marker name `_tnStylesVer`), and new files come from `smartass\tn_template_<ver>.xltx`.
  - `fast_excel` (no redraw, manual calculation, settings restored) wraps writes and formatting.
  - `PrintCommunication=False` is set around page setup.
  - The lint works without per-cell COM.
  - The atlas descendant query uses `CROSS JOIN` and an index on (connection, parent): 0.03 s instead of 11-45 s.
- **Never skip Smart View's default grid** before a retrieve. It was tried and was followed by a 752 error and two Excel crashes.
- **Several agents in parallel** (live-tested 2026-10-14 with 2 subagents).
  - **Per-job worker.** Every user-Excel job runs on its own COM thread.
  - **Turns per Excel.** Turns are kept per Excel process (`smartass\user_excel_<pid>.lock`), because one Excel runs all automation on one
    thread and Smart View writes to the active sheet.
  - **New analyses spread out.** `target='new'` takes a free Excel, or starts another Excel window on the TN template (once, about 25 s; up to
    `MAX_EXCELS`=4 windows). Only then does it wait.
  - **The same workbook queues.** `active` and path targets queue for that Excel.
  - **Hidden-Excel steps.** These (bridge build, daily closed-month lookup) go through `h.on_host`.
  - **Timeouts.** A timeout aborts only that job's token.
  - **Claude Code itself.** Within ONE agent, Claude Code runs write tools one after another. Parallelism comes from several agents
    (subagents, sessions, Copilot).
- **Stage log.** `server.log` logs `stage > ...` at the start of every stage, which pins a crash or hang to its step.
  `open target` has sub-stages (find Excel windows, new workbook: add from template / save as, ready + bridge).
- **Progress notes.** A stage that runs over 10 s sends an MCP progress note (`still working: <stage>`), repeated
  every 10 s, for the tools that build in Excel and for `sv_drill`.

## How it works

- **One Excel.** The MCP drives one hidden, dedicated Excel (`DispatchEx`) on a single COM thread. The VBA bridge `vba\modBridge.bas`
  and Oracle's `smartview_body.bas` are loaded into an in-memory workbook at every start, and that workbook is never saved.
- **Read-only.** Submit, calculate and business-rule functions do not exist in the bridge.
- **Watchdog.** Every call has a deadline. A hang kills **only this server's own Excel**, recorded in `%LOCALAPPDATA%\smartass\pids.json`
  together with the owning server process, and the next call restarts it. The error text names any dialog the hung Excel had open.
- **SSO.** The global watcher `engine\sso_watcher.ps1` (PowerShell, .NET UI Automation) runs from the Startup folder
  (`smartass-sso-watcher.lnk`, created by `installer\install.ps1`; re-run the installer to recreate it). It clicks "Telenor SSO" in the Smart View Login dialog of *any* Excel, and it never types.
  - It writes a heartbeat to `smartass\logs\sso_watcher.heartbeat`. The MCP restarts it if the heartbeat is older than 90 s.
  - The log is `smartass\logs\sso_watcher_ps.log`.
  - A Python `uiautomation` version was removed because it blocked on this dialog.
  - One login covers both pods: plan1 (FinMax, FinLite) and stratmodel (FinPlan).
- **Saving.** Excel cancels unlabelled `SaveAs` silently, so the MCP sets the **Internal** label
  (`f604d2c9-1577-460e-b668-57374a0216c3`) before every save.

## Files

| What | Where |
|---|---|
| Results (`save_xlsx`, refreshed copies) | the output folder (`paths.output` in `%LOCALAPPDATA%\smartass\config.toml`) |
| Query log (audit trail, JSONL) | `%LOCALAPPDATA%\smartass\logs\queries.jsonl` (`paths.query_log`) |
| Run log (one line per `sv_answer`, JSONL; also the workbook's hidden `_runlog` sheet) | `%LOCALAPPDATA%\smartass\logs\runs.jsonl` (`paths.run_log`) |
| Server log | `%LOCALAPPDATA%\smartass\logs\server.log` |
| SSO watcher log | `%LOCALAPPDATA%\smartass\logs\sso_watcher.log` |
| Atlas | `%LOCALAPPDATA%\smartass\atlas.db` |
| Venv (mcp 2.3.0, pywin32 312, psutil, uiautomation) | `engine\.venv` in the install folder (`%LOCALAPPDATA%\smartass\app`, or the git checkout) |

## Routine tasks

- **Refresh the atlas** when `sv_status` says `stale` (older than 30 days), or after a reorganisation. Run `sv_atlas_build`, which takes about 3 minutes for all
  three cubes. Then regenerate the readable atlas with
  `"%LOCALAPPDATA%\smartass\venv\Scripts\python.exe" "<mcp>\gen_atlas_md.py"`, which writes `reference\cubes.md` of this skill.
- **Apply code changes: automatic.** The registered entry `run_server.py` is a relay. It runs the real server
  (`child_server.py`) as a child process. When a tool is called after any edit to `smartview_mcp\*.py`, `vba\*.bas` or
  `child_server.py`, it restarts that child gracefully (its Excel closes and options are restored), then tells Claude Code the tool list changed.
  No `/mcp` is needed. Only an edit to `run_server.py` itself needs a reconnect. `sv_status` shows `server_started` and
  `behind_reload_proxy`. The relay log is `smartass\logs\proxy.log`, and the live test is matrix scenario `S10.3.reload`. `sv_restart`
  restarts only Excel.
- **Tests.**
  - Unit: `python -m pytest -q` in `<mcp>`.
  - Live: `uv run python tests\live\matrix.py [--only TAGS]` in `<mcp>` (see `tests\live\README.md`).
  - Live tests start their own Excel. They no longer touch the server's.

## Errors and what they mean

| Symptom | Meaning / action |
|---|---|
| `rc=1000 … Server says: Member X is out of place` | The grid layout was rejected. Check that the dimension names and member names are right. |
| `rc=1000 … Unsupported Error` on a refreshed workbook sheet | The sheet references members that no longer exist (often after a reorg). Rebuild it with `sv_query`. |
| Warning `Smart View rejected these members (rc=1000), so they were left out: …` | On a pull or query, the engine bisected the grid (at most 40 test retrieves), dropped the rejected members and retrieved the rest. Typical cause: a name that occurs twice in one dimension (FinPlan Account `Other`), or a member Smart View cannot place (`SDx Revenue`). Tell the user which lines are missing and pick a more specific member. |
| `rc=57704 … missing tag: mbrs` | `GetChildren`, `GetParent` and `QueryMembers` are unsupported on EPM Cloud Planning. Use `sv_members` (atlas) instead. |
| `rc=752 Empty response from provider` | A transient EPM or session issue. The MCP retries the connect three times. If it persists, run `sv_restart`. |
| "non-standard ad hoc mode" | FinMax was opened once in "Ad hoc" mode with alias table `none` and no POV row. The MCP retries with alias `Default`, then reconnects. If it persists, run `sv_restart`. |
| `Call was rejected by callee` | Excel is busy (opening or recalculating a large file). It is retried automatically. |
| `did not respond within …s` + open windows | The watchdog killed a hang. Retry. The listed windows show which dialog blocked Excel. |
| `Cannot connect to provider` plus an `<update>` extension manifest | EPM answered the connect with its extension list. This is transient, and connects are retried 3 times. |
| "Smart View returned a different layout than requested" | Rows or columns moved (suppression, or an unknown or duplicate member). The numbers were NOT used. Check the members with `sv_find_member`. |
| Warning `HTTP route failed (…); the grid was read through Smart View in Excel instead` | The grid WAS answered, through Excel; the warning only says why the faster HTTP route was not used. Likely causes: another session owns the drill browser profile (`DrillBusy`, wait or let it close after its idle minutes); the atlas is incomplete or a dimension build failed (run `sv_atlas_build`); the sign-in timed out (the first HTTP call takes ~10-12 s to sign in). No action needed for the answer itself. |
| "Your Excel is in edit mode…" / "Excel is busy" | A cell is being typed in or a dialog is open. The user presses Enter or Esc, then you ask again. |
| "…is in the Produktregnskapet archive" | Archive guard. Use `target='new'`, or `allow_archive=true` only when the user explicitly asks for it. |
| Bare "Error executing tool …" with no reason | The server is running code from before the 13.10.2026 fix (the Excel was dead: `AttributeError: Excel.Application.Ready` in server.log). Run `sv_restart`, then ask the user to reconnect `smartass` with `/mcp`. |
| Unknown-member warning | Smart View turned the text into a comment. Find the right name with `sv_find_member`. |
| `Invalid query: unknown member 'X' (FinMax / All Entities)` from `sv_analyse` | The template member is not in the atlas and EPM confirmed live that it does not exist in that dimension. Fix the param (for `actual_vs_plan`, pass `finmax_entity`). If the live check itself fails, the member is pulled unverified instead. |
| Warning `error cells on '<sheet>': N x #NoAccess (first: …)` / `formula error on line 'k'` | No read access, or a member not valid in that cube, on a data sheet; the answer lines that depend on it are errors. Tell the user these numbers are missing; never report them as zero. |
| `Sign-in to EPM did not complete - open EPM once in Edge and log in, then retry sv_drill` | The drill's silent SSO login stalled (the browser nudges a hung `/oauth2/callback` once). The user opens EPM in Edge once, then you retry. |
| `The drill page loaded but never requested data (also after one reload)` | The Oracle drill page stayed blank. Retry once; if it repeats, the page has changed (screenshot path in the error). |
| `The hidden drill browser crashed … retry sv_drill` / `ran out of time before it could sign in` | The headless browser died or started too slowly within the drill's time limit. Retry. |
| `this cell has no drill-through report: it is not a FinLite base-level intersection …` | The cell's data did not come through Data Integration. Use a FinLite account and a lower entity or period. |
| Progress `still working: <stage> (n s)` | A long stage is running (sent every 10 s). Wait; do not call again. |

## Smart View settings the MCP touches

- **Global, restored on shutdown.** These are the user's preferences: 104 messages → none, 109 home panel → off, 110 comment dialog → off.
  The originals were 104=2, 109=True and 110=False.
- **Sheet level, our sheets only.** 5 indent → none, 13 missing label → `#Missing`, 14 no-access label → `#NoAccess`.
  The user's own workbooks show missing as `#NumericZero` (0).
- **Leave member display (16) and alias tables alone.** Smart View can persist them onto the *connection*. A sheet-level change on 2026-10-13 coincided
  with FinMax opening in alias table `none` mode.
