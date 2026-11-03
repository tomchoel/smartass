# Operating the smartass MCP

Code: the smartass repo, `engine\` (`<mcp>` below). Design and test history: `docs\DESIGN.md`.
Registered with `claude mcp add --scope user smartass -- "%LOCALAPPDATA%\smartass\venv\Scripts\python.exe" "<mcp>\run_server.py"`.

## Working in the user's Excel (analyses)

- **Pulls.** `sv_answer` reads each pull over Smart View's HTTP provider protocol (~0.2 s per grid once signed in,
  no Excel; signs in through the drill's hidden Edge) and falls back to a Smart View ad hoc grid in Excel. The
  result's `routes` say which was used per source sheet. The first HTTP call in a session takes ~10-12 s, because
  it signs in first.
- **Recipe `pnl`** reads its facts live for every call (the area's 36 months, the open month and the latest
  forecast) into an in-memory fact store; nothing is kept between calls. Its check "Live read = source sheet"
  compares those facts with the source sheets the answer reads.
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
- **Data sheets (Smart View fallback route).** Suppression is forced OFF and the missing label is set to `#NumericZero`.
  Every returned label is verified against the requested member, and a moved row fails the pull.
- **State.** Answer specs are stored in the hidden sheet `_svmeta`, so `sv_answer(followup=...)` works after a server
  restart.
- **Failures.** A failed analysis rolls back the sheets it added.
- **Timeouts.** A call stuck in the USER's Excel is never resolved by killing it. The COM thread is abandoned and that job
  stops at its next call. The user's workbook is never saved by the MCP. New analysis files are saved, with the Internal label.
- **Double-clicks.** The MCP's hidden Excel sets `IgnoreRemoteRequests`, so files the user double-clicks open in their own Excel. The
  setting is per instance and does not persist (verified).

## Recipes and follow-ups

A spec `{"version": 1, "question", "gate", "recipe": {"name", "params", "overrides"?}}` (no `pulls`/`layout`) is
expanded server-side inside the Excel job; the closed month is FinMax's Current Reporting Month. The result carries
`period` and an `outline` (pull ids, column headers, section titles, line keys, check labels).

| Recipe | Params | Defaults | Checks |
|---|---|---|---|
| `pnl` | `area` (a business area by name, or a FinMax unit below one: `un.SY.5OK_W` or `5OK_W`; abolished areas are refused with their successors), `period`, `compare`, `detail` (`compact`/`full`) | closed month; all comparisons | Live read = source sheet (`layout.expect` → static checks), period closed, line bridge |
| `actual_vs_plan` | `entity` (FinPlan entity), `month`, `year`, `basis` (`ytd`/`month`), `forecast`, `finmax_entity` | closed month, YTD, latest forecast, FinMax entity mapped | FinPlan line checks; EBITDA FinPlan = FinMax (`severity` Warning) |
| `monthly_trend` | `entity`, `accounts`, `through`, `compare`, `connection` | `un.SY`; C30, C40S, C70, EBITDA; closed month; with last year; FinMax | sum of months = YTD (Accumulated) |
| `cost_base` | `parent`, `accounts`, `period`, `top`, `connection` | `un.SY`; C50S, C60, C70, STA20Y; closed month; top 10 + Other, share %, last year; FinMax | top + other = total |
| `concept` | `concept` (id from the library), `period`, `compare`, `entity` | closed month vs same month last year | the concept's own checks |
| `lines_ranking` | `accounts`, `entity`, `period`, `by`, `projects`, `top`, `compare`, `group_separate` | `un.SY`; YTD closed month; top 10 + Other; last year (`["ly"]`); group companies separate; FinLite | Blocks = lines total; Total = cube; By project = lines total |

The finance defaults, which recipe the gate names, and what `lines_ranking` draws:
[recipes.md](recipes.md) "Server-side recipes".

`recipe.params` cannot change in a follow-up (ask a new question); `overrides` and follow-ups are patches.

**Patch semantics** (`followup=` and `overrides`): dicts merge and `null` deletes; `pulls` merge by `id`,
`layout.sections` by `title`, `lines` by `key`, `checks` by `label`; `{..., "drop": true}` removes an item; unmatched
items are appended; `{"append": [...]}` adds to a list; any other list replaces. Examples:
- add a line to a section by its title:
  `{"layout": {"sections": [{"title": "Year to date", "lines": [{"key": "fp", "label": "EBITDA (FinPlan)",
  "cells": [{"pull": "fp", "row": ["EBITDA"], "col": ["Actual"]}]}]}]}}` (plus the new pull `fp` in `pulls`);
- append a note: `{"layout": {"notes": {"append": ["Excludes one-offs."]}}}`;
- drop a line: `{"layout": {"sections": [{"title": "Per month", "lines": [{"key": "a1", "drop": true}]}]}}`.

A follow-up redraws `Q<n> Answer` in its own block: unchanged pulls are reused (no fetch), changed ones are fetched
again, the answer sheet is backed up (`~Q<n> Answer`, hidden) and restored if the redraw fails. Cells, formats and
charts the user added outside the block are kept; only charts named `smartass chart*` are replaced. The findings
table (Critical / Warning / Info) sits below the notes.

## Speed (bench route `core`, 2026-11-02; details in `docs/audit/`)

- **Typical times.** A new `sv_answer` shows its answer sheet in 1-5 s and finishes in 5-35 s (an HTTP pull is
  ~0.2 s once signed in); a follow-up that reuses its pulls finishes in about 5-7 s; a drill takes 60-90 s.
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
| Query log (audit trail, JSONL) | `%LOCALAPPDATA%\smartass\logs\queries.jsonl` (`paths.query_log`) |
| Run log (one line per `sv_answer`, JSONL; also the workbook's hidden `_runlog` sheet) | `%LOCALAPPDATA%\smartass\logs\runs.jsonl` (`paths.run_log`) |
| Learned defaults (written by `sv_remember`, read by `sv_gate`; readable TOML, edit or delete by hand) | `%LOCALAPPDATA%\smartass\preferences.toml` (`paths.preferences`) |
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
| Warning `Smart View rejected these members (rc=1000), so they were left out: …` | On a pull, the engine bisected the grid (at most 40 test retrieves), dropped the rejected members and retrieved the rest. Typical cause: a name that occurs twice in one dimension (FinPlan Account `Other`), or a member Smart View cannot place (`SDx Revenue`). Tell the user which lines are missing and pick a more specific member. |
| `rc=57704 … missing tag: mbrs` | `GetChildren`, `GetParent` and `QueryMembers` are unsupported on EPM Cloud Planning. Use `sv_members` (atlas) instead. |
| Progress `EPM gave an empty response …; retrying` | `rc=752 Empty response from provider` (or 57704) on a connect or retrieve in the engine's hidden Excel. It is retried in place once. If it persists, the engine says so, restarts **only its own hidden Excel** once and retries. |
| `EPM kept answering 'Empty response from provider' after a restart of the engine's Excel …` | The restart did not clear it. Wait a minute and ask again; if it persists, the user signs in to Smart View once in their own Excel, then run `smartass doctor --bundle`. |
| Progress `EPM gave an empty response in your Excel; retrying once` / `… trying again` | The same in the user's Excel (Smart View fallback route): retried once; the user's Excel is never restarted. |
| `Smart View in this Excel is not signed in correctly: EPM answers 'Empty response from provider' to every connect …` | The user's Excel has a broken Smart View session and no Login to click. The user signs in once in that Excel (Smart View ribbon → Panel → connect FinMax) or closes that Excel, then you ask again. |
| `Smart View's sign-in window stopped responding (Windows marks it 'Not responding') … The engine restarted its hidden Excel; retry.` | A hung Login window in the engine's own Excel is detected within about 5-7 s, and that Excel is restarted. Retry. `… it happened again`: the user signs in once in their own Excel, then retry. |
| `Smart View's sign-in window stayed open for more than 30s (the automatic SSO click did not complete)` | The SSO watcher did not click (`LOGIN_GRACE_S`, default 30 s). The engine restarted its hidden Excel; retry. If the message names an app holding WebView2 (often PowerPoint), ask the user to close it. |
| Progress `Smart View's Login is open in your Excel; waiting for the automatic SSO click` | Sent after 5 s of an open Login in the user's Excel. Wait; the watcher (or the user) clicks 'Telenor SSO'. |
| `The Smart View Login window in your Excel is not responding; nothing was closed. …` | A Login in the user's Excel hung for 10 s. The user closes that Login window (or that Excel), then you ask again. The engine never closes the user's Excel. |
| HTTP `409 {"error": "busy: one request at a time"}` | `smartass serve`: `POST /gate` or `/answer` while another request runs. Retry when it has finished. |
| "non-standard ad hoc mode" | FinMax was opened once in "Ad hoc" mode with alias table `none` and no POV row. The MCP retries with alias `Default`, then reconnects. If it persists, run `sv_restart`. |
| `Call was rejected by callee` | Excel is busy (opening or recalculating a large file). It is retried automatically. |
| `did not respond within …s` + open windows | The watchdog killed a hang in the engine's Excel. Retry. The listed windows show which dialog blocked Excel. |
| `Cannot connect to provider` plus an `<update>` extension manifest | EPM answered the connect with its extension list. This is transient, and connects are retried 3 times. |
| "Smart View returned a different layout than requested" | Rows or columns moved (suppression, or an unknown or duplicate member). The numbers were NOT used. Check the members with `sv_find_member`. |
| Warning `HTTP route failed (…); the grid was read through Smart View in Excel instead` | The pull WAS answered, through Excel; the warning only says why the faster HTTP route was not used. Likely causes: another session owns the drill browser profile (`DrillBusy`, wait or let it close after its idle minutes); the atlas is incomplete or a dimension build failed (run `sv_atlas_build`); the sign-in timed out (the first HTTP call takes ~10-12 s to sign in). No action needed for the answer itself. |
| "Your Excel is in edit mode…" / "Excel is busy" | A cell is being typed in or a dialog is open. The user presses Enter or Esc, then you ask again. |
| "…is in the Produktregnskapet archive" | Archive guard. Use `target='new'`, or `allow_archive=true` only when the user explicitly asks for it. |
| `Q<n> Answer was not drawn by sv_answer, so it has no answer spec to patch` / `There is no answer to follow up` | A follow-up needs an answer drawn by `sv_answer` in the target workbook. Ask it as a new question with `spec=`. |
| `sv_answer` returns `gate: "ambiguous"` with `questions` and "Nothing was pulled: the gate is still ambiguous for this question" | `sv_answer` re-ran the gate on `spec.question` with `spec.gate.answers`. Ask the questions in the user's language, call `sv_answer(question=<same>, answers={item: answer})` (a hand-written spec: pass the same `answers` as `spec.gate.answers`). Rewording the question is not the fix. |
| `This departs from the standard analysis: … Use the recipe and defaults sv_gate named, or pass deviation='<the user's explicit request>'` | Another recipe than the gate's, a free-form spec where the gate named a recipe, a param that switches off a default (`top` ≠ 10, `compare` without last year, `group_separate: false`), or a free-form ranking that is not top 10 + Other. Use the gate's recipe; only if the user asked for the departure, write a spec with `deviation` (it is written on the sheet and in the run log). |
| `Sheet text that is not English: …` / "Write sheet text in English" | Norwegian words or æ/ø/å in a title, note, header, check label, assumption, driver or deviation. Translate those strings; supplier, project and member names may stay as they are. The chat stays in the user's language. |
| `The drill table(s) Lines_Q<n>_<alias> do not exist in this workbook (Lines tables here: …)` | Drill cells point at a Lines table of another question. Cause (0x800A03EC, 2026-11-02): `sv_drill` wrote `Q1 Lines`, then a new spec took Q2 and its drill cells named `Lines_Q2_*`, which did not exist; Excel rejects such a formula. A new spec with drill cells now takes the q of the pending drill (Lines sheets without an Answer). Fix: drill with `sv_drill(q=<n>, id='<alias>')` first, or use the recipe the gate named (`lines_ranking` drills itself). |
| Warning `The ledger lines could not be drilled, so this shows the cube total only` / `Drill failed: …` | `lines_ranking`'s drill failed: `Q<n> Answer` shows the cube total per account. Ask again to rank the suppliers. If only last year's drill failed, the last-year columns show the cube total. |
| Finding (Warning) `Ledger lines differ from the cube by … NOK (… lines); shown as 'Unexplained gap'` | `lines_ranking`'s lines do not tie to the cube by more than 1 NOK; the difference is its own line. Ask again to re-drill if needed (there is no automatic re-drill). |
| Finding `Q<n> FinMax: unknown members X, Y` | Smart View turned those members into comments: their lines are empty. Find the right names with `sv_find_member` and send a follow-up with the corrected pull. |
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
