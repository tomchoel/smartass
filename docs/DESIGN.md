# Smart View MCP: design brief (2026-10-13)

The decisions come from the grilling session (memory: smartview-mcp-plan). This file records what the research changed or confirmed.

## Architecture

- **Engine: Excel + Smart View via COM, not REST.** EPM REST `exportdataslice` exists, but SSO users cannot use basic auth. OAuth needs a one-time IDCS admin registration plus a device-code login, with a refresh token that lasts 7 days per pod. It also exports *"only cells for which the user has read-write access"*, which risks blank actuals. Excel stays the engine. REST is a possible future path if an admin registers an OAuth client.
- **Python 3.14, venv `%LOCALAPPDATA%\svkit\venv`** (outside OneDrive). `mcp==2.3.0` (SDK v2: `from mcp.server import MCPServer`), pydantic 2.14.1, pywin32 312, psutil 7.2.2, uiautomation 2.0.29, openpyxl, pytest, anyio. Runs as a stdio server registered with `claude mcp add --scope user`, using the venv's `python.exe` by absolute path.
- **Smart View installed: 25.100 (64-bit), 2026-10-13.** Software Center labels it "25.200". WebView2 login is on (`disableWebView2=0`). `smartview.bas` is unchanged in substance (2293 lines).
- **One dedicated STA worker thread owns all COM.** Tools are `async` and hand work to that thread. A watchdog enforces a deadline per call. On a hang it kills **only our own PID**: the `Hwnd`→PID is verified, the process creation time matches, and the command line contains `-Embedding`. Then it rebuilds Excel and re-connects. Claude Code's own tool timeout is about 28 h, so we can't rely on it.
- **Excel instance:** `DispatchEx` (never `Dispatch`/`GetObject`), with the PID checked against processes that existed before. Settings: `Visible=False`, `DisplayAlerts=False`, `AutomationSecurity=1`, `AskToUpdateLinks=False`. `EnableEvents`/`Interactive` are only changed after a spike test, because they may break Smart View or the SSO dialog.
- **The VBA layer ships as a prebuilt `svbridge.xlsm`.** That avoids importing modules at runtime. Every entry point has an `On Error` handler that returns a string. `Application.Run` **returns values**, so the Log-sheet workaround is gone.
- **SSO:** UIA watcher on its own thread (never the COM thread). It scopes `#32770 "Login"` windows to our PID and invokes "Telenor SSO". Verified working. FinMax/FinLite live on `plan1-telenorepm`, FinPlan on `stratmodel-telenorepm`, so FinPlan likely needs its **own** SSO login. The watcher handles that.
- **Coexistence with excel-mcp (decided 2026-10-13).** smartview MCP = the data pipe (EPM → table/xlsx); excel-mcp = the workbench (formatting, pivots, charts on saved files). The hand-off: smartview saves AND closes the file, then excel-mcp opens it.
  - smartview kills ONLY the PIDs it started itself (PID + creation time recorded in `svkit\pids.json`). It never sweeps "all -Embedding" processes, because that would kill live excel-mcp sessions. The produktregnskapet skill's troubleshooting must be corrected to match.
  - The **global SSO watcher** is a separate, always-on process started at logon. It clicks "Telenor SSO" in the Login dialog of ANY Excel (smartview MCP, excel-mcp, the user's own). It only ever invokes that one button and never types. This also fixes excel-mcp's startup hangs (0x80080005). The MCP keeps a built-in fallback watcher scoped to its own PID in case the global one isn't running.
- **Work files** go in `%LOCALAPPDATA%\svkit\` (short paths, no OneDrive locks). Outputs, logs and the atlas index summary go in `prodreg2026\smartview\`. The SQLite atlas lives in `%LOCALAPPDATA%\svkit\atlas.db`.

## Smart View settings applied at connect (`HypSetOption`)

| Code | Setting | Value | Why |
|---|---|---|---|
| 104 | Log message display | 3 (none) | No modal message boxes |
| 109 | Home panel on startup | False | Avoid panel/dialog churn |
| 110 | Comment dialog on refresh | False | No modal dialogs |
| 16 | Member display | 1 (name + description) | We parse `code alias (prefix)`. Switch to 0 for pure name grids |
| 6/41/44 | Suppress missing rows/cols/blocks | per query (default on for sv_query) | Smaller grids |
| 33 | Navigate without data | True during pivot/zoom, False before final retrieve | Speed |
| 105/113 | Route log to file | on, under `svkit\logs` | Diagnostics |

## Tools

**Primary:**
- `sv_status` (connection health, version, pid)
- `sv_connect`
- `sv_query(connection, pov, rows, columns, suppress, save_xlsx, source='auto'|'http'|'com')` → table (HTTP route first, Excel fallback; `route` in the answer)
- `sv_find_member` (atlas)
- `sv_refresh_workbook(copy_of_path)`
- `sv_restart`

**Advanced:** `sv_member_info` (member properties). The ad hoc tools (`sv_adhoc_*`: zoom, pivot, keep/remove only, options) and `sv_show` were removed in the 2026-10 reliability audit: no agent called them in 30 days. The VBA bridge keeps the `B_*` macros the engine itself uses.

**Never exposed:** any submit, calc, or business-rule function.

## Guards

- Size guard: more than 50k cells needs `large=true`.
- Every call is logged to `prodreg2026\smartview\logs\queries.jsonl` (connection, POV, axes, cells, duration, return code).
- Errors are actionable. On rc≠0 the tool calls `HypGetLastError` for the server text. An unknown member is detected (Smart View silently turns it into a comment) and the tool suggests `sv_find_member`.
- Known issue 39163160 (refresh fails after idle): retry once automatically.
- Results over the token budget are truncated with a pointer to the saved xlsx.

## Return codes

| Code | Meaning |
|---|---|
| -15 | Invalid parameter |
| -41 | Not supported by provider (official) |
| -71 | Invalid member |
| -58 / -62 | No grid / no Smart View content on the sheet |
| 1000 / 57704 | Not documented. Call `HypGetLastError`. 57704 is also seen by others on `HypQueryMembers` |

**rc=1000 "Unsupported Error" from one member (2026-10-14).** A FinPlan pull of 654 accounts failed whole because of
two members: `Other` (the name occurs twice in Account, under *Other cash items* and *Statistics*) and `SDx Revenue`.
`ops.query` (sv_query) and `userexcel.pull` (sv_ws_pull/sv_ws_build) now bisect each multi-member axis on rc=1000 (at
most 40 test retrieves, about 1 s each; 8–9 s for a 3–7 member grid), drop the rejected members, retrieve the rest and
warn. The warning is stored in `_svmeta`, so answers rebuilt later still report it, and `_answer_job` adds it as a note
on the answer sheet.

## Spike results (2026-10-13, Smart View 25.100) — scripts in `spikes\`

- **Regression after the upgrade: OK.** WHS C30 YTD Sep 26 = 1,994,151,357.55, identical.
- **SSO is shared across pods.** One click at Excel startup covered both FinMax (plan1) and FinPlan (stratmodel). The FinPlan connect took 1.2 s with no prompt.
- **Speed.** Excel start ~6-8 s plus a connect of ~5 s the first time. A full entity tree under `un.SY` (1,478 rows) takes: zoom all levels 1.4 s, retrieve 2.4 s, `Value2` read 0.09 s.
- **Member API on the Planning provider.**
  - Works: `HypFindMember`, `HypGetDimensions`, `HypGetSheetInfo`, and **`HypGetMemberInformationEx`**. The last returns Name, Dimension, Level, Generation, ParentMbrName, ChildMbrName, ChildCount, AliasNames, Consolidation Operator, Data Storage and more.
  - Fails with 57704 "missing tag: mbrs": `HypGetChildren`, `HypGetParent`, `HypQueryMembers`, `HypOtlGetMemberInfo`.
  - Fails: `HypExecuteQuery` (MDX) with -9; `HypGetSubstitutionVariable` with 10005; `HypListApplications` with -15.
  - `HypGetPOV`: entry point missing in HsAddin. `HypGetMembers` returns only the top member.
  - **The atlas is therefore built by ad hoc zoom (all levels) per dimension, plus MemberInfoEx on demand.**
- **FinPlan dims:** Years, Period (columns); Account, Entity (rows); Scenario, Version, All Products, All Balance Types (POV).
- **Saving from Excel needs a sensitivity label.** Unlabelled `SaveAs` is cancelled silently, with no error and FullName still "BookN"; `EnableEvents` doesn't change that. Set it with `wb.SensitivityLabel.SetLabel(info, info)` (LabelId `f604d2c9-1577-460e-b668-57374a0216c3`, SiteId `1676489c-5c72-46b7-ba63-9ab90c4aad44`, AssignmentMethod 1), then SaveAs, which works for both xlsx and xlsm. The user chose **Internal** as the default label; the ID is the most common label in prodreg2026 (199/342 files), pending the user's confirmation that it is "Internal".
- **The excel-mcp `create` writes xlsx format even with a `.xlsm` name.** The earlier "svkit_template.xlsm" files were VBA-less xlsx. Removed.
- **Bridge.** It is built in memory at every start from the `.bas` files (AccessVBOM=1, under 1 s) and never saved. `Application.Run` returns values and arrays (tuples) fine.
- **Do NOT toggle `EnableEvents` mid-session.** Retrieve works with it False, but after toggling back to True a retrieve on a new sheet failed with 57704. Keep the default (True), which is proven across all runs.
- **Global Smart View options are user preferences.** Originals were 104=2, 109=True, 110=False. The MCP sets its values at start and restores them on shutdown.

## Build results (2026-10-13) — MCP v0.1 registered in Claude Code (user scope)

- **Unit tests:** 7 pass (`tests/test_grid.py`).
- **Live smoke (`tests/live_smoke.py`):** passes. WHS query returns 1,994,151,357.55 in 1-2 s. A cold start (Excel + SSO + connect) takes 15-27 s. Unknown members are flagged. save_xlsx works with the label. FinPlan works.
- **Watchdog (`live_smoke.py hang`):** passed twice. A forced hang is killed after 5.2 s; the next query restarts, signs in and returns correct numbers in about 10 s. One earlier run hung 180 s on restart (cause unseen), so timeouts now log the hung Excel's open windows.
- **Atlas:** all 3 cubes, built in about 3 min.
  - FinMax: 14 dims, e.g. Projects 139,578, Reserves 104,764, Entities 25,465, Products 19,755, Accounts 8,008.
  - FinLite: 10 dims.
  - FinPlan: 8 dims (Account 1,388, Entity 337).
- **Refresh (`tests/live_refresh.py`)** on a copy of `2025\2 Input\20 Inntekt-varekost\Wholesale\2025_Inntekt og varekost WHS_Large Sites.xlsx`:
  - Both sheets refreshed and saved with the label; the source was untouched.
  - 690 of 753 numbers are identical to the archive (22.01.2026 extract).
  - The 63 that differ are all `3WHS_C` parent rows, which now return 0. The **1 June 2026 reorg** moved `4REV_W` from `3WHS_C` (now an empty node under `2CWK`) to `3WHS_W`.
  - **Leaf data is unchanged; parent rows are recomputed on today's hierarchy.** Skill trap.
- **Transient odd mode seen once:** FinMax opened as sheet type "Ad hoc" with alias table "none", no POV row, and HypGetDimensions returning "Accounts"/"TimePeriods". The connect just before it returned rc=752 "Empty response from provider".
  - Fix: connect retry, plus a standard-mode check after every new grid (retry with alias Default, then reconnect).
  - Dimension names are taken from grid labels, and `canon_dim` accepts both forms.
- **Unexplained automation Excel** with a "Save As" dialog (pid 34388, created while our test Excel was blocked). Not ours, so it was left alone; probably excel-mcp reconnecting.

## Analysis in the user's Excel (built 2026-10-13, design in memory `smartview-analysis-ux`)

- **New modules.**
  - `userexcel.py`: attach via the window, versioned bridge xlam, login wait, live data pulls, `_svmeta`, rollback, job generations.
  - `answer.py`: declarative answer spec → `Q<n> Svar` with live INDEX/MATCH formulas.
  - `templates.py`: `division_pnl`, `monthly_trend`, `actual_vs_plan`, `cost_base`.
  - `styling.py`: the tn-rapportstil engine applied to an *open* workbook (built by a sub-agent).
  - `atlas_history.py`: snapshots and a diff, giving `sv_atlas_changes`.
- **New tools.** `sv_analyse`, `sv_ws_pull`, `sv_ws_answer`, `sv_atlas_changes`.
- **Live regression** (`tests/live_analysis.py full`): all four templates plus a free-form pull/answer/rebuild-after-restart.
  - Every check is OK: P&L 9/9, trend 4/4 (sum of months = Accumulated), plan 6/6.
  - Technology YTD Sep 26 EBITDA is (1 112 834 010), the same as in earlier sessions. FinPlan Actual − FinMax = −859 812.
- **Independent audit (sub-agent).** It found no EPM writes and no archive writes. Fixed:
  - **HIGH:**
    - no COM before the login is cleared;
    - kill only a single, unambiguous own start-up Excel;
    - the deadline starts when the job starts, abandoned jobs stop via a generation counter, and the old thread gets a sentinel;
    - **suppression forced off and every label verified (no silent row shift)**;
    - every tool error is readable via a guard decorator.
  - **MEDIUM:**
    - macros force-disabled when refreshing user files;
    - rollback in every write path, restoring DisplayAlerts;
    - a `Kilder SV` fallback name;
    - a single query retry;
    - positional formulas when labels are duplicated;
    - waterfall total bars;
    - sheet activation before every bridge call;
    - validation in pull;
    - an archive check before opening and before saving;
    - correct WRITE/ADMIN annotations;
    - docstrings;
    - sessions persisted in `_svmeta`;
    - busy retries incl. 0x800AC472 everywhere;
    - a non-blocking `sv_restart`.
  - **LOW:**
    - `_Q` matches our sheets only;
    - dead code removed: the Python SSO watcher, `B_ExecuteQuery` (an MDX insert could write), `B_RefreshAll`;
    - an SSO watcher heartbeat with restart when stale;
    - original Smart View options persisted once (crash-safe).
- **Known limits.**
  - A pull of 200 cost centres × 4 accounts takes about 2-3 min, mostly per-cell formula writes and styling.
  - A failed follow-up rebuild clears the previous answer sheet.
  - The atlas purges a cube before rebuilding it.

## Open spikes (resolved 2026-10-13: 1-5)

1. Smart View 25.200 upgrade (user, via Software Center). Then re-extract `smartview.bas` and confirm the connections survived.
2. Do `EnableEvents=False` / `Interactive=False` break Smart View or the SSO dialog?
3. Does FinPlan (stratmodel pod) prompt a second SSO login, and does the watcher clear it?
4. Which member functions work on the Planning provider (QueryMembers, GetMemberInformation, FindMember, ExecuteMDXEx, GetMembers)? This decides how the atlas gets built: via API or via the zoom-all-levels fallback.
5. Is reading a 470×80 GL grid with `Value2` within the time budget?

## Sources

- Smart View VBA 24.100: https://docs.oracle.com/en/applications/enterprise-performance-management/smart-view/24.100/dvesv/toc.htm
- Options: …/24.100/dvesv/hypgetoption.html
- Return codes: …/23.200/dvesv/vba_return_values.html
- 26.100 readme (SSO/WebView2 fixes, known issues): https://docs.oracle.com/en/applications/enterprise-performance-management/smart-view/26.100/reamv/index.html
- EPM REST exportdataslice: https://docs.oracle.com/en/cloud/saas/epm-cloud/prest/export_dataslices.html
- OAuth: https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/prest/authentication_oath.html
- MCP Python SDK: https://py.sdk.modelcontextprotocol.io/
- Spec: https://modelcontextprotocol.io/specification/latest/server/tools
- Claude Code MCP: https://code.claude.com/docs/en/mcp
- Anthropic tool design: https://www.anthropic.com/engineering/writing-tools-for-agents
- COM: IMessageFilter retry (learn.microsoft.com), pyvbaharness (watchdog pattern)

## Reload relay (2026-10-14)

Claude Code never restarts a stdio MCP server by itself (docs: "Stdio servers are local processes, and Claude Code
doesn't reconnect them automatically"), so every code change needed `/mcp` → Reconnect. A stale server also tempted
side scripts around the tools.

- **Relay.** `run_server.py` is now a small stable relay. The real server is `child_server.py`, run as a child process. JSON-RPC lines
  pass through unchanged.
- **Reload.** Before a request, the relay compares the code fingerprint: mtime and size of `smartview_mcp\*.py`, `vba\*.bas`
  and `child_server.py`. If it changed and nothing is in flight:
  1. Close the child's stdin. The lifespan shuts its Excel and restores Smart View options, with a 60 s grace before the tree is killed.
  2. Start a new child.
  3. Replay the client's `initialize` under id `relay-init-N`, and drop that answer.
  4. Send `notifications/initialized`.
  5. Send `notifications/tools/list_changed` to Claude Code.
- **Crashes.** A crashed child is restarted the same way. Requests it still owed get a JSON-RPC error ("retry").
- **Status.** `sv_status` reports `server_started` and `behind_reload_proxy`. The log is `svkit\logs\proxy.log`.
- **Test.** `tests\live_relay.py excel` is ALL OK. It covers handshake, 20 tools, a real sv_query (Wholesale EBITDA YTD Sep 26 =
  1 858 190 823), no reload without a change, reload plus list_changed after a change, the old Excel closed, crash recovery,
  and a clean exit with no processes left.
- **Limit.** Only an edit to `run_server.py` itself needs a reconnect.

Same day:
- `sv_ws_answer` now saves a file the MCP created in `analyser\` (an earlier `sv_ws_pull` had created it).
- Charts take `{lines, columns: [i, j]}` for one series per column, e.g. this year vs last year per month.
- The value axis is in MNOK through `DisplayUnit=xlMillions (-6)`, because scaling format codes are garbled in the nb-NO locale.

## Performance pass (2026-10-14)

Every tool result now carries a `timings` field with the seconds per stage (`smartview_mcp\timing.py`).

The benchmark is the BUS EBITDA monthly flow: 2 pulls plus a 30-line answer with a chart, into a new workbook in the user's Excel.

| Stage | Before | After | Fix |
|---|---:|---:|---|
| Pull: dimension lookup | 10-20 s, once hung for minutes | 0 | Dimension labels and aliases are cached on disk (`svkit\dims_cache.json`), learned from the pull's own sheet. The hidden server Excel is no longer started for user-Excel jobs (`host.ensure()` on demand). |
| Pull: default grid | 1.6 s | 0 | Skipped when the dimensions are known. Falls back to the default grid on a failed retrieve or a non-standard sheet. |
| Answer: `ensure_styles` (53 styles, about 1 300 property sets, run twice) | 25-160 s | 0.02 s | Version marker name `_tnStylesVer`. New workbooks are created from `svkit\tn_template_<ver>.xltx`; existing ones get one `Styles.Merge`. |
| Answer: cell writes | 1.6-17.6 s | 1.6 s | `fast_excel`: ScreenUpdating off and manual calculation, then the user's settings are restored. |
| Answer: page setup | 14 s | 1.3 s | `PrintCommunication=False`. |
| Answer: row heights | 8.8 s | 0.3 s | One call per distinct height. |
| Answer: lint | 3.3-3.9 s | 0.3 s | Checked against the styles assigned. It still catches text overflow. |
| **Pull** | 11-120+ s | **6-8 s** | EPM retrieve itself is 1.4-2.2 s. |
| **Answer** | 33-214 s | **7 s** | |

The key points are written by the model in chat (free-form) or filled in by Python f-strings (templates). The server never calls an LLM.

Found on the way:
- Copilot (`copilot.exe`) runs its own smartview relay and server, a second instance driving the same Excel.
- The hidden server Excel crashed once in Smart View's `HsAddin.dll` (0xc0000005).

Follow-up the same day:
- **One call.** `sv_ws_build` runs the pulls and the answer together. Specs got shorthands (`repeat`, `sum` patterns,
  chart line patterns), `spec_patch` follow-ups, and server-drafted factual key points (`specs.py`, `tests/test_specs.py`).
  Live, through the relay (`tests/live_build.py`): ALL OK. The build took 42.5 s, 24.9 s of it an SSO-refresh connect. A patch takes 4.3 s, and a spec is 1.6 KB instead of about 30 KB.
- **Turns in the user's Excel.** A machine-wide lock (`excel_turn`) lets several agents use smartview at once.
- **Relay announces `tools.listChanged: true`.** Before this, Claude Code ignored `list_changed`: reloaded code ran, but new tools and parameters
  stayed invisible until `/mcp`. This needs one reconnect.
- **Atlas query.** `descendants()` went from 11-45 s to 0.03 s (CROSS JOIN plus a (connection, parent) index).
- **Default grid restored.** The skipped default grid was reverted after two Excel crashes.
- **Template regression** (Technology, through the registered tools) matched the 13.10 numbers. Times:

  | Template | Before | After |
  |---|---:|---:|
  | P&L | 90-240 s | 40 s, including the daily closed-month lookup |
  | Trend | | 21.5 s |
  | Plan | | 28.8 s |
  | Cost base | | 22 s plus SSO waits |

## Parallel agents and session files (2026-10-14)

- **Per-job worker.** Every user-Excel job runs on its own STA thread (`host.call`).
  - Steps that need the hidden Excel go through `h.on_host`: the bridge build, the fallback for a new file, and the daily closed-month lookup.
  - A timeout aborts only that job's `JobToken`.
- **Turns per Excel process.** Each Excel process has its own lock file, `svkit\user_excel_<pid>.lock`, taken by `resolve_target` and released when the job ends.
  - New analyses take a free Excel, or start one on the TN template (about 25 s, up to `MAX_EXCELS`=4).
  - The same workbook queues.
  - Live-tested with 2 subagents and with 2 relays (`tests\live_sessions.py`).
  - Claude Code runs one agent's write tools sequentially, so parallelism comes from several agents.
- **One file per session.** `session.py` identifies the session by its relay ancestor (pid plus start time, stable across code reloads). The project comes from the relay's working folder.
  - Default `target='session'` gives `analyser\<project>\<date time> <first question>.xlsx`, reused for every question (Q1, Q2 ...) and reopened if closed.
  - `active` is used only on request. `new` makes a new session file.
  - The state is kept in `svkit\sessions.json`.
- **Server instructions** carry the essential rules (cube choice, members, closed month, one-call flow, format) for agents without the skill, e.g. Copilot.
- **Copilot.** Its relay predates the `listChanged` fix. Code updates reach it; new tools appear after one Copilot restart.

## Drill-through (sv_drill, 2026-10-15)

- **Route.** Smart View `B_DrillReports` returns the cell's drill link. Headless Chrome with Telenor SSO opens it, and we replay the page's own `getData` request with offset paging (500 lines per page). Lines go to `svkit\drill.db` and the session workbook.
- **Read-only.** Never call `HypExecuteDrillThroughReport` (it opens the default browser) and never build or alter SQL. Only the page's own body is replayed, with `limit`/`offset` changed.
- **Process safety.** Own Chrome profile `svkit\drill_chrome` plus a lock file (one drill at a time across sessions). Chrome is recognised by `--user-data-dir`; only our Chrome is killed, never the user's Chrome or Edge. It closes after 5 idle minutes.
- **Text and reports.** Mojibake is repaired exactly; letters lost in EPM go through `knowledge\text_fixes.json` plus the user's `%LOCALAPPDATA%\smartass\text_fixes.json` (user entries win; updates never touch it); the rest is reported, never guessed.
- **Docs.** Spec: `specs\2026-10-15-sv-drill-design.md`. Plan: `plans\2026-10-15-sv-drill.md`. Skill reference: `~\.claude\skills\smartass\reference\drill.md`.
