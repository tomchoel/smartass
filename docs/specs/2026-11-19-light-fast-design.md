# smartass 0.3: light and fast (design, 2026-11-19)

Status: **draft for the owner's approval.** Nothing in here is built. Evidence comes from throwaway spikes run live on
2026-11-19 (session scratchpad; not kept).

## Intent

The owner wants the lightest, fastest smartass that gives the same answers: less waiting, and nothing left in the
engine that is only there for historical reasons. Since 19.11 the user's own OAuth token reaches EPM directly
(PR #8), which makes most of the machinery around sign-in, hidden Excel and hidden browser unnecessary. This design
says what replaces it and what is deleted.

Where the time goes. The "today" column for the Excel job comes from the last stage timings the engine logged
(`runs.jsonl`, 2-3 November, the v1 pipeline): **the v2 engine logs no stage timings**, so its real figures are
unknown and phase 0 measures them. The other rows were measured on 19.11.

| Step | Today | Measured with the new way |
|---|---|---|
| Fetch grids from EPM (warm) | 0.2 s | unchanged |
| First grid after a cold start | about 20 s (browser sign-in) | 0.7 s (live since PR #8) |
| Excel job for one answer | about 10 s of a 13 s median (150-400 COM calls) | 1.3-1.5 s incl. a 50 000 x 12 lines sheet |
| Drill lines for one cell (2 839 lines) | 21-28 s cold, 6-9 s warm (hidden browser) | 7.4 s, no browser (3.4 s of it is EPM's own query) |
| Fallback when HTTP fails | about 50 s (hidden Excel + Smart View) | 0.2-0.3 s (Planning REST) |
| Atlas build (member lists) | minutes in hidden Excel, + 2 min for cards | about 33 s over REST, with account type and time balance |

Reliability is the second reason. Since 5 November `server.log` holds 37 failed `fin_analyze` / `fin_refresh` calls;
26 of them were Excel/COM failures (10 "Excel was busy", 11 COM attribute errors such as `Range.Value2` and
`Add.Name`, 5 other COM errors). A job of a few dozen COM calls has far fewer places to fail than one of several
hundred.

What the Excel spike covered: opening the built file, the one copy call, closing it, breaking the link, making the
table, writing the formula block, and for a follow-up the replacement of the old sheets. It did **not** cover finding
the target workbook, the read-back, `_svmeta`, sheet ordering or the save, and it ran in a clean hidden Excel.

## Goals

Goals 1 and 2 are targets to steer by, not merge gates: phase 0 gives the real baseline, and each phase reports
against it. The gate for phase 3 is the Excel job alone (run-log key `excel job._total`, save included), measured in the
owner's real Excel with `scripts/bench_answers.py --rounds 3` in one growing workbook, as the baseline was: a median
of **at most 3 s** for an answer without lines sheets and **at most 5 s** with a lines sheet, in every round (the
baseline: 15.6, 25.8 and 46.0 s). Missing the gate is reported with the stage table, not hidden.

1. A new answer without drill lines is in Excel within **4 s** of the tool call when the server is warm; a cold
   server adds at most 1 s.
2. An answer with one drill of a few thousand lines within **12 s**.
3. No hidden Excel, no VBA, no SSO clicker, no hidden browser, no Playwright. The only Excel smartass touches is the
   user's own.
4. The engine is about a third smaller, and it installs with `uv sync`, `smartass setup` and `smartass login`: no
   Smart View, no "trust access to the VBA project", no Edge policy. 0.3 is verified on the owner's PC only. Rolling
   it out to colleagues is a later step that waits for the owner's choice of OAuth client and a test with a
   colleague's EPM role; it does not hold back any deletion in this design.
5. Same numbers. The test set's golden numbers pass unchanged, and every answer still passes
   `checks.sheet_equals_model`.

## Non-goals

- New analysis features, new block types, changes to the AnalysisSpec v2 contract or to tool names and reply fields.
- A rewrite in another language or framework: the waiting is in Excel and EPM, not in Python.
- Other EPM pods than plan1 (FinMax, FinLite) and stratmodel (FinPlan).
- Caching of answers or drills, and any Smart View refresh feature (owner decisions of 2026-11-03 stand).
- The Excel MCP side (`xl_snapshot`, `xl_revert`, `xl_tidy`): kept as it is.

## Anti-goals (a failure even if it works)

- A number that differs from EPM without the answer saying so.
- A crash, a hang or a dialog in the user's Excel caused by smartass.
- A change to a sheet smartass does not own, or a save of a workbook that is not smartass's own answer file.
- Any request that can write to EPM.
- A token or session key in a log, a reply or a file that is not DPAPI-protected.

## Constraints

- Windows, desktop Excel (Microsoft 365), Python 3.12, the user's own EPM role (no admin rights are assumed).
- Answers go into the user's own Excel, into the session workbook; new files are copies of the labelled seed (a save
  through Excel fails silently without the sensitivity label).
- Used from Claude Code and Copilot through MCP. The Excel task pane and the localhost HTTP API are not used.
- The owner uses it now; colleagues later. Install must stay simple, but no update or release machinery is needed yet.
- Owner decisions 2026-11-19: **the Q sheets belong to smartass** - a follow-up may replace them whole, and the
  user's own notes and formulas live on the user's own sheets, which are never changed. **Nothing of the user's that
  points at a Q sheet needs to survive its replacement** (formulas, names, charts, pivots): such references are not
  re-pointed.

## Approaches

| | What | Trade-off |
|---|---|---|
| **A. Token everywhere, file-built sheets (chosen)** | All EPM traffic over the token; answers built as a file and moved into the open workbook in a handful of COM calls; the hidden Excel, VBA, SSO clicker, browser, pane, HTTP API and updater are deleted. | Largest change. Drill lines keep relying on EPM's undocumented page endpoints - as the browser route does today. |
| B. Speed up in place | Batch the COM writing (20-40 calls), keep the browser for drill and the hidden Excel as fallback. | Lower risk; answers at maybe 4-6 s; every legacy part stays. |
| C. Task pane writes the sheets (Office.js) | `insertWorksheetsFromBase64` from the add-in. | The pane must be open and is not used; open bugs in that API; a localhost bridge to maintain. |
| D. Rewrite in another language | - | Removes no bottleneck; discards 20 000 lines of tests. |

The owner handed the how to the builder; A is the builder's pick, on the measurements above.

## Design

### 1. EPM access: one token session per pod

One small HTTP client per pod (stdlib; a cookie jar, the bearer from `epm/oauth.py`, no redirects followed), used by
everything that talks to EPM. It only sends what an allowlist admits:

| Purpose | Requests |
|---|---|
| Grids, closed month, drill links | `POST /HyperionPlanning/SmartView` with the bodies `svxml.check_read_only` admits (as today) |
| Grid fallback, member lists | `POST .../plantypes/{cube}/exportdataslice`; `GET` under `/HyperionPlanning/rest/v3/applications/` |
| Drill lines | `GET /aif/ui/model/applications/servicetype`, the drill link's `FORMAT=SUI` form, `GET /aif/ui/model/view`; `POST /aif/ui/model/drilldown/getData` |

Anything else is refused before it is sent, as `check_read_only` does today.

**Grids.** The Smart View provider XML over the token stays the first route: it is built, fast, and returns the labels
Excel shows. `BrowserTransport` is deleted; `BearerTransport` loses its fallback argument.

**Second engine: Planning REST `exportdataslice`** (Oracle's documented API, "any role", stateless). It takes over
the role of the Excel fallback (`data/adhoc.py`, `xl/pull_adhoc.py`): a grid the provider route could not fetch is
fetched once over REST, and the reply's `routes` and warning say so. It must return the same `Fetched` shape, so
everything after the fetch is unchanged. Known differences to handle: dimensions are named without "All " and
spaces; the answer carries member names, not labels (labels come from the atlas); large grids are split by a cell
budget and stitched (the Apache-2.0 `pbj-pbcs-client` pages this way; no cell cap is documented, assume 1 M).
A grid is only accepted when its layout check passes, as on the provider route.

**No sign-in.** With no browser there is no silent fallback. A tool call that finds no valid sign-in starts the
device-code login itself and answers with one message: the page, the code, and "ask again when you have approved".
The server finishes the login in the background. `smartass login` stays for the command line.

**Drill lines over the token** (spike: 2 839 lines identical to the browser route, in the same order):

1. `GET /aif/ui/model/applications/servicetype` with a `page_id` header. The response header
   `x-appbuilder-client-id` is a key the server issues per session; every later request sends it back, together
   with `page_id` and `X-Requested-With`. Without it every POST is redirected to `/interop/epmexit.jsp`.
2. The drill link in its `FORMAT=SUI` form gives EPM's own query (as `drill_direct.sui_url` / `parse_sui` today).
3. `GET /aif/ui/model/view?page=DRILL_THROUGH&appId=..&partitionkey=..` gives the column definition.
4. `POST /aif/ui/model/drilldown/getData?limit=..&offset=..` gives the lines.

A redirect to `epmexit.jsp` or a non-JSON answer means the key or session is gone: take a new key once and repeat.
The rest of the drill stays: `drillstep`, `drill_split`, `drill_derive`, `drill_lines`, the one-call page size, the
tie to the cube. EPM answers one SUI query at a time per user; `getData` calls run in parallel. Deleted:
`drill_browser.py`, `drill_tabs.py`, the browser half of `drill_direct.py`, the warm-up, the profile slots, and the
`playwright` dependency.

**Atlas over REST.** `GET .../applications/{app}/dimensions/{dim}/members/{dim}` returns a whole hierarchy in one
call with alias, parent, data storage, account type and time balance (spike: every FinMax and FinPlan dimension in
about 33 s; member counts match `atlas.db`). `smartass atlas build` uses it; account type and time balance go
straight into the cards. The consolidation operator is not in the answer, so `cards_fetch` keeps deriving it from
data (it already runs over HTTP). Oracle documents this endpoint as Service Administrator only, although it answers
for the owner's role on both pods: when it answers 401/403, the build walks the hierarchy with the provider's
`EnumMembers` instead and says so.

**Deleted with this part:** `xl/host.py`'s hidden Excel and `xl/hostproc.py`, `xl/dialogs.py`, `xl/login.py`,
`xl/pull_adhoc.py`, `data/adhoc.py`, the bridge half of `epm/ops.py`, `epm/sso.py`, `engine/sso_watcher.ps1`,
`engine/vba/`, the Smart View global-option handling, and the doctor checks for Excel-as-engine, Smart View, VBA
access, Smart View start-up and the drill browser. The COM job runner for the **user's** Excel (turn-taking, edit
mode, attach) stays: `xl_*` tools and the answer job need it.

### 2. Excel output: build a file, move it in

*Refined 2026-11-19 after phase 0 and a survey of the drawing code (`docs/audit/2026-11-19-render-map.md`).* The
baseline (`docs/audit/2026-11-19-baseline-v2.md`) puts the Excel job at 33.7 s of a 34.9 s answer, about 280 COM
calls, and each call costs 50-70 ms once a workbook holds some 60 sheets (measured in the owner's Excel; it is not
recalculation). The answer's cell contents are already written with one `Range.Formula2` call: the calls go to the
look (named styles, heights, merges, page setup), the Data and Lines sheets, the charts, the Sources log, sheet
handling and the save. So the file carries **the look and the bulk data**, and the contents keep being written in
bulk after the move-in.

`render/layout.py` stays the single description of an answer's contents. New, pure: `render/look.py` gives the look
of every cell of the Answer (what `xl/styling.apply_report` applies today through COM).

1. **Inside the job, read q** from the workbook's sheet names (as today), refuse a sheet or table name that is
   already taken, and **build** a temporary .xlsx with XlsxWriter (outside OneDrive): the Answer sheet with its
   formats, widths, heights, merges, conditional formats, page setup and native charts but **no cell contents**; the
   Data sheet with its header texts and the fact rows; the Lines sheets with their rows. **No Excel Table in the
   file.**
2. **Move in** (screen updating and events off, calculation manual, all restored): open the file read-only and
   hidden, copy its sheets after the question's place in the workbook in **one** `Worksheets([...]).Copy` call,
   close it, break the leftover link.
3. **Finish inside the workbook:** `ListObjects.Add` for `Data_Q<n>` and every `Lines_Q<n>_b<i>`; one `Formula2` for
   the whole Answer block (as today), one for the Data header's formulas, one cell per Lines sheet; calculate the
   new sheets only; one read-back of the block against the model (`checks.sheet_equals_model`); the status cell; the
   Key point rows whose calculated text needs another height.
4. **Workbook-wide parts stay COM and are slimmed:** the Sources log (rows appended in one write, no AutoFit), the
   spec in `_svmeta`, the sheet order, and the save - an explicit `Save` of our own file, with AutoSave paused during
   the job: left to AutoSave, the last answer had not reached the disk 30 s later (live 19.11.2026, commit 5058b03).
   The user's workbook is never saved.

**A follow-up, a refresh or a redraw of Q<n>** replaces its sheets (Answer, Data and every Lines sheet of that
question): rename the old sheets and tables out of the way ("old ..."), move the new ones in and finish them, then
delete the old ones. If any step before the delete fails, the new sheets are removed and the old names restored.
Anything the user typed on a Q sheet, and references from the user's own sheets to it, are not kept (owner decision
above).

**The table rule.** A sheet that contains an Excel Table is never copied between workbooks. Spike: after a second
copy of a table-bearing sheet into the same workbook, replacing sheets crashed EXCEL.EXE (access violation) in 13 of
14 runs. With the tables made inside the workbook, one answer plus six follow-ups in a row passed in a clean Excel
(21 of 21, 1.3-1.5 s each) and in the owner's real Excel with Smart View, think-cell and Arixcel loaded and AutoSave
on (7 of 7, 2.3-3.6 s each, 50 000-line sheet included). The writer asserts that the built file has no table.

**Known differences from today's look** (the owner judges them in the comparison): cells carry direct formats
instead of the "TN ..." named styles. (Per-point label positions on line charts and the lifted label on thin waterfall
steps were on this list; the move-in now sets them by COM after the copy, as the COM drawing did, so they no longer
differ.)

**Deleted with this part:** `xl/owned.py` (owned ranges, backup sheets), the in-place rewrite of the Data table,
COM chart drawing in `render/charts.py`, `xl/drill_xl.write_lines` and the COM writers in `render/excel.py` that the
file replaces. `xl/styling.apply_report` stays: `xl_tidy` uses it on the user's own sheets. `ChartSpec`,
`keypoints`, `layout`, `readback.compare`, `check` and `sheetorder` stay.

### 3. Slimmer product

- **Removed:** `addin/tn-claude` (task pane and its server), `smartass serve` with `ops/httpapi.py`, the HTTP
  secret and its Startup entry, the `smartass xl` sub-commands (the pane's), `ops/update.py` and the release-folder
  settings. The sign-in renewal already runs in the MCP server.
- **Install:** `uv sync`; `smartass setup` (user config, MCP entries for Claude Code and Copilot, skill link);
  `smartass login`, which builds the atlas when there is none (about half a minute, plus the cards). The atlas is no
  longer copied from a release folder; `smartass atlas build` rebuilds it at any time, and a tool call without an
  atlas says which command to run. The `[oauth]` settings move to `engine/defaults.toml` once the owner has settled
  which OAuth client colleagues use.
- **Doctor:** config, EPM reachable, oauth sign-ins, atlas, knowledge, output folder.
- **Dependencies:** `playwright` out, `xlsxwriter` in; `mcp`, `pydantic`, `pywin32`, `psutil`, `openpyxl` stay.
- Version 0.3.0; `docs/DESIGN.md`, the skill's `reference/` and `docs/troubleshooting.md` rewritten to what is
  left.

## Order of work

Each phase is merged on its own (branch, PR, CI green), and **the old route is deleted only after the new one has
matched it live.**

| Phase | Delivers | Proof before the old part is deleted |
|---|---|---|
| 0. Baseline | stage timings logged again for every v2 answer; the test-set questions timed on today's engine | - |
| 1. Drill over the token | drill lines without a browser; sign-in started from the tool call | the test set's drill questions and three live cells (one over 5 000 lines, one split cell, one "too wide"): same lines, same tie to the cube as the browser route |
| 2. REST second engine, atlas over REST | fallback without Excel; `atlas build` in under a minute | 20 grids from the run log fetched on both routes: identical `Fetched`; a rebuilt atlas equal to today's apart from members added since |
| 3. File-built answers | the Excel job within the gate above | every test-set question drawn both ways: same cell values, formulas and formats on the Answer sheet, charts compared by eye once; six follow-ups in a row in the owner's real Excel with Smart View loaded |
| 4. Slim down | pane, HTTP API, updater and hidden-Excel leftovers gone; install and docs for 0.3.0 | `smartass doctor` clean on a fresh checkout; the whole suite and the test set green |

Phases 1 and 2 do not depend on phase 3 and can be built in parallel.

## Verification

- **Numbers:** the offline test set (9 questions, golden numbers) unchanged; `checks.sheet_equals_model` on every
  answer.
- **Speed:** the goals above read from `runs.jsonl` stage timings on the owner's PC, before and after each phase.
- **Excel stability:** an automated live check that draws an answer and replaces it six times in one Excel session,
  run before every release; plus a unit test that the built file has no table.
- **Read-only:** tests that the HTTP client refuses every path and method outside the allowlist.

## Risks

| Risk | Handling |
|---|---|
| Oracle changes the undocumented endpoints (provider XML, drill page) | Grids fall back to documented REST. Drill has no documented alternative: it fails with a clear message, as the browser route would. |
| The Excel findings come from one Excel build, in a clean hidden instance | Phase 3 starts with the six-follow-ups check in the owner's real Excel (Smart View and other add-ins loaded), before anything else is built on it. If it fails there, phase 3 falls back to approach B's batched writes behind the same plan. |
| The temp workbook flashes or steals focus when it is opened | Screen updating off, window hidden, active sheet and selection restored; judged by the owner in phase 3's first demo. |
| The spike's drill view definition differed from the browser's (not examined) | First task of phase 1: explain the difference before relying on it. |
| A colleague's role may not read member metadata over REST | `EnumMembers` walk as fallback; checked when colleagues are brought in. |
| An unused sign-in ends after 7 days | Renewed daily while a server runs; otherwise the tool call starts a new login. An identity-domain admin can raise the lifetime. |

## Left to the builder

File and module names; how the user-Excel job runner is cut loose from `ExcelHost`; retry and timeout values; how
closely charts must match pixel for pixel (same type, series, colours, labels and title is required); whether the
TN look is written as direct formats or named styles, as long as `xl_tidy` reports no BLOCK finding on a fresh
answer; the paging budget of the REST engine; test layout.

## The builder's own calls (not decided by the owner)

- Python and the current model, knowledge layer, tool names and reply fields stay.
- XlsxWriter builds the file (formula and chart support; fastest measured for the lines sheet: 2.4 s for 50 000 x
  12 against 3.7 s for one `Range.Value2` call).
- REST `exportdataslice` is only a fallback, not the first route.
- A missing sign-in starts the login from the tool call instead of failing.
- The pane's `smartass xl` sub-commands go with the pane.
- The atlas is built at the first login instead of copied from a release folder.
- Speed goals are targets; the only timing gate is the Excel job in phase 3.
- Phases 0 to 4 in the order above.
