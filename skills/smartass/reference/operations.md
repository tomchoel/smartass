# Operating the smartass MCP (v2)

## How it runs

- **One MCP entry, `smartass`.** Claude Code starts `engine\run_server.py`, a small relay that runs the real
  server `engine\child_server.py` (`smartass.server`) as a child process and passes JSON-RPC through.
- **Code changes roll only on request.** The relay replaces its child when `engine\VERSION` changes
  (`scripts\bump.ps1` bumps it together with `pyproject.toml`) or when a client calls `fin_status(restart=True)`.
  Editing a `.py` file alone does nothing to a running server. After a reload, Claude Code may keep an old tool
  schema: run `/mcp` and reconnect smartass. Only an edit to `run_server.py` itself always needs a reconnect.
  After an update, reconnect every open session: one still on 0.4.0 in the same Excel has no first-contact check, so
  a 0.5.0 save after the reply can run in the middle of its job.
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
  server's environment sets the seconds; 0: no pause). With the add-in, the answer's own file is saved by Excel
  right after the reply (0.4-1.1 s the user no longer waits for): the job asks for the save in one call, Excel waits
  0.3 s so the answer is painted first (`SV_SAVE_DELAY`), saves, and switches AutoSave back on. A follow-up that
  comes before that save ran has it run first; one that comes while it runs waits for it. Over COM, with
  `SV_SAVE=sync` in the server's environment, and for the first answer of a workbook the question created, the job
  saves before it replies (a new workbook's file holds its first answer before the session points at it). Never the
  user's own workbook. Two accepted risks for a later answer: a save after the reply that fails while no further
  answer comes is shown only by `fin_status` (`save_after_reply`; the next answer would say it and save); and if
  Excel crashes after the reply before that save ran (0.3 s plus the save's own time; longer while a cell is being
  edited, as the save waits until Excel is ready), that answer is lost without a note - the file on disk is the
  workbook as it was before that answer. The Excel job takes about 2-2.5 s for an answer without a lines sheet (2.7 s at
  61 sheets), about 4 s with a lines sheet, 1.6-1.9 s for a follow-up, and 5.3-6.1 s for the first answer of a new
  workbook (live, a fresh Excel, 20.11.2026, before the cuts of phase 3c Task 1b; 0.4.0 took 7.5 s there).
  After those cuts: (controller fills in). A follow-up or `fin_refresh` replaces the question's `Q<n>`
  sheets and tables whole: anything typed on them goes with them, so keep your own notes on your own sheets.
  A formula elsewhere in the workbook that points at a `Q<n>` sheet or its table (`='Q3 Answer'!E12`,
  `=SUM(Data_Q3[Value])`, a cross-sheet SUM, an array formula, a chart series, a defined name; also on another
  answer's sheet) is re-pointed at the new sheets before the old ones are dropped, so it shows the new answer's
  value: the same cell address, so a cell the new layout no longer has shows what is there now, or #REF! when the
  new answer has no such sheet. The reply warns about a formula Excel refused (a locked cell on a protected sheet, a
  table column the new answer lacks, a legacy array formula longer than 255 characters), about one the user changed
  while the answer was drawn (each write first checks the formula is still what was read; an edited cell, or a row,
  chart or name added in between, is left alone), and about a sheet whose formulas cannot be read (a protected
  sheet with a hidden formula, or one too large to scan) when a cell there shows #REF! after the follow-up that did
  not before; always with Calculation manual, and for a sheet too large to check as well (a used range of more than
  about 100-200 million cells, such as one a stray format stretched to the last row and column) - those show #REF!.
  When the workbook cannot be scanned at all, the old sheets are dropped anyway and the reply says the formulas pointing at them were not checked. **Not re-pointed** (they show #REF!): conditional
  formats, data validation lists, pivot tables built on a `Data_Q<n>` table, chart titles and labels linked to a
  cell, sparklines, form controls' linked cells and lists, shapes and text boxes linked to a cell, and links from
  other workbooks. Nothing of the user's is saved and Excel's settings are put back (also those an earlier job that
  did not finish left changed: the next job in that Excel puts back what it recorded in its state file, `jobstate\`,
  and says so); when the answer's workbook is the active one it ends on the new Answer, otherwise the user's active
  sheet is put back.
- **The add-in.** The first answer in an Excel builds a small Excel-DNA add-in (`smartass_run_<hash>.xll`, compiled
  with the .NET compiler that ships with Windows, under `%LOCALAPPDATA%\smartass\inproc`) and loads it into that
  Excel with `RegisterXLL`, silently, for the rest of that Excel's life; each batch of an answer's Excel work is
  then one `Application.Run` into it. Anything that keeps it from loading or answering (no compiler, a 32-bit
  Excel, a policy, a dialog: `RegisterXLL` is given 60 s) makes that Excel use the COM route instead - the same
  work, about twice as slow for a follow-up (4.9-6.2 s against 2.5-2.9 s in-process, measured while another session
  worked in the same Excel) - with one warning in `logs\server.log` ("in-process executor not used in Excel pid
  ..."). On a very large workbook the COM route costs more: every follow-up reads the defined names one call at a
  time (3-5 s for 2 000 names), and a sheet whose used range reaches both the last row and the last column costs some
  700 000 calls (an estimate, not measured: minutes); `fin_refresh` pays both once per question.
  `SV_INPROC=0` in the server's environment forces the COM route; such a server does not run another session's
  waiting save after the reply at its first contact, so that save can run between its batches. The add-in runs
  nothing but job files the engine writes to `inproc\jobs` (a formula calling it gets #N/A); the user's settings it
  turns off while it works are put back by each batch itself, also when a step fails.
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
| The outcome of each Excel's last save after a reply (no figures; `last.json`: the last one a job read) | `inproc\saved\` |
| The Excel settings a running job changed, for the next job in that Excel if it dies (no figures; removed at the job's end) | `jobstate\<Excel's pid>.json` |

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
| Warning in `server.log`: `in-process executor not used in Excel pid <n>: <why>` | The add-in did not load in that Excel; answers run over COM (about twice as slow for a follow-up, same result). The reason says what (no compiler, 32-bit Excel, `RegisterXLL` refused or blocked). A load that failed is tried again in a new Excel or after a server restart. A compile error is kept in `inproc\smartass_run_<hash>.failed` and not tried again for a day (delete the file to try now); an update that changes the add-in is built anew. |
| `Excel operation None failed: the batch in Excel: the batch's outcome is unknown: …` | The add-in took one batch of the answer's Excel work, and what it did cannot be told (its answer could not be read, or the call failed after it took the job). The batch is not run again; the job is undone as far as the workbook shows (the old answer back, or the new one kept and said). Check the workbook, then retry once. If it repeats, every batch in that Excel fails the same way: restart that Excel, or set `SV_INPROC=0` in the server's environment (read when the server starts). |
| `…; the new answer is in place (the old one is gone)` (`and saved`) | A follow-up or `fin_refresh` failed after its old sheets were dropped, or its final save failed: the new answer stays (a rollback would leave neither; `and saved`: our own analysis file was saved, the user's workbook never is). The Sources rows and the stored spec may be missing: run the same follow-up again to record them. |
| Warning `the save of <file> after the last answer failed (<Excel's message>); it is saved now` | Excel's save of the answer's file after the previous reply failed (the file was locked, the network was gone); this answer saved it in the job. Nothing to do unless it repeats: then set `SV_SAVE=sync` and report it. `... could not be saved now either`: save that workbook yourself. `fin_status` shows the last outcome (`save_after_reply`). |
| Warning `an earlier smartass job in this Excel did not finish and left <settings>: put back now` | A job died (a killed server, a put-back Excel refused) and left Excel's settings changed; this answer put them back from its state file (`jobstate\`). Nothing to do. |
| Warning `AutoSave could not be switched back on in <file> ...` | Excel refused AutoSave on in the answer's file. `the next answer in this Excel switches it on` / `this answer switches it on` / `it is switched on now`: nothing to do; `switch it on yourself`: switch AutoSave on in that workbook. |
| Warning `<file> was closed before smartass saved it after the last answer` | The workbook was closed in the moment between the reply and Excel's save of it; Excel asked whether to save it when it closed. Open it and check the last answer is there. |
| Repeated `HostError` | `fin_status(restart=True)`, then retry; if it repeats, `smartass doctor --bundle` to the maintainer. |
| HTTP `409 busy` / `401` / `413` | `smartass serve`: one request at a time / wrong secret / body over 1 MB. |

Drill errors (sign-in, blank drill page, too wide, no drill report): [drill.md](drill.md).
Doctor checks one by one: `docs\troubleshooting.md` in the repo.
