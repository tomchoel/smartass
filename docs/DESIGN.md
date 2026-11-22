# smartass: design as built (v2, 2026-11)

The target design and its reasoning: [`specs/2026-11-04-v2-overview.md`](specs/2026-11-04-v2-overview.md) and
[`specs/2026-11-04-v2-1-engine-design.md`](specs/2026-11-04-v2-1-engine-design.md). This file says what was built.
v1 (the `sv_*` gate/recipe pipeline, retired 2026-11-04) is described in
[`archive/DESIGN-2026-11-v1.md`](archive/DESIGN-2026-11-v1.md); the log up to October 2026 in
[`archive/DESIGN-2026-10.md`](archive/DESIGN-2026-10.md).

## Shape

- **The model interprets, the engine validates.** The finance agent writes an **AnalysisSpec v2** (plan, queries,
  typed blocks). The engine checks it, fetches, computes every value in Python and draws it in Excel. There are no
  recipes and no regex gate.
- **One model, two outputs.** Every output cell is an expression tree. `evaluate(frame)` gives the number;
  `to_excel` gives the formula (SUMIFS / XLOOKUP on labels in table `Data_Q<n>`). The sheet is read back once and
  must equal the model (`checks.sheet_equals_model`).
- **Read-only towards EPM.** Every request to EPM that carries your sign-in goes through one HTTPS boundary,
  `epm/epmhttp.py` `EpmSession`, whose allowlist admits reads only (Runtime, below); the doctor's `epm` check alone
  is a plain GET of `epm.base` without a token, to see the server answers. No submit, calculate, import, job or
  business-rule request exists in the engine.
- **Your Excel is never saved, and closed in one case only.** The engine has no Excel of its own (since 0.7.0). It
  starts an Excel only as yours, visible, on a workbook (Runtime), and saves only its own answer files. It quits an
  Excel only when it started that Excel itself for a new workbook that could not be saved and nothing else is open
  in it (`target._discard_new`).

## Tools (9, `smartass/server.py`)

| Tool | Does |
|---|---|
| `fin_analyze(spec \| patch, q, target)` | New question from a spec, or a JSON merge patch on the stored spec of `Q<n>` |
| `fin_refresh(target)` | Re-runs every v2 answer in the workbook; relative periods roll to the latest closed month |
| `fin_lookup(query, kind, cube)` | Metrics by name or alias (description, members, `differs`), atlas members and product concepts (`kind="concept"`), with caveats; `kind="answer"`: past answers from answer memory |
| `fin_drill(cell, group_by, top)` | A spec with one `lines` block: FinLite ledger lines ranked and tied to the cube |
| `fin_critic(q, scores, notes)` | The stored answer's critic brief + the rubric (`knowledge/agent/critic.md`); with `scores` (six criterion scores 1–5) computes the score (mean rounded half up) and stores it |
| `fin_status(restart)` | Version, atlas age, knowledge build and refresh due, doctor problems, `save_after_reply` (the newest outcome of a save Excel did after a reply, and how many are still to come); `restart=True` gives a fresh server process (without the relay: new EPM sessions and a new job runner) |
| `xl_snapshot`, `xl_revert`, `xl_tidy` | Snapshot / undo around Excel MCP edits; TN style + tn-formelstil check |

Every other Excel job (read, write, pivots, Power Query, charts) is Excel MCP. Error text: every tool runs under
`core/serverkit.guard`, which turns any exception into one readable message with one next step.

## Pipeline (`pipeline.py`; the Excel job in `exceljob.py`)

1. **Prepare** (no Excel): validate the spec (`model/spec.py`, pydantic, strict), resolve metrics and members
   (`knowledge/resolver.py` `KnowledgeResolver` over `knowledge.db` + `atlas.db`), run the hard guards
   (`model/validate.py`: unknown member or metric, open month without `preliminary`, period mismatch), plan the
   grid requests (`model/requests.py`, `model/series.py`) and fetch them in parallel over Smart View HTTP
   (`data/provider.py` → `epm/svhttp.py`). `lines` blocks drill first (`epm/drillstep.py`, concurrently).
2. **Second engine** (`data/provider.py`): a grid request the Smart View provider route could not fetch is fetched
   once over Oracle's Planning REST `exportdataslice` (`epm/rest.py`, through the same `EpmSession` of the pod),
   parsed by the same `facts_from_parsed`. Not for an error REST meets as well (`svhttp.NoFallback`: no token, a
   token refused twice, a pod without a scope, a stale atlas, an alias clash, the data-source lookup busy) or a
   missing sign-in. A member the cube does not have where the grid puts it (the provider's `WrongMember`; REST's
   HTTP 400 names it) is left out of the request when it is on an axis of several (at most 3, every grid of the
   request): it gets no fact, the path of a missing cell, and one warning and one warn caveat per member (the
   Answer's Caveats, the Sources sheet's Known limits) say its figures are missing and every total with it is
   INCOMPLETE. A POV member, an axis's only member or a fourth refused member fails that request with an error that
   names it, the grid and `smartass atlas build`: the grid is not fetched, and like any grid neither engine could
   fetch it is a warning (`Prepared.failed`) and the answer is drawn without its numbers. No Excel is asked for a
   grid. `routes` counts `rest`, and one warning says how many grids came over REST and that a cell you may
   not read reads as empty there (Oracle exports only cells with read-write access). Live gate:
   `docs/audit/2026-11-21-rest-vs-provider.md`.
3. **Compute** (`model/analyze.py`): facts → `model/frame.py` → block builders (`model/blocks/*`) → expression
   trees (`model/expr.py`) → values, findings, material lines and `key_values` (`model/findings.py`).
4. **Excel job** (one job in the user's Excel, `exceljob.excel_job`): the answer's sheets are built as one .xlsx
   without Excel (`render/xlsx.py`, charts `render/xlsx_charts.py` on the rules of `render/charts.py`) from a pure
   `render/layout.py` plan - `Q<n> Answer` (its look: `render/look.py`), `Q<n> Data`, `Q<n> Lines <alias>` for
   drills - under `%LOCALAPPDATA%\smartass\tmp`. Its Excel work is recorded as **batches** of operations
   (`xl/inproc/batch.py`: get / call / set on named objects, the verbs `each` for the only loop, `nudge` and
   `exist`, a `finally` list), and each batch runs in one go inside Excel (`xl/inproc/runner.py`: one
   `Application.Run` into a small Excel-DNA add-in, `Runner.cs`, built on first use and loaded once per Excel
   process) or, where the add-in cannot load or `SV_INPROC=0`, over COM from Python - the same operations, the
   same workbook. Three batches and a save (`exceljob._file_draw`): **look** (every read the build needs: the
   sheets and tables, the base, which workbook is active, the kind of a Sources sheet, the stored specs' keys; then
   AutoSave off while the job writes); **bring in** (the answer's table names looked up among the defined names
   first, a lookup per name - a new answer's number comes from the look - then a follow-up's old sheets parked, the
   file's sheets inserted with one `Sheets.Add(Type:=<the file's path>)` - no workbook opened, nothing copied, no
   link left; proved equal to the Open + Copy + Close it replaced by `scripts/check_insert.py` - and moved to their
   places, then `xl/movein.py`'s finish: the tables `Data_Q<n>` / `Lines_Q<n>_<alias>`, the formulas, the chart
   labels' position, calculate, the labels' lift, and the reads the close-out needs: the read-back block, the status
   cells; the Answer active);
   **close out** (the 'Sheet = model' status (`render/readback.py`), Key point heights and page breaks, a
   follow-up's old sheets dropped, the Sources log (`xl/sources.py`), the spec in hidden `_svmeta`, sheet order
   (`xl/sheetorder.py`)); then the save (our files only, AutoSave back on): with the add-in Excel does it right
   after the reply (`xl/inproc/later.py`: one Run asks for it; the add-in waits `SV_SAVE_DELAY`, 0.3 s, so the
   answer is painted, then saves as a queued macro and writes an outcome the next job reads; a batch that comes
   first runs it first, one that comes while it runs is answered "busy" and asked again; a failed one is a note of
   the next reply, which saves in the job), over COM, with `SV_SAVE=sync` and for the first answer of a workbook the
   request created (`Target.new`: its file holds the answer before the session points at it) in the job. Every
   quiet batch puts the user's settings back itself, and what a job that did not finish left changed is put back by
   the next: right after its look, before its first quiet batch, a job records in the state file
   `%LOCALAPPDATA%\smartass\jobstate\<Excel's pid>.json` (`xl/jobstate.py`: that Excel's start time, the four
   quiet settings as found, what it sets, our file whose AutoSave the look switched off; no figure) and removes it at
   its end; the next job in that Excel process, at its first contact, puts back each setting that still reads what
   the dead job set and says so, and keeps only what Excel refused to put back for the job after it. Calculation is
   set to manual only where formulas are written - the bring-in, a close-out that re-points formulas elsewhere, and the point-back after a failure - since putting automatic back recalculates
   what is dirty in every open workbook; after the bring-in the job pauses once (`SV_PAINT_PAUSE`, 0.15 s) so Excel
   paints the answer with its numbers and charts when the workbook is on screen. A failure rolls back as before
   (the clean-up runs as batches of its own, also for a job abandoned by a timeout).
   **The first answer of a new workbook** (target `new`, or `session` before the session has one, in a running
   Excel) builds the workbook itself: `xl/target.py` chooses the Excel as for any new workbook (a free turn first)
   and only reserves a free file name in the analysis folder (created empty and exclusively, so a parallel job takes
   the next; never the name of a workbook open in that Excel); `render/xlsx.build(book=NewBook)` writes the
   answer's sheets, then the Sources log furnished as `xl/sources.py` furnishes it over COM (the same data), the
   sensitivity label as custom document properties (`target.label_properties`) and the TN template's theme, in the
   base a workbook added from the template has (`target.new_base`, stored by the look of the first one made that
   way); Excel only opens it (`Workbooks.Open`, stage `draw.open`) and the job goes on with the look, the bring-in
   (the finish only: no name check, park or insert), the close-out (the log's rows: its furniture and print setup
   came with the file) and the save, in the job - no `Workbooks.Add`, no `SaveAs`, no furniture in Excel. A
   failure once the name is reserved closes the workbook without saving and deletes that file (only it, found by its
   full path), and the error says so; the workbook becomes the session's once the answer is in it and saved. A new
   workbook is still made from the template (label, SaveAs, insert) without a label id, before a base is stored,
   without the template, and in an Excel started for a parallel agent. With no Excel running (`_new_target`): a copy
   of the labelled seed when there is one (found also when Excel keeps the label in `docMetadata/LabelInfo.xml`, as
   it does on this tenant); else, with a stored base or without the template, a "Sources" workbook built without
   Excel (`render/xlsx.build_sources`: the label as document properties, the TN theme), kept as the seed only when
   it was built in the stored base (written to a temp file, then replaced: two cold jobs can write it at once; a
   stored base whose column width is not a standard one is not used, so with such a base every cold start builds a
   plain file and keeps none - an open item); else your Excel launched on the TN template and saved as
   (`_spawn_excel`, `_save_new`), whose look stores the base.
   A seed or a built file is opened by launching Excel on it (`_open_path`); none of this needs an Excel of the
   engine's own.
   The built workbook has no TN cell styles until something needs them (`styling.ensure_styles` merges them
   from the template: `xl_tidy`, `apply_report`). Proved item by item against one made today's way by
   `scripts/check_insert.py --parts newbook` (20.11.2026, then in the engine's hidden Excel of 0.6.0; since 0.7.0 the
   script starts a hidden Excel of its own).
   Measured (docs/audit/2026-11-20-inproc-bench.md, live): 1.9-2.7 s in the Excel job for an answer without a
   lines sheet as the workbook grows to 61 sheets, 3.9-4.3 s with a lines sheet, a follow-up 1.6-1.9 s, at 5
   batches per answer (4 since the review fixes). What is left is mostly Excel's own work: the insert (about 0.5 s)
   and the save of our OneDrive file. The first answer of a new workbook, built as its file (phase 3c): 2.0-3.4 s in
   the Excel job in a script's own hidden Excel on the OneDrive folder, against 8.0-11.6 s made from the template
   there (the Open itself 1.0-1.8 s; 20.11.2026); 10.6-11.3 s made from the template in the owner's Excel (0.4.0 run
   log). Measured in the owner's Excel built as its file (a fresh Excel, 20.11.2026 22:16, before the save after the
   reply): the first answer of a new workbook 5.3-6.1 s against 7.5 s made from the template by 0.4.0 in the same
   Excel (goal 5 s or under): open target 2.1-2.6 s, of which "new workbook" 2.0 s with no Excel in it; `draw.open`
   1.05-1.08 s (Excel opening the file: Excel's own); look 0.83-1.14 s (0.03 s in an existing workbook);
   everything else 0.9 s together. Since then (phase 3c Task 1b; the live figure after it is owed) the session's
   parents are read from one snapshot of the process table instead of one per ancestor: 1.5 s of "new workbook",
   once per server process, measured 1.75-2.15 s before and 0.20-0.22 s after in a plain script. And a server process
   that meets an Excel holding its build of the add-in uses it as it is instead of registering it again: in a script's
   own hidden Excel the look of a new server's first answer went from 0.45-0.49 s to 0.04-0.09 s. That this was most
   of the live look is inferred, not measured live. Ledger and figures: docs/audit/2026-11-20-first-answer.md.
5. **Reply**: `q`, `workbook`, `sheets`, `key_values` (the only numbers the agent may quote), `material`,
   `findings`, `checks`, `layout` (layout check findings with fixes), `charts`, `warnings`, `caveats`, `routes`,
   `closed`, `saved_user_workbook: false`, and for depth standard or deep `critic_brief` (`model/critic.py`). The run is stored in answer memory
   (`answers.py`) at the end of the Excel job and of every redraw.

**Follow-ups and packs** (`followup.py`): a patch is an RFC 7386 merge on the stored spec; `exceljob.redraw_job`
replaces the question's sheets whole: the old ones are parked under 'old ...' names, the new answer is moved in where
they were, and the old ones are dropped. Cells a user typed on a `Q<n>` sheet go with it (their own notes belong on
their own sheets). A formula elsewhere in the workbook that points at the question's sheets or tables follows the park
to the 'old ...' names (Excel rewrites it) and would turn #REF! with the drop: the bring-in batch finds such formulas
(the verb `mentions`, inside Excel: every other sheet's used range halved, its longer side first, into blocks of at
most 100 000 cells, only the blocks SpecialCells/Intersect finds a formula in halved further and read - a protected
sheet block by block, a hidden formula there makes it "could not be read" (the mentions found before stay), and so
does a sheet that would take more than 2 000 steps, a step being one test or one read of a block: a protected sheet
of more than about 100-200 million cells (a used range to the last row and some 100-190 columns wide), or formulas
spread over too many blocks - on a sheet whose used range a stray format stretched to XFD1048576 (blocks of about
256 rows by 390 columns) a formula column taller than about 85 000 rows; such a sheet is warned about on every
follow-up when it is too large to check as well (the check below reads at most 2 000 blocks: a used range of more
than about 100-200 million cells, a stretched one included), else when a cell there turns #REF! -, no Range.Find,
which would leave its options as the Find dialog's and skips filtered rows; the charts' series; the defined names),
Python rewrites exactly the references to a parked name with a namesake in the new answer (`xl/repoint.py`: never
text in a string literal, another workbook's sheet, or a longer name that contains it), and the close-out writes them right before the drop with Calculation manual, each a compare-and-set
inside Excel (the verb `swap`: a formula changed since the bring-in read it is left alone and reported) that may
fail (a warning). A sheet that could not be read is a warning only when a cell there shows #REF! after the drop that
did not before (its values can be read: where its cells show #REF! - the error told by its code, not its text in
Excel's language - is read before and after the drop, bounded as the scan is, and a cell that shows it only after is
said; with the user's Calculation manual, always, as that cannot be told without calculating their sheet; and
always for a sheet whose values are too many to read that way (the verb `errors` runs only for a sheet that could
not be scanned). A failure before the drop points them back first, by the same compare-and-set (the parked sheets
must still be there), then the rollback and the restore. Not covered (Excel offers no exact, cheap way): conditional
formats and data validation (their formulas come only in the user's separators and change through a type-specific
Modify), pivot caches, linked chart titles and labels, sparklines, form controls' links, linked shapes, other
workbooks. A failed redraw removes the new sheets and gives the old ones their
names back; once the old ones
are dropped the new answer stays and the error says so. `fin_refresh` redraws every stored v2 question; one failure
never stops the rest.

## Packages

| Package | Contents |
|---|---|
| `core/` | `config` (paths, settings, DPAPI for the stored sign-ins), `logsetup`, `timing` (stages, progress), `session` (session workbooks, `sessions.json`), `serverkit` (guard, Excel jobs, lifespan) |
| `model/` | spec, periods, requests, series, resolver interface, frame, expressions, blocks, validate, findings, lines (pure; no Excel, no network) |
| `data/` | the provider (the Smart View provider route, Planning REST as its second engine), recorder (offline replay for the test set) |
| `render/` | layout and look (pure), the built file (`xlsx`, `xlsx_charts`), chart rules, read-back, layout check |
| `epm/` | the HTTPS boundary (`epmhttp`), OAuth sign-in (`oauth`), Smart View's provider XML (`svhttp`, `svxml`), grids, Planning REST (`rest`: the second grid engine; `members`: the atlas's trees), the drill over the token (`drill_token`, `drillcore`, `drillstep`) and its lines |
| `knowledge/` | atlas (SQLite), semantic files (`semantic`, `seed`), `build` (knowledge.db), `KnowledgeResolver`, monthly `refresh` |
| `xl/` | the user's Excel (`userjobs`: the job runner; `hostproc`: a job's COM thread and the sweep of an older release's Excels; `jobs`: turns; `instances`, `target`, `dialogs`: Smart View's Login guard; `sheets`), `styling`, `formulas`, `movein`, `sources`, `sheetorder`, `inproc` (batches of Excel operations, their COM executor and the in-process add-in), `drill_xl`, snapshots (`workbench/history`), vendored TN style and formula checker |
| `ops/` | `cli`, `doctor`, `setup_user`, `login`, `atlas_build`, `pane_leftovers` (an older release's task pane in Office's settings, read only) |
| `testset/` | the B-question test set: specs in git, recordings and golden numbers outside git |

## Runtime

- **Relay** (`engine/run_server.py`, the MCP entry `smartass`): runs `child_server.py` as a child and passes
  JSON-RPC through. It replaces the child when `engine/VERSION` changes (`scripts/bump.ps1`) or on
  `fin_status(restart=True)`, and when the child crashed; a `.py` edit alone never restarts it. Log: `proxy.log`.
- **Bearer route** (`epm/oauth.py`; required since 0.6.0): the user's OAuth access token carries both the grids
  and the drill lines over plain HTTPS; smartass starts no browser. Grids: `svhttp.BearerTransport` posts the Smart
  View provider XML through the pod's `EpmSession`. Drill lines: `drill_token` on an `EpmSession` (the drill's own
  rules) takes a session key, asks EPM for the drill query, then pages
  getData; at most `DRILL_PARALLEL` getData calls are in flight, and a lost session is renewed once for all pieces
  (draining the requests still out with the old key). One device-code sign-in per pod (`smartass login`, or the page
  and code a tool reply shows when the sign-in is missing), stored DPAPI-protected; the refresh token is single use,
  so every refresh runs under a cross-process lock, and the servers renew it daily. Without `[oauth]` in
  config.toml a tool call answers with the set-up text. The pod of `epm.base` must have a scope: the closed month and
  the data-source list are read there with no other route, so without it no answer works (doctor FAIL); a grid on
  another pod without a scope cannot be fetched (a warning that names the scope to add). Live 19.11.2026: grids, the
  closed month and drill links are the old browser route's cell for cell (first grid 0.7 s against about 20 s);
  21.11.2026: drill lines on four cells are the browser route's row for row
  (docs/audit/2026-11-19-drill-token-vs-browser.md).
  The same token opens the documented Planning REST API: `exportdataslice` is the grids' second engine and Get
  Member builds the atlas (below). The drill link (GetDrillThroughReports) has no second route: Planning REST has
  no drill-report request, so a link the provider route cannot fetch is an error.
- **One HTTPS boundary** (`epm/epmhttp.py` `EpmSession`, one per pod and route). Every request to EPM with the
  token goes through it (the doctor's `epm` check alone is a plain GET of `epm.base`, without a token or a cookie,
  to see the server answers), and it refuses before a token is asked for anything that is not exactly on its
  allowlist (`ALLOWED`): the method, the whole path (no pattern), the query names and values that path takes, the header
  names it takes, and for a POST the body. The bearer goes only over https to the session's own pod, and only when
  `[oauth.scopes]` has a scope for it; the URL is built from the pod's origin, never from a host a caller wrote; a
  redirect is never followed; a 401 takes one new token and the request goes once more (the grids, Planning REST and the
  atlas's reads: a gate answer, HTTP 301, 302, 303, 307, 401 or 403); an answer is read up to 64 MB; a token, a session
  key, a cookie or a URL query is never in a log line, an error text or a reply (`drillcore.redact` on every URL shown;
  `logsetup.redact` on every log line, tool error and run-log line, the provider's own error text and the two reply
  warnings that quote an engine's error). The rules, one by one:
  - `POST /HyperionPlanning/SmartView`: the grids' provider route. No query, `Content-Type: text/xml` only, a body
    of at most 16 MB that `svxml.check_read_only` admits (the provider XML's read requests only; no CDATA, comment,
    DOCTYPE or processing instruction, which the provider's parser would read otherwise than the check).
  - The drill's own four (`drill_token`): `GET /aif/ui/model/applications/servicetype` (no query);
    `GET .../FdmeeMain.jsp` with `FORMAT=SUI`, `module=aif.drilldown` and the link's own parameters;
    `GET /aif/ui/model/view` with `page=DRILL_THROUGH`; `POST /aif/ui/model/drilldown/getData` with `limit` and
    `offset`, a JSON object with no keys but the drill query's (`QUERY`, `PARAMS`, `PARAMSTYPE`), whose `QUERY` is
    one read-only SELECT (or WITH ... SELECT: no `;`, no statement word that writes or administers). Only these take
    the drill's headers.
  - `POST /HyperionPlanning/rest/v3/applications/{app}/plantypes/{cube}/exportdataslice` for FinMax/FinMax,
    FinMax/FinLite and FinPlan/FinPlan: the second grid engine. No query, `Content-Type: application/json` only, a
    body of exactly Oracle's grid keys, members by name (no member function).
  - `GET .../plantypes/{cube}/dimensions` for the same three, with exactly `aliasTableName=Default`: the plan
    type's dimension names, which `rest` maps the atlas's names onto.
  - `GET .../applications/{app}/dimensions/{dim}/members/{dim}` for 22 named dimensions (14 FinMax, 8 FinPlan),
    with exactly `fields=name,alias,dataStorage,usedIn,accountType,timeBalance,children`: Oracle's Get Member, the
    atlas's trees. No header, no body.

  Nothing that imports, clears, runs a rule or a job is on the list; `tests/test_epmhttp.py` attacks each rule.
- **Two grid engines** (Pipeline, step 2): the provider route first; a grid it could not fetch goes once over
  Planning REST with the same token, but not for an error REST would meet too (`svhttp.NoFallback`). A member EPM
  refuses is left out of a grid when its axis has others, with one warning and one caveat that every total with it
  is INCOMPLETE (at most 3 per request); a refused POV member or only member stops that grid: it is not fetched, the
  reply warns naming the member, and the answer is drawn without its numbers. No Excel is asked for a grid. A REST
  cell is read only as a plain number (digits, one `.`, an exponent); any other text stops that grid with an error
  that names its shape, never its value. The first route's error is classified on the provider's raw text; only what
  is shown is scrubbed.
- **Atlas** (`atlas.db`, `knowledge/atlas.py`): `smartass atlas build` (`ops/atlas_build.py`) reads every
  dimension tree with Get Member (`epm/members.py`; live 8-12 s of Get Member for both pods and 20-24 s for the
  build in all over four gate runs, the last 8.8 s and 23.6 s: docs/audit/2026-11-22-atlas-rest.md) into a NEW file,
  a copy of the atlas, and swaps it in (SQLite's backup, one transaction) only when every dimension was built (a full
  build then derives the behaviour cards from data, as before: about two minutes more, the CLI's figure); otherwise it
  says "atlas not replaced: <why>" and the atlas is unchanged. One build at a time (a lock next to the atlas); every
  attempt that took the lock is recorded next to it (one refused by the lock is not: the record is the running build's)
  (`atlas.db.last.json`), and `smartass atlas status` shows it as `last_build`. A dimension EPM lists that this
  release has no rule for is skipped, its rows kept, and said ("dimension X is new in EPM; this release cannot read
  it - update smartass"). The rows are the ones the old Smart View zoom gave: the label is the Default alias, else the
  name; the member keys are unchanged. The `member` table keeps the seven columns every release knows; the real
  names behind each key, the alias, data storage, account type and time balance go to a side table `member_real`,
  stamped with its build. An older installed release can still rebuild this file (its purge and insert work); its
  side rows are then stale, and this release ignores them until its next build. With a trusted side table the REST
  engine accepts only a real name of the key it asked for. EPM's account type and time balance are stored, not used
  (the behaviour cards still derive theirs from data). A full build then derives the cards (`atlas cards`,
  read-only grids). The doctor's `atlas` check warns on an atlas without a trusted side table
  (`knowledge/atlas.trust`): the strict check is off there until the next build.
- **Your Excel** (`xl/userjobs.py` `UserJobs`): every job in it (an answer, a follow-up, a drill's sheets, the
  `xl_*` tools) runs on an STA thread of its own, so jobs in different Excels run side by side; work in one Excel
  takes turns (a lock file per Excel process). A job has a deadline (`SV_USER_CALL_TIMEOUT`, 600 s); a job given
  up on is never killed - its next Excel call stops it - and no job closes your Excel. The objects a job made are
  released on its own thread before its turns are freed; a release not done 10 s after the answer gives the job up.
  A job's result is plain data; every failure is a readable text. Excel is reached through its window, never
  created by COM.
- **Smart View's Login in your Excel** (`xl/dialogs.py`): smartass uses no Smart View, but Smart View stays
  installed and can open its Login (at every Excel start when its start-up update check is on: doctor
  `smartview_startup`). Excel's COM blocks while that Login is open, so smartass refuses at once to work while one is
  open in any visible Excel (also for a job whose workbook is in another Excel), and names that Excel by its window
  title. After it launched an Excel and found the workbook there (a read of that Excel's workbook list and the first
  contact of its turn, in 0.6.0's order), it watches that Excel for a Login before the job's first step there: 12 s
  from when the workbook showed; where Smart View's start-up check is on (`HKCU\Software\Oracle\SmartView\extensions`
  `CheckOnStartup`, read from the registry, no COM; a policy setting reads as not set) until 25 s after the launch
  when that is later, as that Login comes 19-22 s after Excel starts. A Login later than that (a slow start) is not
  caught: that job waits until its deadline. Nothing clicks, closes or waits for a sign-in.
- **What can start an Excel:** only `xl/target._launch_excel` (`EXCEL.EXE /x <file>`: your Excel, visible, on a
  workbook), called by `_open_path` (a target file, or a first workbook's file when no Excel runs) and
  `_spawn_excel` (the TN template: a parallel agent's new workbook, or a first workbook with no Excel, no seed and
  no stored base). What smartass ends: an Excel it launched itself for a new workbook that could not be saved, when
  nothing else is open in it (`_discard_new` quits it), and an older release's processes: at server start an Excel
  that release started hidden, when its `pids.json` proves it (pid, creation time, `-Embedding`; then the file
  goes), and in `smartass setup` that release's sign-in watcher and its `smartass serve` of this install (`smartass
  doctor --fix` stops that serve too; ours when its executable is in this install's `.venv`, or its parent's is; it
  is stopped through the process object the listing found, so a pid a new process has taken since is never hit).
  Outside the engine package one more site ends processes. The relay (`run_server.py` `end_tree`) ends a server
  that has not exited 60 s after it was told to stop, with every process under it except each Excel and what runs
  under that Excel: an Excel smartass launched is a child of the server process, and it is your Excel. Anything that
  ends the server's tree from outside ends such an Excel with it: an open item (whether Claude Code does when it ends
  its MCP server is not known). `tests/test_neutral_imports.py` pins the start sites; `tests/conftest.py` and
  `tests/test_suite_guards.py` pin the engine's kills, `tests/test_relay_stop.py` the relay's.

## Knowledge layer (`knowledge/`, sub-spec 2 §1–§4, §8)

- **Source, in git:** `knowledge/semantic/metrics/<id>.toml` (name, aliases, description, unit, sign, balance,
  aggregation, materiality, members per cube, `differs`, status) and `facts/<id>.toml` (text, scope members per cube
  and/or metrics, `valid_from`/`valid_to`, severity info / warn / block, source, status). `semantic.py` validates
  them (pydantic, strict). Since 10.11 (`docs/specs/2026-11-10-knowledge-layer-design.md`, phase A) `knowledge/semantic` is the ONLY source: also `rules/` (scenarios, materiality, P&L chain, signs, cube choice, group companies; one rule per data kind), `entities/` (business areas: members per cube, FinPlan division total `finplan_total`, aliases incl. generic ones, successors), `kpis/` (ratios of two metrics), `vocab/` (vague words with options, grains, currency), `notes/` (meaning and booking pattern per member, 2 366) and account bundles as plain metrics. Facts carry `inherit` (self / down / up to the nearest metric or entity member / both); the build expands them over the atlas (`expand.py`) and a caveat says `via` the member it came from. catalog.toml, capabilities.toml, the three business maps and seed.py are gone. Behaviour cards (phase B, 10.11): Smart View gives no member properties on this provider, so `smartass atlas cards` (also run after a full `atlas build`, about 2 minutes, read-only grids) derives per (member, parent) the operator (+ / ~ / -, from parent vs sum of children over two slices; #NoAccess = unknown), the time balance (flow / balance_last / average) and the account type into atlas.db `card`; `knowledge/semantic/cards/*.toml` holds manual overrides that win per field (provenance derived@ / manual@). `weight(member, cube, parent)` reads the merged card, `~` children are neither counted, shown nor ranked; fin_lookup member hits carry the compact card and touched level/average accounts get one time-balance caveat. New cards rebuild knowledge.db.
- **Build:** `build.py` compiles the files with the atlas into `%LOCALAPPDATA%\smartass\knowledge.db` (metrics,
  members, aliases + FTS5, facts, fact scopes, member labels, meta), written to `.tmp` and swapped in. Errors (a
  metric member not in the atlas, an alias on two metrics) keep the old db; warnings (a historical fact member,
  drafts) do not. `smartass knowledge build`; a missing db is built on first use and said in the warnings.
- **Resolver:** `KnowledgeResolver` serves the model (metrics by id, name or alias, any case; tree, labels and
  existence from the atlas via `atlas_tree.AtlasTree`; a draft metric is refused by the status guard),
  `search()` for `fin_lookup`, and `caveats(touched, metrics, months)`.
- **Caveats:** `prepare` collects the entities and accounts of every grid request and lines cell (FinLite counts as
  FinMax, whose members it shares), the query metrics and the fetched months; every approved fact in scope and in
  validity becomes `{id, text, severity}` in `reply["caveats"]` and a "Caveats" section on the Answer sheet. Drafts
  in scope add one warn line "unapproved knowledge change touches this answer".
- **Monthly refresh** (`refresh.py`): `smartass knowledge refresh [--no-atlas] [--month YYYY-MM]` rebuilds the atlas
  (live), diffs it against `knowledge/atlas_snapshot.json` in the home folder (added, removed, moved, relabelled
  account and entity members), drafts a fact per moved or removed member a metric or fact references, rebuilds
  knowledge.db (meta `refresh_month`), runs the offline test set and writes `knowledge/reports/<yyyy-mm>.md`.
  `smartass knowledge approve <id|all> [--by]` sets `status = approved` and `approved = {by, on}`, then rebuilds;
  `smartass knowledge status` lists the build, drafts and last report. `fin_status` says "knowledge refresh due" when
  an answer has seen a closed month newer than the last refresh (`knowledge/state.json`).
- **Rebuild on change:** `ensure_built` compares the semantic hash in knowledge.db meta (file paths and bytes, plus
  the schema version) with the files and rebuilds first when they differ ("knowledge.db rebuilt: semantic files
  changed" in the warnings); a failing rebuild is a GuardError naming the error and keeps the old db. A newer atlas (`atlas_built_ts`) also rebuilds; if only the atlas changed and the rebuild fails, the old db is kept and the answer warns. An atlas older than 30 days is a warning on every answer and lookup.
- **Product concepts (Plan 2C):** `knowledge/semantic/concepts/<id>.toml`, seeded from `knowledge/concepts.json` by
  `seed_concepts()` (24 approved concepts: `splitt`, `streamix`, 22 booked-only `fam_*` families; texts translated
  to English). A concept has `[[block]]`s of `[[block.item]]`s (booked: account x products or entities; attributed;
  estimated: `base` x `key_share` with `num`/`den` or `same_as`; sum; missing with a note), `[[check]]` (an
  expression over item keys and `$src`, the check's source amount), `[total]` and `notes`. The build stores
  concepts and aliases (`concept`, `concept_alias`, `concept_fts`); an alias that is a metric's or another
  concept's is an error, a concept member missing from the atlas a warning. `fin_lookup(kind="concept" | "any")`
  finds them. The atlas has no consolidation operators, so metric identities (C70 = C50S + C60) are checked in the
  test set, not the tree.

## Products and the concept block (Plan 2C)

- **Products:** `Query.products` (empty = the cube's product total, unchanged). A request with products adds
  `All Products` (the dimension's name in FinMax, FinLite and FinPlan) as a third row dimension; facts carry
  `product`; the Data table has a `Product` column only when a fact has a product (answers without products keep
  their table, requests, replay keys and formulas exactly); `Pick` filters on Product and a constant Product is left
  out of the formula. Unknown products stop the answer. `table` and `ranking` take `rows: "products"`
  (`blocks/common.product_lines`: products, Other, the total over the listed products).
- **Concept block** (`model/blocks/concept.py`): `{type "concept", concept, query}` on a FinMax/FinLite query. The
  concept's amounts are fetched in their own grids, tagged `<query>:<id>:p` (account x query entity x products)
  and `<query>:<id>:e` (account x entities, no product), so they never mix with the query's metric grid. One
  section with a heading per concept block and a line per item; booked = Pick, attributed/estimated = `Base` x
  `Key share` from labelled helper columns (a 100 % key is a typed input, never a constant in a formula), missing =
  no value, then sums, the total and the checks. A check off zero is a finding (`concept_findings`), never the
  reconciliation stop. The notes, missing items and estimates become info caveats (`concept_caveats`).

## Critic, playbooks and answer memory (sub-spec 2 §5–§7)

- **Critic brief** (`model/critic.py`, pure): plan purpose, depth and hypotheses; each material line with
  `backed` and its `evidence` (a drivers block on the same metric whose tree holds the line's entity, a lines block
  on the line's entity, or a commentary driver naming the line with basis `data`); comparisons; warn/block caveats;
  checks; findings; the commentary drivers and their basis; `gaps` (unbacked material lines, warn/block caveats
  whose id or first 30 characters the commentary does not mention, a FinMax/FinLite query of flow metrics without
  `ly`). `fin_refresh` returns one brief per question.
- **Rubric** `knowledge/agent/critic.md`: six criteria 1–5 (right question, numbers reconciled, comparison basis,
  drivers backed by data, caveats respected, clear for a controller); the rounded mean must be at least 4, else fix
  by patch, at most 2 rounds. The plan guide is `knowledge/agent/plan.md`.
- **Playbooks** `knowledge/playbooks/{monthly-pack,variance-commentary,gl-extract,prodreg-input}.md` (method and
  judgement; variance-commentary quotes the vendored Apache-2.0 text unchanged). MCP prompts `plan`, `critic` and
  the four playbooks (the Markdown with a one-line preface); resources `smartass://playbooks/<name>` and
  `smartass://conventions` (the skill's finance rules section).
- **Answer memory** `answers.py`: `%LOCALAPPDATA%\smartass\answers.db` (SQLite, per user, never in git): one row per
  run (time, workbook, q, question, purpose, spec, key values, metrics, entities, periods, critic brief, score,
  notes), FTS5 over question, purpose, metrics and entities; `get` and `search` give the latest run per workbook
  and Q; rows older than 24 months are purged. A failed store is a warning, never a failed answer.

## Presentation (sub-spec 5, `docs/specs/2026-11-05-v2-5-presentation-design.md`)

- **Answer zones** (`render/layout.py`, pure): title (`plan.title`, else the purpose), subtitle, Key points,
  Charts, the tables (main table first), Checks, Notes (conventions, thresholds, findings, caveats), Chart data,
  Source. The plan is not drawn. Column B holds codes, C labels, values from D.
- **Tables:** scenario headers AC, PY, BU, FC and Δ (`blocks/common.scenario_heads`); `periods` gives one column
  group per period expression (each inside the query's periods, `validate._guard_groups`); `chain`
  (`model/chain.py`) completes the P&L chain of a metric table with EBITDA on new questions and patches, recorded
  in `chain_added`; `material` adds a live Material column (●) on the thresholds in Notes.
- **Charts** (`model/blocks/chart.py`, `render/chartdata.py`, `render/charts.py`): a `chart` block draws a block
  on the sheet or its own data (a Chart data table). IBCS styling through COM: AC navy, PY grey, BU outlined, FC
  hatched, variances green or red, labels instead of axes; the waterfall is a stacked column over helper columns,
  never Excel's native waterfall. `guard_chart` refuses a new standard/deep question without a chart; every other
  path gets `implicit_chart`. A failed chart is a warning, never a failed answer. Skill reference
  `reference/charts.md` holds the message → kind table; `tests/v2/test_v5_chart_docs.py` validates its examples.
- **Key points** (`render/keypoints.py`): at most 5; a live number sentence per material line of the main table's
  last column group (at most 4), the commentary driver as the cause, then the agent's own points.
- **Data and Sources:** `Q<n> Data` and Lines sheets carry a provenance header (`render/datasheet.py`): a control
  total on Data, a tie-out to the cube on Lines;
  the Sources sheet (`xl/sources.py`) is an append-only log of pulls with a Known limits block.
- **Layout check** (`render/check.py`, after every draw): zone order, a chart drawn, number formats, cut-off text
  or `###`, the Data header, the Sources rows. The findings go in the reply as `layout`, each with its fix, and the
  agent applies the fix by patch.

## Command line, install and upgrade (`ops/`)

smartass is the MCP server the agent starts over stdio (the relay above); it opens no port, and nothing of it runs at
logon. The command line (`ops/cli.py`, the console script `smartass`) has five commands:

- **`smartass setup`** (`ops/setup_user.py`), per user, each step one line: `config.toml` from a template when there is
  none; the MCP entries for Copilot CLI and Claude Code (an old `smartview` / `smartass2` entry replaced); the
  skill linked for both agents; the Startup entries and processes an older release left, removed and stopped when
  they are this install's (they would start or keep running old code: the old sign-in watcher's entries and process,
  `smartass-engine.lnk` when its target, read through WScript.Shell, is in this install's `.venv`, and this install's
  old `serve`; each stopped through the process object its listing found, said first); the user PATH. It writes no
  Startup entry. When `smartass doctor` would list a leftover that is a file or a setting, its last line says to run
  `smartass doctor`, then `smartass doctor --fix` (or, when only Office's settings or the old svkit record are left,
  to follow doctor's manual steps): `doctor.pending_leftovers` / `fixable_leftovers`, no process listing.
- **`smartass login`** (`ops/login.py`): the device-code sign-in per pod; then, on a PC with no atlas, the atlas build
  (said first; `--no-atlas` skips it; never over an atlas that is there). Exit codes: 0 signed in, 1 a sign-in failed
  (nothing built), 2 not set up, 3 signed in but the atlas was not built.
- **`smartass atlas status | build | cards`** (`ops/atlas_build.py`) and **`smartass knowledge build | refresh |
  approve | status`** (Knowledge layer, below).
- **`smartass doctor [--json | --bundle | --fix]`** (`ops/doctor.py`): the checks `config`, `excel`,
  `smartview_startup`, `epm`, `oauth`, `atlas`, `knowledge`, `output` (what `fin_status` runs too) and `leftovers`
  (the command and the bundle only: it lists processes and reads the Startup folder).

**No atlas.** `knowledge/atlas.present()` (read only, never creates the file, never waits on a lock): a missing file or
one without members is no atlas. `need_atlas()` refuses every entry that needs it with one command
(`no_atlas_command`: `smartass login` while a stored sign-in is missing, else `smartass atlas build`); `fin_lookup`
answers from a knowledge.db that matches the semantic files over an empty atlas in memory, with one warning naming
that command. `fin_status` and `smartass atlas status` make no atlas.db on a fresh home (`connect_read`).

**An upgrading PC.** What an older release left is found by the `leftovers` check and removed by the command that owns
it, each thing said first, each with its own proof: `smartass-engine.lnk` only when its target is in this install's
`.venv`, an old `serve` only when its executable is (or its parent's), the old sign-in watcher's process only when its
command line names this install's script, its Startup entries by their names, the pane's Startup entries by their own
text, the files by their names in this install's home. `smartass setup` the Startup entries and old processes; `smartass
doctor --fix` (which also removes this install's `smartass-engine.lnk` and stops its `serve`) the files - `secret.bin`,
`update_attempt.json`, a Startup entry of the task pane, a crashed run's copy of Smart View's options, the bridge
add-ins - and the settings only an older release read: the `[release]` and `[http]` tables of config.toml, taken out on
its bytes after a backup (`config.toml.bak-<date>`; one of the same day with the same bytes is used again) and written
only when the result parses as the old settings without those two tables, else left and the lines to delete said. The
task pane's registration in Office's settings (`ops/pane_leftovers.py`: a Trusted Add-in Catalogs entry matched by the
folder name only, never read; an `OPEN<n>` value naming `TNClaude64.xll`) is read only and reported with the steps in
Excel: smartass writes nothing under Office's registry keys. The old svkit home is never changed; its record of Smart
View's options counts only when Smart View's stored options (`CAOptionsXML`, read from the registry without COM) differ
from it. An install the deleted installer script made in `%LOCALAPPDATA%\smartass\app` is another install to this one:
its Startup entry and `serve` are reported and left (troubleshooting, "leftovers", has the steps by hand). The pane's
folder left in this install's checkout (`addin/`, now ignored by git) is reported and left: smartass deletes nothing
inside a git checkout. `scripts/check_fresh_home.py` proves a fresh install: setup and doctor on a new home in a temp
folder of its own leave only the sign-in and the atlas to do.

## Verification

- Unit tests: `cd engine; uv run pytest -q` (no Excel or EPM; the stdio smoke test starts
  `child_server.py`).
- Test set: `uv run python -m smartass.testset run` replays recorded grids and compares with the approved golden
  numbers (9 questions).
- `smartass doctor` checks the install (Command line, install and upgrade, above).
