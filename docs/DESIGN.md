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
- **Read-only towards EPM.** No submit, calculate or business-rule function exists in the engine or the VBA bridge.
- **Your Excel is never closed or saved.** Only the engine's own hidden Excel is restarted; only the engine's own
  answer files are saved.

## Tools (9, `smartass/server.py`)

| Tool | Does |
|---|---|
| `fin_analyze(spec \| patch, q, target)` | New question from a spec, or a JSON merge patch on the stored spec of `Q<n>` |
| `fin_refresh(target)` | Re-runs every v2 answer in the workbook; relative periods roll to the latest closed month |
| `fin_lookup(query, kind, cube)` | Metrics by name or alias (description, members, `differs`), atlas members and product concepts (`kind="concept"`), with caveats; `kind="answer"`: past answers from answer memory |
| `fin_drill(cell, group_by, top)` | A spec with one `lines` block: FinLite ledger lines ranked and tied to the cube |
| `fin_critic(q, scores, notes)` | The stored answer's critic brief + the rubric (`knowledge/agent/critic.md`); with `scores` (six criterion scores 1–5) computes the score (mean rounded half up) and stores it |
| `fin_status(restart)` | Version, atlas age, knowledge build and refresh due, doctor problems; `restart=True` gives a fresh server process |
| `xl_snapshot`, `xl_revert`, `xl_tidy` | Snapshot / undo around Excel MCP edits; TN style + tn-formelstil check |

Every other Excel job (read, write, pivots, Power Query, charts) is Excel MCP. Error text: every tool runs under
`core/serverkit.guard`, which turns any exception into one readable message with one next step.

## Pipeline (`pipeline.py`)

1. **Prepare** (no Excel): validate the spec (`model/spec.py`, pydantic, strict), resolve metrics and members
   (`knowledge/resolver.py` `KnowledgeResolver` over `knowledge.db` + `atlas.db`), run the hard guards
   (`model/validate.py`: unknown member or metric, open month without `preliminary`, period mismatch), plan the
   grid requests (`model/requests.py`, `model/series.py`) and fetch them in parallel over Smart View HTTP
   (`data/provider.py` → `epm/svhttp.py`). `lines` blocks drill first (`epm/drillstep.py`, concurrently).
2. **Fallback**: grids HTTP could not fetch are pulled once as ad hoc grids in the user's Excel
   (`data/adhoc.py` → `xl/pull_adhoc.py`), then prepare runs again with those facts fixed.
3. **Compute** (`model/analyze.py`): facts → `model/frame.py` → block builders (`model/blocks/*`) → expression
   trees (`model/expr.py`) → values, findings, material lines and `key_values` (`model/findings.py`).
4. **Excel job** (one job in the user's Excel, `pipeline.excel_job`): the answer's sheets are built as one .xlsx
   without Excel (`render/xlsx.py`, charts `render/xlsx_charts.py` on the rules of `render/charts.py`) from a pure
   `render/layout.py` plan - `Q<n> Answer` (its look: `render/look.py`), `Q<n> Data`, `Q<n> Lines <alias>` for
   drills - under `%LOCALAPPDATA%\smartass\tmp`. Its Excel work is recorded as **batches** of operations
   (`xl/inproc/batch.py`: get / call / set on named objects, the verbs `each` for the only loop, `nudge` and
   `exist`, a `finally` list), and each batch runs in one go inside Excel (`xl/inproc/runner.py`: one
   `Application.Run` into a small Excel-DNA add-in, `Runner.cs`, built on first use and loaded once per Excel
   process) or, where the add-in cannot load or `SV_INPROC=0`, over COM from Python - the same operations, the
   same workbook. Three batches and a save (`pipeline._file_draw`): **look** (every read the build needs: the
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
   (`xl/sheetorder.py`)); then the save (our files only, AutoSave back on). Every quiet batch puts the user's
   settings back itself; Calculation is set to manual only where formulas are written - the bring-in, a close-out
   that re-points formulas elsewhere, and the point-back after a failure - since putting automatic back recalculates
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
   came with the file) and the save - no `Workbooks.Add`, no `SaveAs`, no furniture in Excel. A failure once the
   name is reserved closes the workbook without saving and deletes that file (only it), and the error says so; the
   workbook becomes the session's once the answer is in it. A new workbook is still made from the template (label,
   SaveAs, insert) without a label id, before a base is stored, without the template, in an Excel started for a
   parallel agent, and with no Excel at all (`_new_analysis_file`: a copy of the labelled seed - found now also
   when Excel keeps the label in `docMetadata/LabelInfo.xml`, as it does on this tenant - then Excel launched on
   it). The built workbook has no TN cell styles until something needs them (`styling.ensure_styles` merges them
   from the template: `xl_tidy`, `apply_report`). Proved item by item against one made today's way by
   `scripts/check_insert.py --parts newbook` (hidden Excel, 20.11.2026).
   Measured (docs/audit/2026-11-20-inproc-bench.md, live): 1.9-2.7 s in the Excel job for an answer without a
   lines sheet as the workbook grows to 61 sheets, 3.9-4.3 s with a lines sheet, a follow-up 1.6-1.9 s, at 5
   batches per answer (4 since the review fixes). What is left is mostly Excel's own work: the insert (about 0.5 s)
   and the save of our OneDrive file. The first answer of a new workbook, built as its file (phase 3c): 2.0-3.4 s in
   the Excel job in a hidden Excel on the OneDrive folder, against 8.0-11.6 s made from the template there (the
   Open itself 1.0-1.8 s; 20.11.2026); 10.6-11.3 s made from the template in the owner's Excel (0.4.0 run log).
   Measured in the owner's Excel built as its file: (controller fills in).
5. **Reply**: `q`, `workbook`, `sheets`, `key_values` (the only numbers the agent may quote), `material`,
   `findings`, `checks`, `layout` (layout check findings with fixes), `charts`, `warnings`, `caveats`, `routes`,
   `closed`, `saved_user_workbook: false`, and for depth standard or deep `critic_brief` (`model/critic.py`). The run is stored in answer memory
   (`answers.py`) at the end of the Excel job and of every redraw.

**Follow-ups and packs** (`followup.py`): a patch is an RFC 7386 merge on the stored spec; `pipeline.redraw_job`
replaces the question's sheets whole: the old ones are parked under 'old ...' names, the new answer is moved in where
they were, and the old ones are dropped. Cells a user typed on a `Q<n>` sheet go with it (their own notes belong on
their own sheets). A formula elsewhere in the workbook that points at the question's sheets or tables follows the park
to the 'old ...' names (Excel rewrites it) and would turn #REF! with the drop: the bring-in batch finds such formulas
(the verb `mentions`, inside Excel: every other sheet's used range read in chunks of at most 100 000 cells, only the
chunks SpecialCells/Intersect finds a formula in - a protected sheet chunk by chunk, a hidden formula there makes it
"could not be read" -, no Range.Find, which would leave its options as the Find dialog's and skips filtered rows;
the charts' series; the defined names), Python rewrites exactly the references to a parked name with a namesake in
the new answer (`xl/repoint.py`: never text in a string literal, another workbook's sheet, or a longer name that
contains it), and the close-out writes them right before the drop with Calculation manual, each a compare-and-set
inside Excel (the verb `swap`: a formula changed since the bring-in read it is left alone and reported) that may
fail (a warning). A failure before the drop points them back first, by the same compare-and-set (the parked sheets
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
| `core/` | `config` (paths, settings, DPAPI secret), `logsetup`, `timing` (stages, progress), `session` (session workbooks, `sessions.json`), `serverkit` (guard, Excel jobs, lifespan) |
| `model/` | spec, periods, requests, series, resolver interface, frame, expressions, blocks, validate, findings, lines (pure; no Excel, no network) |
| `data/` | providers (HTTP, ad hoc combo), recorder (offline replay for the test set) |
| `render/` | layout and look (pure), the built file (`xlsx`, `xlsx_charts`), chart rules, read-back, layout check |
| `epm/` | Smart View HTTP/XML, grids, `ops` (bridge grids, atlas dumps), SSO, the drill browser and lines |
| `knowledge/` | atlas (SQLite), semantic files (`semantic`, `seed`), `build` (knowledge.db), `KnowledgeResolver`, monthly `refresh` |
| `xl/` | the hidden Excel (`host`, `hostproc`, `dialogs`), the user's Excel (`instances`, `target`, `jobs`, `login`, `sheets`), `styling`, `formulas`, `movein`, `sources`, `sheetorder`, `inproc` (batches of Excel operations, their COM executor and the in-process add-in), `drill_xl`, snapshots (`workbench/history`), vendored TN style and formula checker |
| `ops/` | `cli`, `doctor`, `setup_user`, `update`, `httpapi`, `atlas_build` |
| `testset/` | the B-question test set: specs in git, recordings and golden numbers outside git |

## Runtime

- **Relay** (`engine/run_server.py`, the MCP entry `smartass`): runs `child_server.py` as a child and passes
  JSON-RPC through. It replaces the child when `engine/VERSION` changes (`scripts/bump.ps1`) or on
  `fin_status(restart=True)`, and when the child crashed; a `.py` edit alone never restarts it. Log: `proxy.log`.
- **Hidden Excel**: one `DispatchEx` Excel on one STA thread with the VBA bridge built in memory at start. Every
  call has a deadline; a hang kills only this server's Excel (PID + creation time + `-Embedding`, recorded in
  `pids.json`, written atomically under a lock). Smart View global options are set while it runs and restored on
  shutdown, and at the next start after a crash.
- **User's Excel**: attached through its window; per-Excel turns (a lock file per Excel PID); never killed. A Login
  window there waits for the SSO watcher or the user.
- **SSO watcher** (`sso_watcher.ps1`, Startup folder) clicks 'Telenor SSO' in any Excel's Login dialog.
- **Drill browser**: headless Edge with its own profile, one drill at a time across sessions, shared with the HTTP
  route for sign-in.
- **Bearer route** (`epm/oauth.py`, `svhttp.BearerTransport`; optional, per pod): the Smart View HTTP requests go
  over plain HTTPS with the user's OAuth access token instead of through the browser. One device-code sign-in per
  pod (`smartass login`), stored DPAPI-protected; the refresh token is single use, so every refresh runs under a
  cross-process lock, and the servers renew it daily. A pod without a scope or a sign-in, and a refused token, fall
  back to the browser transport. Live 19.11.2026 with scopes for plan1 and stratmodel: FinMax, FinLite and FinPlan
  grids, the closed month and drill links are the browser route's cell for cell, and a cold start needs no browser
  (first grid 0.7 s against about 20 s). Drill lines (SUI query + getData) still load in the browser; the SUI GET
  answers over the token too, getData over the token is untested. The same token also opens the documented Planning
  REST API (`exportdataslice`, member properties), which is not used yet.
- **Atlas**: `atlas.db`, rebuilt with `smartass atlas build` (or copied from the monthly release by `smartass
  setup`).

## Knowledge layer (`knowledge/`, sub-spec 2 §1–§4, §8)

- **Source, in git:** `knowledge/semantic/metrics/<id>.toml` (name, aliases, description, unit, sign, balance,
  aggregation, materiality, members per cube, `differs`, status) and `facts/<id>.toml` (text, scope members per cube
  and/or metrics, `valid_from`/`valid_to`, severity info / warn / block, source, status). `semantic.py` validates
  them (pydantic, strict). Since 10.11 (`docs/specs/2026-11-10-knowledge-layer-design.md`, phase A) `knowledge/semantic` is the ONLY source: also `rules/` (scenarios, materiality, P&L chain, signs, cube choice, group companies; one rule per data kind), `entities/` (business areas: members per cube, FinPlan division total `finplan_total`, aliases incl. generic ones, successors), `kpis/` (ratios of two metrics), `vocab/` (vague words with options, grains, currency), `notes/` (meaning and booking pattern per member, 2 366) and account bundles as plain metrics. Facts carry `inherit` (self / down / up to the nearest metric or entity member / both); the build expands them over the atlas (`expand.py`) and a caveat says `via` the member it came from. catalog.toml, capabilities.toml, the three business maps and seed.py are gone. Behaviour cards (phase B, 10.11): Smart View gives no member properties on this provider, so `smartass atlas cards` (also run after a full `atlas build`, ~1 min, read-only grids) derives per (member, parent) the operator (+ / ~ / -, from parent vs sum of children over two slices; #NoAccess = unknown), the time balance (flow / balance_last / average) and the account type into atlas.db `card`; `knowledge/semantic/cards/*.toml` holds manual overrides that win per field (provenance derived@ / manual@). `weight(member, cube, parent)` reads the merged card, `~` children are neither counted, shown nor ranked; fin_lookup member hits carry the compact card and touched level/average accounts get one time-balance caveat. New cards rebuild knowledge.db.
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

## HTTP API (`ops/httpapi.py`, `smartass serve`)

Localhost only, per-user secret (`secret.bin`, DPAPI), one request at a time: `GET /health` (the doctor checks)
and `POST /analyze` (`{"spec", "target"}`, the same `_new_question` as `fin_analyze`). 400 for bad JSON, unknown
fields or an invalid spec, 401 secret, 409 busy, 413 over 1 MB, 503 Excel/EPM error, 500 (logged) otherwise.

## Verification

- Unit tests: `cd engine; uv run pytest -q` (no Excel, EPM or browser; the stdio smoke test starts
  `child_server.py` with the warm-up off).
- Test set: `uv run python -m smartass.testset run` replays recorded grids and compares with the approved golden
  numbers (9 questions).
- `smartass doctor` checks the install (config, Excel, Smart View, VBA access, EPM, oauth sign-ins, drill browser, atlas,
  knowledge (build checks against the current atlas; WARN when knowledge.db is older than the atlas or the semantic
  files, or drafts await approval), output, HTTP API, update).
