# Changelog

## Unreleased

### The answer's Excel work runs inside Excel (phase 3b)

- An answer's Excel work is recorded as batches of operations and each batch runs in one go inside Excel, through a
  small Excel-DNA add-in the engine builds on first use (C# compiled by the .NET Framework compiler that ships with
  Windows) and loads silently with `RegisterXLL` once per Excel process; one `Application.Run` per batch instead
  of dozens of cross-process calls. Where it cannot load (no compiler, a 32-bit Excel, a refused or blocked
  `RegisterXLL`) or `SV_INPROC=0` is set, the same batches run over COM, with one warning in the server log.
- The file job is four batches and a save: look (every read the build needs: the sheets, the base, which workbook is
  active, AutoSave, the Sources log's kind, the stored specs' keys, the defined names), move in (a follow-up's park,
  the insert, the moves, the Answer active), finish (tables, formulas, calculate, the read-back), close out (status,
  Key point heights, chart labels, a follow-up's drop, the Sources rows, the spec, the sheet order) and the save of
  our own file. Each batch keeps Excel quiet and puts the user's settings back itself, also when a step fails;
  between the batches Excel paints the formatted answer, then its numbers and charts, when the workbook is on
  screen. A failure is undone as before (rollback, the old answer's names put back, the new answer kept once the old
  one is gone), also for a job abandoned by a timeout; an error now names the Excel operation that failed and Excel's
  own message, and says when AutoSave was left off in the user's workbook.
- Measured live (`scripts/bench_answers.py --rounds 3` and `scripts/check_followups.py` in the owner's Excel,
  20.11.2026; docs/audit/2026-11-20-inproc-bench.md): the Excel job of an answer without a lines sheet 1.9 / 2.4 /
  2.7 s per round (2.0 / 2.7 / 3.3 s before), with one 4.1 / 3.9 / 4.3 s (4.2 / 4.1 / 5.2 s): the 3 s / 5 s gate
  holds in every round now. A follow-up 1.64-1.91 s (2.08-2.41 s), `fin_refresh` of two answers 3.4 s (4.3 s);
  5 batches and 9 COM calls per answer (33-34 calls before). Over COM (the fallback) a follow-up is 2.5-2.9 s.
- The label of a thin waterfall step is now lifted 10 pt as the chart rules mean it: the lift was set before the
  calculation, which put the label back where Excel places it (measured in a hidden Excel); it is now read and set
  once the answer is calculated, while the other point labels are positioned before it, as before.
- `scripts/check_insert.py` now also runs whole jobs through both executors in a hidden Excel and compares the
  workbooks item by item (76 of 76 checks the same), and makes a follow-up and a new answer fail in their finish
  batch through both: the workbook, Excel's settings and the active sheet are as before. `scripts/check_inproc.py`
  proves the verb `each` and the job's reads equal in both executors (34 of 34).

### The answer's sheets inserted with one call (phase 3b)

- The built file's sheets go into the workbook with one `Sheets.Add(Type:=<the file's path>)` (`xl/movein.py`)
  instead of opening the file, copying its sheets, closing it and breaking the link the copy left: no second
  workbook is opened, no link is left, and the move-in is 2 COM calls instead of 5. `scripts/check_insert.py` (a
  hidden Excel of its own, nothing saved) compared both routes on nine answers (every chart kind, thin waterfall
  steps with their label fixes, a matrix, kpi tiles, Key points, a Lines sheet) and on whole jobs (a second answer,
  a follow-up): cells, formats, borders, sizes, page setup, page breaks, window, names, conditional formats,
  charts, links, where the sheets land, refusals and the Excel around them - 71 of 71 checks the same.
- Measured live (`scripts/check_followups.py` in the owner's Excel, 20.11.2026, every step "checks OK"): the second
  answer 1.86 s and 33 COM calls (2.1 s and 36 before), follow-ups 2.08-2.41 s and 34 calls (2.2-2.9 s and 37), the
  move-in stage 0.66-1.03 s (0.81-1.22 s). The move-in alone in a hidden Excel: 291 ms against 485 ms (median of 5).
- Removed with the copy: the "temporary workbook could not be closed" warning and its retry at the end of a job,
  and the in-process executor's `break_links` verb (no verb ships now; the mechanism stays for the next task).

### Answers built as a file (0.3.0)

- An answer's sheets (`Q<n> Answer` with its charts, `Q<n> Data`, `Q<n> Lines <alias>`) are built as an .xlsx
  without Excel (`render/xlsx.py`, `render/xlsx_charts.py`), moved into the workbook in one copy and finished there
  (`xl/movein.py`), with the user's Excel quiet and left as it was found. The live comparison of all nine test-set
  cases passed against the old COM writers: every difference left is on its allowed list (cell style names, a
  0.2 pt chart height, Data widths within AutoFit's band; `docs/audit/2026-11-19-render-comparison.md`).
- Measured live (`scripts/bench_answers.py --rounds 3`, all nine cases, one workbook growing to 28 answers): the
  Excel job of an answer without a lines sheet went from 15.6 / 25.8 / 46.0 s per round
  (`docs/audit/2026-11-19-baseline-v2.md`) to 2.0 / 2.7 / 3.3 s; over all cases from 33.7 s to 2.8 s, the whole
  answer from 34.9 s to 4.1 s, and from 280 COM calls to 34. The gate of 3 s (5 s with a lines sheet) per round
  held in rounds 1 and 2 and was missed in round 3: 3.29 s, and 5.19 s with a lines sheet. What grows is the save of
  the file (Excel's own) and reads that walk every sheet of a 61-sheet workbook inside one COM call; the controller
  accepted the miss, and the re-gate moves to the next task (`docs/audit/2026-11-19-file-render-bench.md`). In
  short: 2-3 s for an answer without a lines sheet in a workbook of normal size (3.3 s at 61 sheets), 4-5 s with a
  lines sheet, about 6-7 s for the first answer of a new workbook (the workbook is created).
- Follow-ups and `fin_refresh` replace the question's sheets whole (the old ones are parked, the new ones moved in
  where they were, the old ones dropped). Cells typed on a `Q<n>` sheet go with it; a failure puts the old answer
  back, and once the old sheets are gone it keeps the new one and says so (also when the final save fails).
- Known limit (to be fixed in the next phase): a formula on the user's own sheet that points at a `Q<n>` sheet or
  its table (`='Q3 Answer'!E12`, `=SUM(Data_Q3[Value])`) shows #REF! after a follow-up or refresh of that question,
  because the old sheet is deleted. Until then, link to values by pasting them, or re-link after a refresh.
- The answer lands on the new Answer whenever its workbook is the active one, also in an Excel that is not visible.
- The Sources log's print area is `A1:N1048575` (a finite last row), so openpyxl reads our files without a warning.
- Removed: the COM writers (`render/excel.py`'s Data/Answer writers, `render/charts.py`'s drawing,
  `drill_xl.write_lines`), `xl/owned.py` (the stored drawn range is no longer used), `config.RENDER` /
  `SV_RENDER`, and `scripts/compare_renders.py`. `SV_FASTCOM` stays.

### Bearer route: grids without the hidden browser (optional, per pod)

- Smart View HTTP requests go over plain HTTPS with the user's own OAuth access token where a pod has a scope and a
  sign-in (`svhttp.BearerTransport` in front of the browser transport). Anything it cannot authenticate falls back
  to the browser route; drill-through is unchanged (untested over the token).
- Live 19.11 with scopes for plan1 and stratmodel: FinMax, FinLite and FinPlan grids, the closed month and drill
  links over the token, the same as the browser route cell for cell; from a cold start no browser is started (first
  grid 0.7 s against about 20 s). Drill lines still load in the browser.
- `smartass login` (`--status`, `--force`, `--forget`): Oracle's device-code sign-in, one approval per pod; the
  sign-in is DPAPI-protected under `oauth\` in the home folder. The refresh token is single use: refreshes run
  under a cross-process lock, and the MCP server and `smartass serve` renew sign-ins daily (an unused one ends
  after 7 days).
- Off until `config.toml` has `[oauth]` `idcs`, `client_id` and `[oauth.scopes]`; doctor check `oauth`.

### Knowledge layer phase B (0.2.23)

- Behaviour cards derived from data (`smartass atlas cards`, after every full atlas build): operator per (member,
  parent), time balance, account type; manual overrides in `knowledge/semantic/cards/`. Consolidation uses the real
  operators (`~` children out of totals, lines, rankings and drivers). Smart View member-information requests do not
  work on this provider (probe 10.11) and are not allowed.
- New metrics, verified live: EBIT, D&A incl. write-downs, profit before tax, net financial items, net income
  (FinMax), capex, total assets, fixed subscriptions (total, internet, TV, telephony), internet/TV ARPU, FinPlan FTE,
  FinPlan EBITDA margin; D&A gets its FinPlan member. FinPlan Target carries D&A, EBIT, PBT and CAPEX since FY26.
- Facts on FinMax scenarios (only Actual has data), FinPlan net income (no tax), balance sheet, free cash flow (empty),
  write-downs FY26 and volume units; an access rule (numbers depend on the reader's EPM rights).

### Knowledge layer phase A (0.2.22, spec `docs/specs/2026-11-10-knowledge-layer-design.md`)

- One source: `knowledge/semantic` gains `rules/`, `entities/`, `kpis/`, `vocab/`, `notes/` and account-bundle
  metrics, folded from catalog.toml, capabilities.toml and the three 15.10 business maps (all deleted with seed.py).
- Caveats inherit over the atlas tree (`inherit` down / up / both) and say `via` the member they came from, also on
  the Answer sheet: C40S carries the C40510 trap, a division its reorgs.
- `fin_lookup` kinds `entity`, `kpi`, `vocab`, `rule`; member hits carry their meaning note; short and generic aliases
  follow the catalogue's matching rules (longest match wins).
- Scenarios and the default materiality come from the rules (raw-member default now 10 %, Tom's catalogue decision);
  FTE materiality 10 FTE / 5 %.
- The build stops on contradictions (entity placement, FinPlan division level 1xxx vs 2xxx, aliases across kinds,
  rules defined once, required rules approved). A stale atlas warns on every answer; an atlas refresh that breaks the
  knowledge keeps the old knowledge.db.

v2 stands alone: one MCP server `smartass` with nine tools; the agent writes an AnalysisSpec v2 and the engine
validates, computes and draws it (design `docs/DESIGN.md`, `docs/specs/2026-11-04-v2-overview.md`). v1 is retired.

### Fixed

- Critic score: `fin_critic(q, scores, notes)` takes the six criterion scores and computes the score itself (mean
  rounded half up). The agent no longer computes it: a pane run stored 4 for a mean of 4.83.

- Pane live-test fixes (v2): follow-ups (`fin_analyze(patch)`), `fin_refresh` and `fin_critic` default to the
  workbook of the latest answer in answers.db, not the old session file (explicit targets win; no stored answer
  keeps the session workbook); a target that is the name of an open workbook (e.g. an unsaved `Book1`, with or
  without extension) resolves to it in any Excel (two of that name: ambiguous; the archive guard applies);
  `fin_critic(target=...)` finds and scores the answer in that workbook; answer lines, KPI tiles, bridge steps,
  trend headers, drivers/ranking/matrix titles and key_values show metric names (a raw member: its metric if it
  is that metric's only member, else the atlas label) instead of the tokens passed; `xl_tidy` accepts row kinds
  `subtotal`, `total`, `heading` (= sum, grand, section). Pane: one `/api/chat` run at a time (409 with a message,
  send disabled while running); the pane prompt pins `target="active"` on every fin_analyze / fin_critic /
  fin_refresh / fin_drill call, stops on a fin_* error instead of rebuilding the analysis by hand with Excel MCP,
  and asks for metric names from fin_lookup. `npm test` in `addin/tn-claude` tests the server.

### Added

- v2 presentation (sub-spec 5, Plan 5): the answer sheet has fixed zones (title from `plan.title`, Key points,
  Charts, the tables, Checks, Notes, Chart data, Source); the plan is no longer drawn. Scenario headers (AC, PY, BU,
  FC, Δ) and a code column. `TableBlock.periods` adds column groups (default P&L/cost table:
  `{"type":"table","query":"q","periods":["closed","closed-ytd"]}`, month and YTD); `chain` completes the P&L chain
  of a table with EBITDA (new questions and patches); `material` adds a live ● Material column on labelled
  thresholds. New `chart` block (`kind` waterfall, line, column, bar, stacked, stacked100, combo, variance; `title`
  the message; `source` a block or own data laid out as Chart data; `series`, `highlight`, `sort`), drawn in IBCS
  style through COM. A new standard/deep question without a chart is refused (`guard_chart`); other paths get an
  implicit chart. Key points (at most 5) are live number sentences for the material lines, with the commentary
  driver as the cause. `Q<n> Data` gets a provenance header and a control total,
  Lines sheets a tie-out; a Sources sheet logs every pull with known
  limits. The reply has `layout` (the layout check's findings, each with a fix) and `charts`. Skill: answer-sheet
  rules, `reference/charts.md` (message → kind, one example per kind, tested against the schema), spec reference
  for `periods`, `chain`, `material`, `plan.title` and the chart block; plan guide and pane prompt require a chart
  chosen by the message.

- Excel task pane on v2 (Plan 3A): the TN Claude pane source now lives in `addin/tn-claude` (served from the repo
  on `https://localhost:3000`); its chat may use only `mcp__smartass`, `mcp__excel` and read-only tools, with a v2
  pane prompt (fin_analyze `target="active"`, xl_snapshot before Excel MCP edits, xl_tidy, never save). New CLI
  `smartass xl changes|revert|check` (one JSON object on stdout) backs the pane's Endringer and Formelsjekk tabs.
- v2 products and product concepts (Plan 2C): `Query.products` adds the `All Products` dimension to the grid (a
  `Product` column in the Data table only then); `table`/`ranking` take `rows: "products"`; 24 approved concepts
  (`splitt`, `streamix`, 22 `fam_*` families) as `knowledge/semantic/concepts/*.toml` seeded from `concepts.json`,
  found by `fin_lookup(kind="concept")`; the `concept` block shows booked, attributed, estimated (base x key share
  in helper columns) and missing items, the total and its checks as findings; the concept notes become caveats.
  knowledge.db is rebuilt when the semantic files changed (hash differs), not only when it is missing. Test-set
  case `b09` (mobile subscriptions family).
- v2 critic, playbooks and answer memory (Plan 2B): `critic_brief` in standard/deep replies and per question in
  `fin_refresh` (material lines backed by data, gaps); `fin_critic(q, score, notes)` returns the brief and the rubric
  (`knowledge/agent/critic.md`) and stores the score; the plan guide `knowledge/agent/plan.md`; four playbooks
  (monthly-pack, variance-commentary, gl-extract, prodreg-input) as MCP prompts with plan and critic, and as
  resources `smartass://playbooks/<name>` + `smartass://conventions`; answer memory `answers.db` (FTS5, 24 months)
  searched by `fin_lookup(kind="answer")`.
- v2 knowledge layer (Plan 2A): metrics and facts as TOML files in `knowledge/semantic`, compiled with the atlas
  into knowledge.db (build checks); `KnowledgeResolver` replaces the interim catalogue; every reply and Answer sheet
  carries the caveats it touched; `fin_lookup` finds metrics by alias; `smartass knowledge refresh | approve |
  status` for the monthly atlas diff, drafts and impact report; doctor check `knowledge` replaces `catalog`. The
  v1-only advisor texts `gate.md` / `findings.md` and the `products_map.json` side files are removed.
- v2: ad hoc fallback for failed HTTP grids, record/replay, test set (record / approve / run offline),
  text-fix candidates in the reply, wrapped notes, parallel drills.
- v2 ledger lines: lines block and fin_drill (supplier / project / account ranking tied to the cube), clean
  refusals, shorter SUMIFS.
- v2 robustness and toolkit: AutoSave paused while writing, session workbook kept across /mcp reconnects, a lost
  first request retried, xl_snapshot / xl_revert / xl_tidy.
- v2 complete answers: charts (trend, bridge), follow-ups via fin_analyze(patch, q), fin_refresh for monthly
  packs, drivers skip pass-through nodes, Q<n> Data sorted with data sheets.
- v2 model core (`engine/smartass/model`): typed analysis spec, period algebra, fact frame, expression trees
  (value and Excel formula from one tree), seven computed blocks (table, kpi_tiles, trend, matrix, bridge, ranking,
  drivers), data guards, material lines and key values. Not wired to tools yet (Plan 1B).
- v2 first answer (`smartass2` server, beside v1): `fin_analyze`, `fin_lookup`, `fin_status`; Smart View HTTP
  provider, interim atlas resolver, Q<n> Answer as live formulas on table Data_Q<n>, "Sheet = model" readback;
  answer files in `<analyser>\v2`. Not yet: follow-ups, packs, drill, xl_* tools, ad hoc fallback (Plan 1C).
- `sv_gate`: the advisor gate. A fixed checklist (period, scenario, entity, measure, currency/sign, granularity)
  against the semantic catalogue, the capability map and the user's learned defaults; `clear`, `ambiguous` (1-3
  questions) or `infeasible` (reason + proxy). It never pulls. `sv_remember` saves a gate answer as a default.
- `sv_answer`: one question from an answer spec v1 (`core/contract.py`): pulls over Smart View's HTTP provider
  (Smart View in Excel as the fallback), `Q<n> Answer` shown first, live formulas, checks, a findings table
  (Critical / Warning / Info), material lines and commentary drivers (data with evidence, or a hypothesis).
  - Follow-ups patch the stored spec (`followup=`): unchanged pulls are reused, the answer is redrawn inside the
    tool-owned range, a hidden backup restores it if the redraw fails.
  - Recipes draft the spec server-side: `pnl` (area or unit P&L, month and YTD vs last year, budget, forecast,
    bridges, KPIs, landing), `actual_vs_plan`, `monthly_trend`, `cost_base`, `concept`.
  - Run log: one JSON line per answer in `runs.jsonl` and a row in the workbook's hidden `_runlog` sheet.
- Recipe `pnl` reads its facts live on every call into an in-memory store; its check "Live read = source sheet"
  compares them with the source sheets. Materiality: 10 % of |comparison| (counted from NOK 500k) or NOK 5 million;
  a month swing of NOK 5 million is flagged; revenue, personnel cost and EBITDA are always commented. Consumer and
  Infrastructure Growth (abolished 1 June 2026) are refused with their successors.
- Login hardening: a Smart View Login that Windows reports as not responding counts as stuck at once (hidden Excel:
  restarted once; your Excel: "not responding; nothing was closed" after 10 s, never closed); a note after 5 s of
  an open Login; "Empty response from provider" is retried with a progress note, and in the hidden Excel a second
  one restarts only that Excel and retries once. Progress notes during every long stage.
- HTTP API: `POST /gate` and `POST /answer` beside `GET /health` (same calls as the MCP tools; 400 / 401 / 409 busy
  / 413 / 503).
- `[svhttp]` in `engine\defaults.toml`: the data source names and `allow_test`; environment variables
  `SV_HTTP_FINMAX_SERVER`, `SV_HTTP_FINPLAN_SERVER`, `SV_HTTP_ALLOW_TEST`, `SV_HTTP_TIMEOUT`.
- Semantic catalogue `knowledge\catalog.toml` (P&L lines and checks per cube, business areas, scenarios, KPIs,
  materiality) with a `smartass doctor` check.
- Faster answer sheets: answer cells written as one block, check colours set once per run of cells.

### Changed

- **One server, `smartass`, runs v2.** `run_server.py` -> `child_server.py` -> `smartass.server`; the side-by-side
  `smartass2` entry, `run_server_v2.py` and `child_server_v2.py` are gone (`smartass setup` removes a leftover
  `smartass2` MCP entry). Engine code lives in `engine/smartass/` (`core`, `model`, `data`, `render`, `epm`,
  `knowledge`, `xl`, `ops`); answer files go to the plain analyser folder again (no `v2` subfolder).
- **Reload on request only.** The relay restarts its server when `engine/VERSION` changes (`scripts/bump.ps1`
  bumps it with `pyproject.toml`; `smartass.__version__` reads it) or on `fin_status(restart=True)`, never on a
  `.py` edit.
- **Runtime fixes.** `pids.json` is written atomically under a lock file; the HTTP API secret is DPAPI-protected in
  `secret.bin` (a plain `http.secret` in `config.toml` is moved there once by `smartass setup`); Smart View options
  left by a crashed run are restored at the next start; `smartass doctor` checks Excel's VBA project access
  (Trust Center).
- **HTTP API:** `POST /analyze` (`{"spec", "target"}`, the same run as `fin_analyze`) replaces `/gate` and
  `/answer`; `GET /health` stays.
- **CLI:** `smartass atlas build | status` (the atlas build is no longer a tool); `doctor`, `setup`, `update` and
  `serve` run on the v2 modules (`python -m smartass.ops.cli`).
- `fin_status` returns the doctor's problems and takes `restart`.
- The SSO grace for the hidden Excel's sign-in (`LOGIN_GRACE_S`, env `SV_LOGIN_GRACE`) defaults to 30 s.
- Engine modules reorganised: packages `excel/` (the user's Excel) and `workbench/` (xl_* tools), `answer` +
  `answer_write`, `host` + `hostproc`.
- Tool count 35 -> 25 (v1 rework), then 8 (v2).

### Removed (breaking)

- **v1 retired:** the package `engine/smartview_mcp` (gate, recipes, answer spec v1, workbench tools), its tests,
  the live bench (`tests/live`, replaced by `python -m smartass.testset`) and the skill's `recipes.md`. Tool
  renames:

  | v1 | v2 |
  |---|---|
  | `sv_answer` (and `sv_gate`, `sv_remember`, recipes) | `fin_analyze` (AnalysisSpec v2; follow-ups by `patch`) |
  | `sv_drill` | `fin_drill` (or a `lines` block) |
  | `sv_find_member`, `sv_members`, `sv_member_info`, `sv_dimensions` | `fin_lookup` |
  | `sv_status`, `sv_restart`, `sv_atlas_status` | `fin_status` (`restart=True`) |
  | `sv_atlas_build`, `sv_atlas_changes` | `smartass atlas build` / `smartass atlas status` (CLI) |
  | `xl_workbooks`, `xl_overview`, `xl_read`, `xl_write`, `xl_rows`, `xl_add_sheet`, `xl_format`, `xl_format_sheet`, `xl_trace`, `xl_check`, `xl_changes`, `xl_revert` | Excel MCP + `xl_snapshot` / `xl_revert` / `xl_tidy` |
  | (new) | `fin_refresh` (monthly packs) |

- HTTP endpoints `POST /gate` and `POST /answer` (use `POST /analyze`).
- Tools `sv_query`, `sv_ws_pull`, `sv_ws_build`, `sv_ws_answer`, `sv_analyse`, `sv_concept`, `fin_pnl`,
  `sv_scope`, `sv_warehouse`, `sv_refresh_workbook`: use `sv_gate` + `sv_answer` (recipes for P&L, plan, trend,
  cost base and concepts); refresh an old workbook in Excel with Smart View's Refresh.
- The DuckDB warehouse (`warehouse.duckdb`, `paths.warehouse`, the doctor's warehouse check) and the `duckdb`
  dependency.
- Norwegian sheet names: sheets are `Q<n> Answer`, `Q<n> <cube>`, `Q<n> Lines` and `Sources`; old `Svar` /
  `Kilder` / `Linjer` sheets are no longer recognised as answers.

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
