# smartass: design as built (2026-11)

The target design and its reasoning are in
[`specs/2026-11-02-pipeline-architecture-design.md`](specs/2026-11-02-pipeline-architecture-design.md); the
acceptance set is [`specs/2026-11-02-benchmarks.md`](specs/2026-11-02-benchmarks.md). This file says what was built.
The design brief and engineering log up to October 2026 (spikes, Smart View settings, return codes, the reload relay,
the performance pass, parallel agents, the first drill) are in
[`archive/DESIGN-2026-10.md`](archive/DESIGN-2026-10.md).

## Shape

- **One pipeline.** Every finance question goes `sv_gate` (scope) → `sv_answer` (pull → answer → check). A follow-up
  patches the same answer; a recipe drafts a spec server-side. There is no second way to pull or build an answer.
- **One contract.** Answer spec v1 (`core/contract.py`, pydantic, strict): `question`, `gate`, `assumptions`,
  `pulls`, `layout`, `commentary`, or `recipe` instead of `pulls` + `layout`. The agent writes it, recipes produce
  it, follow-ups patch it (`contract.patched`), and it is stored with the answer in the hidden `_svmeta` sheet.
- **Read-only towards EPM.** No submit, calculate or business-rule function exists anywhere in the engine or the
  VBA bridge.
- **Your Excel is never closed.** The engine restarts only its own hidden Excel; the user's workbooks are never
  saved (only the engine's own analysis files are).

## Tools (25)

| Group | Tools | Module |
|---|---|---|
| Pipeline | `sv_gate`, `sv_remember`, `sv_answer` | `tools_core.py` |
| Health, atlas | `sv_status`, `sv_restart`, `sv_atlas_build`, `sv_atlas_status`, `sv_atlas_changes` | `tools_status.py` |
| Members | `sv_dimensions`, `sv_find_member`, `sv_members`, `sv_member_info` | `tools_query.py` |
| Drill | `sv_drill` | `tools_drill.py` |
| Workbench | `xl_workbooks`, `xl_overview`, `xl_read`, `xl_write`, `xl_rows`, `xl_add_sheet`, `xl_format`, `xl_format_sheet`, `xl_trace`, `xl_check`, `xl_changes`, `xl_revert` | `tools_xl.py` |

Ten tools were retired in the rework (phase 6): the free-form pull/build/answer group, the template and concept
tools, the P&L tool and its warehouse tool, the scope tool, the chat query tool and the workbook refresh tool. Each
overlapped `sv_answer`, `sv_gate` or a deleted subsystem; the list is in `CHANGELOG.md`.

## The pipeline (`core/`)

1. **Scope: the advisor gate** (`core/scope.py`, `core/period.py`, `core/semantic.py`, `core/prefs.py`). A fixed
   checklist (period, scenario, entity, measure, currency/sign, granularity) is resolved against the semantic
   catalogue (`knowledge/catalog.toml`), the capability map and the user's learned defaults
   (`preferences.toml`, written by `sv_remember`). It suggests `clear`, `ambiguous` (1–3 questions, nothing is
   pulled) or `infeasible` (reason + proxy). The closed month comes from the HTTP route; the gate never starts Excel.
2. **Pull** (`core/source.py`, `svhttp.py`, `svxml.py`). Each pull is fetched over Smart View's HTTP provider
   protocol from the drill's hidden Edge (no Excel, ~0.2 s per grid once signed in) and written as a value grid in
   one block on its source sheet `Q<n> <cube>`. If HTTP fails, the grid is pulled as a Smart View ad hoc grid in the
   user's Excel (`excel/adhoc.py`); the result's `routes` say which.
3. **Answer** (`answer.py`, `answer_write.py`, `specs.py`, `formulas.py`, `styling.py`). `Q<n> Answer` appears first
   (`core/progress.py`: question, assumptions, status line). The layout is rendered as live formulas over the source
   sheets, in TN report style, with checks, the findings table, notes, source and chart. The formula checker
   (skill `tn-formelstil`) runs on every built sheet.
4. **Check and review** (`core/check.py`, `core/findings.py`, `core/commentary.py`). Reconciliation checks, formula
   error cells, sheet = fetch and unknown members become Critical / Warning / Info findings. Material lines (the
   catalogue's materiality rule) are returned in `material`; the agent explains them in `commentary.drivers`, each
   backed by data (with evidence) or labelled a hypothesis.
5. **Reply and log** (`core/reply.py`, `core/runlog.py`). Save + activate (our files only), the result dict
   (table, checks, review, material, outline, routes, timings), one JSON line per run in `runs.jsonl` and a row in
   the workbook's hidden `_runlog` sheet.

**Follow-ups** (`core/followup.py`, `core/owned.py`). A follow-up patches the stored spec. Pulls whose definition
did not change are reused (not fetched); the answer is redrawn inside the tool-owned range, so the user's cells and
charts outside it survive. The answer sheet is backed up (`~Q<n> Answer`, hidden) and restored if the redraw fails.

## Recipes (`core/recipes/`)

Pure functions from a few parameters to a draft spec, expanded inside the Excel job (closed month, catalogue,
member labels): `pnl` (B01, B13, B15), `actual_vs_plan` (B05, B14), `monthly_trend` (B07), `cost_base` (B06),
`concept` (B09, `concepts.py` + `knowledge/concepts.json`). `overrides` patch the draft like a follow-up.

**Live P&L facts.** Recipe `pnl` reads its facts live on every call: `pnl_facts.py` resolves the area or unit
(abolished areas are refused with their successors), `extract.py` reads the grids the catalogue defines (HTTP first,
`ops.query` over COM as the fallback) into an in-memory `factstore.FactStore`, and the P&L calc module's `compute`
with the
`analytics/` package (frames, measures, variance, materiality, KPIs, bridges, landing, reconciliation) produces the
facts. `fin_spec.py` turns them into the answer spec; its check "Live read = source sheet" compares the facts with
the source sheets the formulas read. Nothing is kept between calls (the DuckDB warehouse was removed).

## Excel

- **The engine's hidden Excel** (`host.py`, `hostproc.py`, `ops.py`, `dialogs.py`, `vba/`). One `DispatchEx` Excel on
  one STA thread, with the VBA bridge built in memory at every start. Every call has a deadline; a hang kills only
  this server's own Excel (PID + creation time + `-Embedding`, recorded in `pids.json`) and the next call restarts
  it. It is started only when needed (COM fallback, bridge build).
- **The user's Excel** (package `excel/`): `instances.py` (attach through the window, the versioned bridge add-in),
  `target.py` (session / active / new / path targets, archive guard), `jobs.py` (per-Excel turns, retries, fast
  mode), `login.py` (Login dialog wait), `sheets.py` (question numbering, rollback, `_svmeta`, `Sources`),
  `adhoc.py` (the Smart View fallback pull). A call stuck in the user's Excel is never resolved by killing it.
- **The workbench** (package `workbench/`): `books.py` (find and read), `changes.py` (write, rows, sheets),
  `fmt.py` (TN formats, trace, formula check), `history.py` (snapshot before every change, change log, revert).
- **Sheet names** are English only: `Q<n> Answer`, `Q<n> <cube>`, `Q<n> Lines` (table `Lines_Q<n>`), `Sources`
  (`Sources SV` when the user has an unrelated `Sources`). `sheetorder.py` fixes the order after every answer.

## Login hardening (design §3.2)

- **Hidden Excel.** A Smart View Login window that Windows reports as hung (`IsHungAppWindow`) counts as stuck at
  once (seen within about 5–7 s); otherwise the SSO watcher gets `LOGIN_GRACE_S` (30 s). A stuck sign-in restarts
  the hidden Excel once. `rc=752/57704 Empty response from provider` is retried in place once with a progress note;
  if it survives, `ProviderEmpty` restarts only the hidden Excel and retries once (`ExcelHost.call`).
- **The user's Excel (never restarted).** A Login open for 5 s sends a note ("waiting for the automatic SSO
  click"); a Login hung for 10 s fails fast with "not responding; nothing was closed". The ad hoc pull says so on
  the first "Empty response" connect and retries a transient retrieve once.
- **Progress notes** (`timing.say`, MCP progress) tell the user what is happening during any long stage.

## Drill (`tools_drill.py`, `drill_*.py`)

`sv_drill` reads the FinLite drill page in a headless Edge (own profile, one drill at a time across sessions, closes
after 5 idle minutes), replays the page's own data request with paging, splits big or too-wide cells, reconciles
the lines to the cube, repairs text, caches lines in `drill.db` and writes `Q<n> Lines`. The report is an
`sv_answer` spec with `drill` cells. Never Smart View's own drill-through (it opens the user's browser).

## Atlas and members

`atlas.py` holds every dimension tree of the three cubes in SQLite (`atlas.db`, built by `sv_atlas_build` in about
3 minutes); `atlas_history.py` diffs builds (`sv_atlas_changes`). Member lookups and the gate read it locally.

## HTTP API (`httpapi.py`, `smartass serve`)

Localhost only, per-user secret header, one request at a time: `GET /health`, `POST /gate` and `POST /answer`, which
call the same `tools_core.gate()` / `tools_core.answer()` as the MCP tools. 400 for bad input, 409 when busy, 413
for a body over 1 MB, 503 for an Excel/EPM error. `/answer` in the `serve` process uses its own hidden Excel.

## Process model

- `run_server.py` is a stable relay per client session; the server behind it restarts by itself after a code change.
- The SSO watcher (`sso_watcher.ps1`, Startup folder) clicks 'Telenor SSO' in any Excel's Login dialog.
- Logs: `%LOCALAPPDATA%\smartass\logs` (`server.log` with `stage >` lines, `queries.jsonl`, `runs.jsonl`).
  `smartass doctor` checks the install cheaply; `--bundle` zips redacted diagnostics.

## Verification

- Unit tests: `cd engine; uv run pytest -q` (no Excel or EPM needed).
- Live bench (route `core`, gate cases): `engine/tests/live/bench.py`; results in `docs/audit/`.
- Live reliability matrix (sign-in, dialogs, killed Excel, drill, workbench): `engine/tests/live/matrix.py`.
