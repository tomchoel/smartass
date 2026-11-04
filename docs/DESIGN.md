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

## Tools (8, `smartass/server.py`)

| Tool | Does |
|---|---|
| `fin_analyze(spec \| patch, q, target)` | New question from a spec, or a JSON merge patch on the stored spec of `Q<n>` |
| `fin_refresh(target)` | Re-runs every v2 answer in the workbook; relative periods roll to the latest closed month |
| `fin_lookup(query, kind, cube)` | Metric names (interim catalogue) and atlas members |
| `fin_drill(cell, group_by, top)` | A spec with one `lines` block: FinLite ledger lines ranked and tied to the cube |
| `fin_status(restart)` | Version, atlas age, doctor problems; `restart=True` gives a fresh server process |
| `xl_snapshot`, `xl_revert`, `xl_tidy` | Snapshot / undo around Excel MCP edits; TN style + tn-formelstil check |

Every other Excel job (read, write, pivots, Power Query, charts) is Excel MCP. Error text: every tool runs under
`core/serverkit.guard`, which turns any exception into one readable message with one next step.

## Pipeline (`pipeline.py`)

1. **Prepare** (no Excel): validate the spec (`model/spec.py`, pydantic, strict), resolve metrics and members
   (`knowledge/atlas_resolver.py` over `atlas.db` + `knowledge/interim.py`), run the hard guards
   (`model/validate.py`: unknown member or metric, open month without `preliminary`, period mismatch), plan the
   grid requests (`model/requests.py`, `model/series.py`) and fetch them in parallel over Smart View HTTP
   (`data/provider.py` → `epm/svhttp.py`). `lines` blocks drill first (`epm/drillstep.py`, concurrently).
2. **Fallback**: grids HTTP could not fetch are pulled once as ad hoc grids in the user's Excel
   (`data/adhoc.py` → `xl/pull_adhoc.py`), then prepare runs again with those facts fixed.
3. **Compute** (`model/analyze.py`): facts → `model/frame.py` → block builders (`model/blocks/*`) → expression
   trees (`model/expr.py`) → values, findings, material lines and `key_values` (`model/findings.py`).
4. **Excel job** (one job in the user's Excel): `Q<n> Data` (table `Data_Q<n>`), `Q<n> Answer` from a pure
   `render/layout.py` plan written by `render/excel.py`, charts (`render/charts.py`), `Q<n> Lines <alias>` for
   drills, TN style (`xl/styling.py`), read-back (`render/readback.py`), the spec stored in hidden `_svmeta`, sheet
   order (`xl/sheetorder.py`), save (our files only).
5. **Reply**: `q`, `workbook`, `sheets`, `key_values` (the only numbers the agent may quote), `material`,
   `findings`, `checks`, `warnings`, `routes`, `closed`, `saved_user_workbook: false`.

**Follow-ups and packs** (`followup.py`): a patch is an RFC 7386 merge on the stored spec; the answer is redrawn in
place inside its owned range (`xl/owned.py`), so the user's cells outside it survive; a failed redraw restores the
backup. `fin_refresh` redraws every stored v2 question; one failure never stops the rest.

## Packages

| Package | Contents |
|---|---|
| `core/` | `config` (paths, settings, DPAPI secret), `logsetup`, `timing` (stages, progress), `session` (session workbooks, `sessions.json`), `serverkit` (guard, Excel jobs, lifespan) |
| `model/` | spec, periods, requests, series, resolver interface, frame, expressions, blocks, validate, findings, lines (pure; no Excel, no network) |
| `data/` | providers (HTTP, ad hoc combo), recorder (offline replay for the test set) |
| `render/` | layout (pure), COM writer, charts, read-back |
| `epm/` | Smart View HTTP/XML, grids, `ops` (bridge grids, atlas dumps), SSO, the drill browser and lines |
| `knowledge/` | atlas (SQLite), atlas resolver, interim metrics from `knowledge/catalog.toml` |
| `xl/` | the hidden Excel (`host`, `hostproc`, `dialogs`), the user's Excel (`instances`, `target`, `jobs`, `login`, `sheets`), `styling`, `formulas`, `owned`, `sheetorder`, `drill_xl`, snapshots (`workbench/history`), vendored TN style and formula checker |
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
- **Atlas**: `atlas.db`, rebuilt with `smartass atlas build` (or copied from the monthly release by `smartass
  setup`).

## HTTP API (`ops/httpapi.py`, `smartass serve`)

Localhost only, per-user secret (`secret.bin`, DPAPI), one request at a time: `GET /health` (the doctor checks)
and `POST /analyze` (`{"spec", "target"}`, the same `_new_question` as `fin_analyze`). 400 for bad JSON, unknown
fields or an invalid spec, 401 secret, 409 busy, 413 over 1 MB, 503 Excel/EPM error, 500 (logged) otherwise.

## Verification

- Unit tests: `cd engine; uv run pytest -q` (no Excel, EPM or browser; the stdio smoke test starts
  `child_server.py` with the warm-up off).
- Test set: `uv run python -m smartass.testset run` replays recorded grids and compares with the approved golden
  numbers (8 questions).
- `smartass doctor` checks the install (config, Excel, Smart View, VBA access, EPM, drill browser, atlas, catalogue,
  output, HTTP API, update).
