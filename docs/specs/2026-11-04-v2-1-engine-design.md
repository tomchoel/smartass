# smartass v2 · sub-spec 1: engine

Status: draft for review (2026-11-04). Parent: [overview](2026-11-04-v2-overview.md) (decisions D1–D16).
Covers the model, data, render and runtime layers and the MCP tool surface. Knowledge content is sub-spec 2;
this spec fixes only the contracts knowledge plugs into (`fin_lookup`, metric resolution, caveats).

## 1. Tool surface (8 tools)

| Tool | Does | Replaces (v1) |
|---|---|---|
| `fin_lookup(query, kind?)` | Members, metrics, facts, past answers; every hit carries its caveats | `sv_find_member`, `sv_members`, `sv_member_info`, `sv_dimensions`, `sv_gate` |
| `fin_analyze(spec \| patch, q?, target?)` | Validate → fetch → compute → render → check; a patch with `q` changes an existing answer | `sv_answer` (spec, recipe, follow-up), `sv_remember` |
| `fin_drill(cell)` | Ledger / invoice lines behind a FinLite cell, reconciled; lines become a `Data` source for blocks | `sv_drill` |
| `fin_refresh(workbook?, through?)` | Re-runs every saved spec in a workbook to the latest closed month (or `through`), no AI | new |
| `fin_status(action?)` | Health, sign-in, atlas and knowledge age, restart of the hidden Excel | `sv_status`, `sv_restart`, `sv_atlas_*` |
| `xl_snapshot(range \| sheet \| workbook)` | Snapshot before an Excel MCP edit to an open workbook | part of `xl_write` and others |
| `xl_revert(id?)` | Undo to a snapshot | `xl_revert`, `xl_changes` |
| `xl_tidy(sheet)` | TN report style + tn-formelstil check (zero BLOCK findings) | `xl_format_sheet`, `xl_format`, `xl_check`, `xl_trace` |

Gone: `xl_workbooks`, `xl_overview`, `xl_read`, `xl_write`, `xl_rows`, `xl_add_sheet` (Excel MCP attach does
these). The atlas build is internal (sub-spec 2, monthly refresh).

## 2. The analysis spec (contract)

One pydantic model, strict (`extra="forbid"`), published as JSON Schema in the tool's input schema so the client
validates before calling. Periods are **relative or absolute**, so a saved spec can roll forward (D12).

```json
{"version": 2,
 "question": "Hvordan går EBITDA i Business mot i fjor og budsjett?",
 "plan": {"depth": "standard",
          "purpose": "Explain Business EBITDA YTD vs LY and budget",
          "hypotheses": ["Revenue down on fixed voice", "Opex savings in consultants"],
          "audience": "controller"},
 "queries": [
   {"id": "fm", "cube": "FinMax", "metrics": ["EBITDA", "Revenue", "COGS", "Opex"],
    "entity": "un.SY.2BUS", "periods": "closed-ytd", "compare": ["ly"]},
   {"id": "fp", "cube": "FinPlan", "metrics": ["EBITDA"], "entity": "1BUS Business",
    "periods": "closed-ytd", "scenarios": ["Actual", "Target", "Forecast:latest"]}],
 "blocks": [
   {"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA", "Revenue"], "compare": "ly"},
   {"type": "table", "title": "P&L YTD", "query": "fm", "rows": "metrics",
    "columns": ["value", "ly", "var", "var_pct"]},
   {"type": "bridge", "query": "fm", "from": "ly", "to": "value", "total": "EBITDA",
    "steps": ["Revenue", "COGS", "Opex", "Other items"]},
   {"type": "drivers", "query": "fm", "metric": "Opex", "compare": "ly", "by": "entity", "depth": 2},
   {"type": "commentary"}]}
```

- **`plan`** is required unless `depth` is `quick`. It is stored with the answer and shown to the critic.
- **`queries`** name metrics from the semantic layer (sub-spec 2) or raw members (`ac.C30`). Dimensions not named
  take the metric's defaults (currency NOK, balance from the metric). `periods`: `closed`, `closed-ytd`,
  `closed-12m`, `fy`, `Sep 26`, `Jan 26..Sep 26`; `compare`: `ly`, `budget`, `forecast:latest`.
- **Blocks** (closed list; each a pydantic class):

| Block | Renders | Options |
|---|---|---|
| `table` | Rows × columns with subtotals | `rows` (metrics / a dimension's children / members), `columns` (value, ly, budget, forecast, var, var_pct, share), `top`, `other`, `sort` |
| `trend` | Months × series + chart | `series`, `cumulative`, `chart` |
| `bridge` | Waterfall from one total to another | `from`, `to`, `steps`, `by` |
| `ranking` | Top N + Other + total, share % | `by` (entity, product, supplier, project), `top`, `group_separate` |
| `matrix` | Two dimensions crossed | `rows`, `columns`, `measure` |
| `kpi_tiles` | 2–6 headline numbers with variance | `metrics`, `compare` |
| `drivers` | Variance decomposition down a dimension, recursive to `depth`; at the leaf, optionally drill to suppliers/lines | `metric`, `by`, `depth`, `drill` |
| `lines` | Drill lines as a table | `cell`, `group_by` |
| `commentary` | Key points and drivers the agent writes after seeing values | filled by a patch |

- **Patches** (follow-ups) are JSON merge patches on the stored spec, addressed by `id` (queries) and block index or
  `key`. Unchanged queries are not fetched again.

## 3. Model layer (`engine/smartass/model/`)

Pure Python; no Excel, no network. This is where all numbers come from (D7).

1. `spec.py`: the contract above, JSON Schema export, period algebra (`closed-ytd` → concrete months given the
   closed month).
2. `resolve.py`: metrics → members per cube via the semantic layer; members validated against the atlas; returns
   the concrete `GridRequest`s and the caveats for every touched member/metric.
3. `frame.py`: fetched facts as one long table (`cube, query, account, entity, product, period, scenario, version,
   balance, currency, value, fetched_at, route`) — the same rows that become `Data_Q<n>`.
4. `expr.py`: an **expression tree** per output cell: `Fact(filters)`, `Sum([...])`, `Sub`, `Div`, `AbsDiv`,
   `Ref(cell)`. Two interpreters: `evaluate(frame)` → number, and `to_excel(layout)` → formula text
   (`SUMIFS(Data_Q3[Value],Data_Q3[Account],"ac.C30",…)`, `SUM(range)`, variance %, empty when the base is zero).
   One tree, two outputs: the sheet can only differ from the model if Excel itself differs (checked).
5. `blocks/*.py`: each block builds its trees from the frame. `drivers` ranks children by contribution to the
   variance, recurses to `depth`, and requests drills at the leaf when `drill` is set.
6. `validate.py`: the hard guards (D9) — unknown member or metric, draft (unapproved) metric used, open month
   without the `preliminary` label, a block total that does not reconcile with the cube's own parent value
   (tolerance 1 NOK), a sign that contradicts the metric. Soft findings (no LY, no top-N + Other, no source) go
   to `review`.
7. `findings.py`: checks, materiality (metric threshold from the semantic layer), material lines that need a
   driver, `key_values`.

## 4. Data layer (`engine/smartass/data/`)

- `provider.py`: `Provider.fetch(GridRequest) -> list[Fact]` and `closed_month(cube)`. Chosen per call:
  `SmartViewHttp` first, `AdhocExcel` as fallback; the route is recorded per fact.
- `smartview_http.py` + `svxml.py`: lifted from v1 `svhttp.py` / `svxml.py` unchanged in behaviour (session pool,
  read-only guard), behind the interface. Contract test against a recorded known grid detects protocol drift.
- `adhoc_excel.py`: lifted from v1 `excel/adhoc.py`.
- `drill/`: lifted from v1 `drill_*.py` (headless Edge, paging, split, reconcile, `drill.db`).
- `epm_rest.py`: placeholder interface only; implemented when Tech grants an OAuth client (sub-spec 4 tracks it).

## 5. Render layer (`engine/smartass/render/`)

- `data_sheet.py`: writes `Q<n> Data` as Excel table `Data_Q<n>` (one block write).
- `writer.py`: lays blocks out top to bottom on `Q<n> Answer`, writes `to_excel` formulas in one block per
  block, charts per block, a status line, the plan and assumptions, findings and the source line. All sheet
  text in English.
- `style.py`: TN report style (lifted from v1 `styling.py` / vendored `tn_rapportstil`).
- `readback.py`: reads the answer range back once and compares with `evaluate` (check "Sheet = model").
- `sheetorder.py`: lifted from v1.
- Packs: the spec is stored in hidden `_svmeta` per question; `fin_refresh` re-resolves relative periods to the new
  closed month, refetches and redraws each question in place, keeping the user's cells outside the owned range
  (v1 `owned.py` behaviour, lifted).

## 6. Runtime (`engine/smartass/runtime/`)

Lifted from v1 with these fixes:

| Fix | v1 evidence |
|---|---|
| Reload only on a new version file or `fin_status(restart)`, never on any `*.py` change | `run_server.py:44-55` |
| `pids.json` written atomically (temp + `os.replace`) under a file lock | `hostproc.py:89-91` |
| HTTP API secret in Windows DPAPI (`keyring`), not plain text | `setup_user.py:35,43` |
| Smart View global options restored at **startup** from the saved originals, not only at clean exit | `config.py:180`, `hostproc.py:227` |
| A failed spec save is a warning in the reply, never swallowed | `core/pipeline.py:264` |
| No `eval` anywhere (expressions are trees) | `specs.py:196` |
| `doctor` checks VBA project access (needed for the bridge) | not checked in v1 |

Kept as is: hidden Excel host on one STA thread with deadlines and PID-verified kills, user-Excel attach by
window, per-Excel turns, login wait and SSO watcher, snapshots/revert (`workbench/history.py`), logs and
`runs.jsonl`.

## 7. Errors

Every error names one next step (v1 convention). Hard-guard failures return no sheet, only the reason and the
fix (for example "metric 'Opex ex one-offs' is draft: approve it or use 'Opex'"). Provider failure on both
routes returns the sign-in or connectivity step. A partial render rolls back to the backup sheet (v1 behaviour).

## 8. Testing

- **Unit (no Excel, no EPM):** spec validation, period algebra, resolve, every block's trees, `evaluate`,
  `to_excel` formula text (golden strings), validate guards, drivers ranking.
- **Recorded facts:** real fetches for the test set are recorded once to a local folder outside git (the repo's
  no-data rule holds) and replayed, so the model and the numbers are tested offline on Tom's PC.
- **Render:** `fakexl` for layout; live render checked by the test-set runner.
- **Test set:** B01–B16 rewritten as v2 specs plus new questions to 20–30, with golden numbers; the runner reports
  numbers, guards, sheet = model and the critic score (sub-spec 2).

## 9. Migration

- New package `engine/smartass/` beside `engine/smartview_mcp/`; a second MCP entry `smartass2` during the build.
- v1 stays in use until v2 passes the test set; then one PR deletes `smartview_mcp`, renames the entry, and
  records the tool rename table in `CHANGELOG.md` with a version tag.
- Lifted modules move with their tests.

## 10. Out of scope here

Knowledge content, the monthly refresh, playbooks and the critic (sub-spec 2); the Excel pane (3); installer,
releases and stewardship (4); EPM REST implementation (when access exists).
