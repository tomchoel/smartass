---
name: smartass
description: Pull Telenor Norge financial data from Oracle EPM Cloud (FinMax/FinLite actuals, FinPlan plan/forecast) through the read-only smartass MCP (fin_* tools) and answer with an analysis in the user's Excel (answer sheet with live formulas on a data table). Use for any number or analysis from FinMax, FinLite, FinPlan, Smart View, Hyperion or EPM - actuals, YTD, P&L, EBITDA, revenue/COGS, opex, FTE/årsverk, cost base per cost centre (un.SY), monthly trend, actual vs budget/forecast, variance drivers, monthly packs, invoice/journal lines behind a number - for metric and member names, and for safe edits to an open workbook (snapshot, revert, TN style).
---

# smartass: finance analyses in Excel

Call the 9 `mcp__smartass__*` tools directly, never a side script. Read-only towards EPM. Answers go to the
session workbook (`target="session"`, Q1, Q2 …); `"active"` only for the user's open workbook, `"new"` for a
separate file. Method lives in MCP prompts: `plan`, `critic` and the playbooks `monthly-pack`,
`variance-commentary`, `gl-extract`, `prodreg-input` (also resources `smartass://playbooks/<name>`).

## 1. The loop: plan → build → explain → critic

1. **Plan.** Decide cube, entity, metrics, periods and comparisons. Unsure of a name: `fin_lookup(query, kind,
   cube)` (metrics by name or alias, e.g. `driftskostnader`, atlas members, product concepts; with caveats).
   Past answers: `fin_lookup(query, kind="answer")`. Write `plan`: `depth`, `purpose`, `hypotheses`.
2. **Build.** `fin_analyze(spec=...)` fetches, computes every value and writes `Q<n> Answer` (live formulas) +
   `Q<n> Data` (table `Data_Q<n>`). Unknown members, an open month or totals that do not reconcile stop it.
3. **Explain.** Add the `commentary` by patch (drivers with `data` evidence or labelled `hypothesis`).
4. **Critic** (`standard`, `deep`, every pack question). `fin_critic(q)` returns the brief and the rubric; score
   the 6 criteria 1–5, fix the gaps by patch while the mean is below 4 (at most 2 rounds), then store it with
   `fin_critic(q, score, notes)`.

## 2. The spec

`{"version": 2, "question", "plan", "queries": [{"id", "cube", "metrics", "entity", "periods", "compare"}],
"blocks": [...]}`. Strict: unknown fields are refused. Periods `closed`, `closed-ytd`, `fy`, `Sep 26`, ranges.
`budget`/`forecast` only in FinPlan. Blocks: `table`, `kpi_tiles`, `trend`, `bridge` (an EBITDA bridge uses
Revenue, COGS, Opex, Other items), `ranking`, `matrix`, `drivers`, `lines`, `commentary`, `concept`. Products:
query `products`, `rows: "products"`. Options, example, variance %:
[reference/spec.md](reference/spec.md); concepts: [reference/concepts.md](reference/concepts.md).

## 3. After the first answer

- **Follow-up:** `fin_analyze(patch=..., q=n)`: JSON merge patch on the stored spec (lists replaced whole).
  Follow-ups, `fin_critic` and `fin_refresh` default to the workbook of the latest answer; an open workbook can
  be targeted by name (`target="Book1"`).
- **Monthly pack:** `fin_refresh()` re-runs every answer to the latest closed month; then prompt `monthly-pack`.
- **Ledger lines:** a `lines` block or `fin_drill(cell, group_by, top)` (FinLite, Periodic; ties to the cube,
  report the gap; 60–90 s): [reference/drill.md](reference/drill.md).

## 4. Reporting

- Quote numbers **only from `key_values`**; never compute sums, differences or shares yourself.
- Pass on, Critical first: `findings`, checks not OK, `warnings` (rejected members and error cells are missing
  numbers, never zero), the `material` lines with their drivers, and the `caveats` that bear on the answer.
- Chat in the user's language; sheet text is English. The user's workbook is never saved.

## 5. Editing open workbooks (Excel MCP)

Before an edit `xl_snapshot(workbook, sheet, range)`; undo with `xl_revert(change_id)`. Finish with
`xl_tidy(workbook, sheet, rows, last_col)`: TN style and the tn-formelstil check, zero BLOCK findings.

## 6. Finance rules

- **Cube:** actuals FinMax (FTE `ac.STA20Y`); FinLite lighter (no FTE; drill source); plan only FinPlan.
- **Format:** full NOK `1 000 000`, negatives `(1 000 000)` in tables, `−1 000 000` in text; FinMax signs (costs
  negative); variance = current − comparison (positive = better); variance % empty on a zero or tiny base (the
  floor sits in the sheet's Thresholds).
- **P&L chain (FinMax):** C30 + C40S = C48; C50S + C60 = C70; C48 + C70 = EBITDA_BOI; + C73 = EBITDA.
- **Members:** `un.SY.<code>`, `ac.<code>`, `Sep 26`; history restated to today's organisation:
  [reference/cubes.md](reference/cubes.md). Never write to the Produktregnskapet archive.

## 7. Health

`fin_status()` (version, atlas, knowledge); repeated errors: `fin_status(restart=True)`. Knowledge refresh: the
user runs `smartass knowledge refresh`, reads the report, then `smartass knowledge approve <id|all>`. A parameter
missing after a reload: ask for `/mcp` and reconnect smartass. [reference/operations.md](reference/operations.md).
