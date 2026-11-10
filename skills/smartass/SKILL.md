---
name: smartass
description: Pull Telenor Norge financial data from Oracle EPM Cloud (FinMax/FinLite actuals, FinPlan plan/forecast) through the read-only smartass MCP (fin_* tools) and answer with an analysis in the user's Excel (answer sheet with live formulas on a data table). Use for any number or analysis from FinMax, FinLite, FinPlan, Smart View, Hyperion or EPM - actuals, YTD, P&L, EBITDA, revenue/COGS, opex, FTE/årsverk, cost base per cost centre (un.SY), monthly trend, actual vs budget/forecast, variance drivers, monthly packs, invoice/journal lines behind a number - for metric and member names, and for safe edits to an open workbook (snapshot, revert, TN style).
---

# smartass: finance analyses in Excel

Call the 9 `mcp__smartass__*` tools directly, never a side script. Read-only towards EPM. Answers go to the
session workbook (`target="session"`, Q1, Q2 …); `"active"` only for the user's open workbook, `"new"` for a
separate file. Method lives in MCP prompts: `plan`, `critic` and the playbooks `monthly-pack`,
`variance-commentary`, `gl-extract`, `prodreg-input`.

## 1. The loop: plan → build → explain → critic

1. **Plan.** Decide cube, entity, metrics, periods and comparisons. Unsure of a name: `fin_lookup(query, kind,
   cube)` (metrics by name or alias, members, concepts; with caveats). Past answers:
   `fin_lookup(query, kind="answer")`. Write `plan`: `depth`, `title`, `purpose`, `hypotheses`.
2. **Build.** `fin_analyze(spec=...)` computes every value and writes `Q<n> Answer` (live formulas) +
   `Q<n> Data`. Unknown members, an open month or unreconciled totals stop it.
3. **Explain.** Add the `commentary` by patch (drivers with `data` evidence or labelled `hypothesis`).
4. **Critic** (`standard`, `deep`, every pack question). `fin_critic(q)` returns the brief and the rubric; score
   the 6 criteria 1–5 and store them with `fin_critic(q, scores, notes)`: the server computes the score. On
   `passed: false`, fix the gaps by patch and score again (at most 2 rounds); report the returned score.

## 2. The spec

`{"version": 2, "question", "plan", "queries": [{"id", "cube", "metrics", "entity", "periods", "compare"}],
"blocks": [...]}`. Strict: unknown fields are refused. `budget`/`forecast` only in FinPlan. Blocks: `table`,
`kpi_tiles`, `trend`, `bridge` (EBITDA: all four steps), `ranking`, `matrix`, `drivers`, `lines`, `commentary`,
`concept`, `chart`. Products: `rows: "products"`. Periods, options: [reference/spec.md](reference/spec.md);
concepts: [reference/concepts.md](reference/concepts.md).

**The answer sheet:**

- Fixed order: title (`plan.title`, the message), Key points, Charts, tables, Checks, Notes, Chart data, Source.
  The plan is not on the sheet.
- Every answer needs a `chart` block chosen by the message (bridge `waterfall`, trend `line`, AC vs PY
  `column`, ranking `bar`): [reference/charts.md](reference/charts.md). A new standard/deep question without one
  is refused.
- P&L or cost: default table `{"type":"table","query":"q","periods":["closed","closed-ytd"]}` (month + YTD;
  query periods `closed-ytd`). A table with EBITDA gets the whole P&L chain.
- Key points: at most 5; the engine writes the material lines, your commentary drivers give the cause.
- The reply's `layout` findings each carry a fix: apply it by patch.

## 3. After the first answer

- **Follow-up:** `fin_analyze(patch=..., q=n)`: JSON merge patch on the stored spec (lists replaced whole).
  Follow-ups, `fin_critic` and `fin_refresh` default to the latest answer's workbook; target an open one by name
  (`target="Book1"`).
- **Monthly pack:** `fin_refresh()` re-runs every answer to the latest closed month; then prompt `monthly-pack`.
- **Ledger lines:** a `lines` block or `fin_drill(cell, group_by, top)` (ties to the cube, report the gap;
  60–90 s): [reference/drill.md](reference/drill.md).

## 4. Reporting

- Quote numbers **only from `key_values`**; never compute them yourself.
- Pass on, Critical first: `findings`, checks not OK, `warnings` (rejected members and error cells are missing
  numbers, never zero), the `material` lines with their drivers, and the `caveats` that bear on the answer.
- Chat in the user's language; sheet text is English. The user's workbook is never saved.

## 5. Editing open workbooks (Excel MCP)

Before an edit `xl_snapshot(workbook, sheet, range)`; undo with `xl_revert(change_id)`. Finish with
`xl_tidy(workbook, sheet, rows, last_col)`: TN style and the tn-formelstil check, zero BLOCK findings.

## 6. Finance rules

- **Cube:** actuals FinMax, drill FinLite, plan only FinPlan.
- **Format and signs:** full NOK `1 000 000`, negatives `(1 000 000)`; FinMax signs (costs negative); variance = current - comparison (`fin_lookup` kind 'rule': full rules).
- **P&L chain (FinMax):** C30 + C40S = C48; C50S + C60 = C70; C48 + C70 = EBITDA_BOI; + C73 = EBITDA.
  (fin_lookup kind 'rule' has the full rules.)
- **Members:** `un.SY.<code>`, `ac.<code>`, `Mon YY`; history restated to today's organisation:
  [reference/cubes.md](reference/cubes.md). Never write to the Produktregnskapet archive.

## 7. Health

`fin_status()`; repeated errors: `fin_status(restart=True)`. Knowledge refresh: the
user runs `smartass knowledge refresh`, reads the report, then `smartass knowledge approve <id|all>`. A parameter
missing after a reload: ask for `/mcp`, reconnect smartass. [reference/operations.md](reference/operations.md).
