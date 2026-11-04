---
name: smartass
description: Pull Telenor Norge financial data from Oracle EPM Cloud (FinMax/FinLite actuals, FinPlan plan/forecast) through the read-only smartass MCP (fin_* tools) and answer with an analysis in the user's Excel (answer sheet with live formulas on a data table). Use for any number or analysis from FinMax, FinLite, FinPlan, Smart View, Hyperion or EPM - actuals, YTD, P&L, EBITDA, revenue/COGS, opex, FTE/årsverk, cost base per cost centre (un.SY), monthly trend, actual vs budget/forecast, variance drivers, monthly packs, invoice/journal lines behind a number - for metric and member names, and for safe edits to an open workbook (snapshot, revert, TN style).
---

# smartass: finance analyses in Excel

Call the `mcp__smartass__*` tools directly, never a side script. Read-only towards EPM. Answers go to the
session workbook (`target="session"`, questions Q1, Q2 …); `"active"` only when the user asks for *their* open
workbook, `"new"` for a separate file.

## 1. The loop: plan → build → critic

1. **Plan.** Decide cube, entity, metrics, periods and comparisons. Unsure of a name: `fin_lookup(query, kind,
   cube)` (metrics by name or alias, e.g. `driftskostnader`, and atlas members; each hit with its caveats). Write `plan`: `depth` (`quick` | `standard` | `deep`),
   `purpose` (required unless quick), `hypotheses` you will test.
2. **Build.** `fin_analyze(spec=...)`: one call fetches (Smart View HTTP; an ad hoc grid in Excel as fallback),
   computes every value and writes `Q<n> Answer` (live formulas) + `Q<n> Data` (table `Data_Q<n>`). Unknown
   members or metrics, an open month or totals that do not reconcile stop it with the reason and the fix.
3. **Critic** (skip for `quick`). Add the `commentary` by patch, then `fin_critic(q)`: score the brief on the
   rubric, fix by patch below 4 (2 rounds), store with `fin_critic(q, score, notes)`.

## 2. AnalysisSpec v2 essentials

```json
{"version": 2, "question": "Wholesale P&L year to date vs last year",
 "plan": {"depth": "standard", "purpose": "Explain the Wholesale EBITDA bridge YTD vs LY"},
 "queries": [{"id": "q1", "cube": "FinMax", "metrics": ["Revenue", "COGS", "Opex", "Other items", "EBITDA"],
              "entity": "un.SY.2WHS", "periods": "closed-ytd", "compare": ["ly"]}],
 "blocks": [{"type": "kpi_tiles", "query": "q1", "metrics": ["Revenue", "EBITDA"]},
            {"type": "table", "query": "q1", "title": "P&L"},
            {"type": "bridge", "query": "q1", "total": "EBITDA", "steps": ["Revenue", "COGS", "Opex", "Other items"]}]}
```

- **Strict:** unknown fields are refused; a block's metrics and series must be in its query.
- **periods:** `closed`, `closed-ytd`, `closed-12m`, `fy`, `Sep 26`, `Jan 26..Sep 26`. Relative periods roll
  forward in `fin_refresh`. Open month: `preliminary: true`.
- **compare:** `ly`; `budget` and `forecast` (latest) **only in FinPlan**: actual vs plan is done entirely in
  FinPlan (entity alias e.g. `1TEC Technology`; columns `["value", "budget", "var", "forecast", "var"]`).
- **Blocks:** `table` (rows `metrics` or `entities` + `metric`; columns value, ly, budget, forecast, var,
  var_pct, share; `top`, `other`, `sort`), `kpi_tiles` (2–6), `trend` (series, `cumulative`, chart), `bridge`
  (`total`, `steps`, `base`; an EBITDA bridge uses every step: Revenue, COGS, Opex, Other items), `ranking`
  (`metric`, `top`), `matrix`, `drivers` (`metric`, `depth` 1–4, `drill` to ledger lines), `lines` (`cell`,
  `group_by` supplier | project | account, `top`; needs no query), `commentary` (`key_points`, `drivers`).
- **Commentary drivers:** `{"line", "text", "basis": "data", "evidence": "<cell key>"}` or
  `"basis": "hypothesis"`, written from the reply's evidence, never from memory.

## 3. After the first answer

- **Follow-up:** `fin_analyze(patch=..., q=n)`: a JSON merge patch on the stored spec (lists such as `blocks` are
  replaced whole); default q = the active `Q<n> Answer`, else the latest. Same sheets, rebuilt in place.
- **Monthly pack:** `fin_refresh()` re-runs every answer in the workbook to the latest closed month, no AI needed;
  then update the commentary by patch.
- **Ledger lines:** a `lines` block, or `fin_drill(cell, group_by, top)` with
  `{"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"}` (FinLite, Periodic;
  `Yr 26` works). Ranks suppliers, projects or accounts and ties the lines to the cube: report the gap. 60–90 s;
  wait for the progress notes. Details: [reference/drill.md](reference/drill.md).

## 4. Reporting

- Quote numbers **only from `key_values`**; never compute sums, differences or shares yourself.
- Pass on, Critical first: `findings`, checks not OK, `warnings` (rejected members and error cells are missing
  numbers, never zero), the `material` lines with their drivers.
- Read the reply's `caveats` (reorg moves, abolished units, account traps; also on the sheet) and pass on those
  that bear on the answer.
- Chat in the user's language; sheet text is English. Name the sheets; the user's workbook is never saved.

## 5. Editing open workbooks (Excel MCP)

Every other Excel job (read, write, format, pivots, charts) is Excel MCP on the attached workbook. Before an
edit `xl_snapshot(workbook, sheet, range)`; undo with `xl_revert(change_id)`. Finish with
`xl_tidy(workbook, sheet, rows, last_col)`: TN style (given `rows` + `last_col`) and the tn-formelstil check,
zero BLOCK findings.

## 6. Finance rules

- **Cube:** actuals FinMax (FTE `ac.STA20Y`); FinLite lighter (no FTE; drill source); plan only FinPlan.
- **Format:** full NOK `1 000 000`, negatives `(1 000 000)` in tables, `−1 000 000` in text; FinMax signs (costs
  negative); variance = current − comparison (positive = better); variance % empty when the base is zero.
- **P&L chain (FinMax):** C30 + C40S = C48; C50S + C60 = C70; C48 + C70 = EBITDA_BOI; + C73 = EBITDA.
- **Members:** `un.SY.<code>` (2WHS, 2BUS, 2SSM, 2EI, 2TEC …), `ac.<code>`, `Sep 26`. History is restated to
  today's organisation (the 1 June 2026 reorg emptied 2CON into 2SSM + 2EI). Tops: [reference/cubes.md](reference/cubes.md).
- Never write to the Produktregnskapet archive.

## 7. Health and reloads

`fin_status()`: version, atlas age, knowledge, doctor problems. Repeated errors: `fin_status(restart=True)`.
Stale atlas: `smartass atlas build` in a terminal. Monthly, when it says knowledge refresh due: the user runs
`smartass knowledge refresh`, reads the report, then `smartass knowledge approve <id|all>`. After a reload Claude Code may keep an old
tool schema: if a parameter seems missing, ask the user to run `/mcp` and reconnect smartass.
Errors, files and logs: [reference/operations.md](reference/operations.md).
