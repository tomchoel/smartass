---
name: smartass
description: Pull Telenor Norge financial data from Oracle EPM Cloud (FinMax/FinLite actuals, FinPlan plan/forecast) through the read-only smartass MCP (sv_* tools) and answer with an analysis in the user's Excel - source data sheets plus a formatted answer sheet with live formulas. Use for any number or analysis from FinMax, FinLite, FinPlan, Smart View, Hyperion or EPM - actuals, YTD, P&L, EBITDA, revenue/COGS, opex, depreciation, FTE/årsverk, cost base per cost centre (un.SY), monthly trend, actual vs budget/forecast, GL extract/uttrekk, invoice/journal lines behind a number - and for finding member names.
---

# Smart View analyses (FinMax · FinLite · FinPlan)

## 1. When to use

Any number, table or analysis from EPM (FinMax, FinLite, FinPlan), a member name lookup, or the ledger lines behind a
FinLite cell. The `smartass` MCP is read-only towards EPM and works **inside the user's Excel**. Call the
`mcp__smartass__*` tools directly, never a side script that imports the package. Code changes reload by themselves
on the next call; if `sv_gate` or `sv_answer` is missing from the tool list, ask the user to type `/mcp` →
smartass → Reconnect once.

- **Default target: the session file.** Every session keeps its own workbook,
  `<analyser folder>\<project folder>\<date time> <first question>.xlsx` (`paths.analyser` in
  `%LOCALAPPDATA%\smartass\config.toml`). The first question creates it; later questions become Q2, Q3 … in it. If
  it was closed, it is reopened.
- **Other targets.** `target='active'` only when the user says to put it in *their* open workbook; `target='new'`
  for an explicitly separate file.
- **Parallel sessions** are fine: only writing in the same Excel window takes turns.

## 2. The pipeline: gate → answer → follow-ups → commentary

### 2.1 Gate (always first)

`sv_gate(question)` runs the fixed scope checklist (period, scenario, entity, measure, currency/sign, granularity)
against the catalogue, the capability map and the user's learned defaults. Procedure: `knowledge/advisor/gate.md`.

| Outcome | Do |
|---|---|
| `clear` | Proceed. State every item with source convention/learned (and each warning) as an assumption. Call `sv_answer` now with the gate's `recipe` (`{name, params}`); `recipe: null` → a free-form spec. |
| `ambiguous` | Pull nothing. Ask its questions (1–3) in ONE message, in the user's language. Then call `sv_gate(question, answers={item: answer})` with the SAME question (e.g. `{"measure": "opex"}`), pass the same `answers` as `spec.gate.answers`, and offer `sv_remember(item, phrase, value)` to save an answer as the user's default. |
| `infeasible` | Say why in one or two sentences, offer the proxy, and build it labelled as a proxy. |

Copy the result into the spec's `gate` (`outcome`, `checklist`, `questions` / `reason` + `proxy`, `answers`) and
`assumptions`. Use the gate's `closed_month`: the month after it is open and carries reversals.

**Always build sheets.** Every number question ends in answer sheets built by `sv_answer`: never ask whether to
build a sheet, and never answer with a table in chat only. Several sheets per request are fine (`Q<n> Answer`, a
topic sheet such as `Q<n> By project`, source and Lines sheets).

**Departures need the user.** Use the recipe and params the gate named. Another recipe, a free-form spec where the
gate named one, or a param that switches off a finance default (`top` ≠ 10, `compare` without last year,
`group_separate: false` …) only when the user explicitly asked for it: pass `deviation="<the user's request>"` in the
spec; it is written on the sheet ("Deviation from the standard analysis (user's request): …") and in the run log.

**Language:** the chat is in the user's language; sheet text is English (title, subtitle, notes, key points,
section titles, column headers, check labels, assumptions, drivers, deviation). Supplier, project and member names
stay as they are; the question is shown as asked.

### 2.2 Answer (one call)

`sv_answer(spec)` writes `Q<n> Answer` first (question, assumptions, a status line), fetches each pull over HTTP
(Smart View ad hoc grid in Excel if HTTP fails) into a source sheet, renders the answer with live formulas and checks
it. Spec v1 (`engine/smartview_mcp/core/contract.py`):
`{version: 1, question, gate, assumptions, pulls: [{id, connection, rows, columns, pov, title?}], layout: {title,
subtitle?, columns, sections, checks, key_points?, notes?, source?, chart?}, commentary?}`.

Worked example (bench B02: Business EBITDA per month 2026 vs 2025, with a reconciliation to Accumulated):

```json
{"version": 1, "question": "EBITDA per month for Business 2026 vs 2025",
 "gate": {"outcome": "clear", "checklist": {"period": "Jan-Sep 26 vs Jan-Sep 25", "scenario": "Actual",
          "entity": "un.SY.2BUS", "measure": "ac.EBITDA", "currency_sign": "NOK, FinMax signs",
          "granularity": "month"}},
 "assumptions": ["Closed month Sep 26."],
 "pulls": [
  {"id": "m", "connection": "FinMax",
   "rows": [{"dimension": "All Accounts", "members": ["ac.C30", "ac.C40S", "ac.C70", "ac.C73", "ac.EBITDA"]}],
   "columns": [{"dimension": "All Time Periods", "members": ["Jan 25", "…", "Dec 25", "Jan 26", "…", "Sep 26"]}],
   "pov": {"All Entities": "un.SY.2BUS", "All Scenarios": "Actual", "All Currencies": "NOK",
           "All Balance Types": "Periodic"}},
  {"id": "ytd", "connection": "FinMax", "title": "Q1 FinMax YTD",
   "rows": [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
   "columns": [{"dimension": "All Time Periods", "members": ["Sep 26", "Sep 25", "Dec 25"]}],
   "pov": {"All Entities": "un.SY.2BUS", "All Scenarios": "Actual", "All Currencies": "NOK",
           "All Balance Types": "Accumulated"}}],
 "layout": {
  "title": "Business Norway — EBITDA per month 2026 vs 2025", "subtitle": "NOK · FinMax · Actual · un.SY.2BUS",
  "columns": [{"header": "Actual 2026"}, {"header": "Actual 2025"}, {"header": "Variance"},
              {"header": "Variance %", "kind": "prosent1"}],
  "sections": [
   {"title": "EBITDA per month (Periodic)",
    "repeat": {"over": "months:Jan..Dec", "line": {"key": "m{i}", "label": "{v}", "cells": [
      {"pull": "m", "row": ["ac.EBITDA"], "col": ["{v} 26"]}, {"pull": "m", "row": ["ac.EBITDA"], "col": ["{v} 25"]},
      {"var": [0, 1]}, {"varpct": [0, 1]}]}},
    "lines": [{"key": "ytd", "label": "YTD Jan–Sep", "kind": "grand",
               "cells": [{"sum": "m[0-8]"}, {"sum": "m[0-8]"}, {"var": [0, 1]}, {"varpct": [0, 1]}]},
              {"key": "fy", "label": "Full year 2025", "kind": "sum", "cells": [null, {"sum": "m*"}, null, null]}]},
   {"title": "Reconciliation (Accumulated)", "lines": [
     {"key": "eacc", "label": "EBITDA YTD Sep", "cells": [{"pull": "ytd", "row": ["ac.EBITDA"], "col": ["Sep 26"]},
                                                         {"pull": "ytd", "row": ["ac.EBITDA"], "col": ["Sep 25"]},
                                                         null, null]},
     {"key": "fyacc", "label": "EBITDA Dec 25",
      "cells": [null, {"pull": "ytd", "row": ["ac.EBITDA"], "col": ["Dec 25"]}, null, null]}]}],
  "checks": [{"label": "Sum of months = Accumulated YTD", "expr": "ytd-eacc", "columns": [0, 1]},
             {"label": "Sum of 2025 months = Accumulated Dec 25", "expr": "fy-fyacc", "columns": [1]}],
  "source": "Source: Oracle EPM Cloud (FinMax) via Smart View.",
  "chart": {"type": "column", "lines": "m*", "columns": [0, 1], "title": "EBITDA per month"}}}
```

- **Cells:** `pull` (+ `row`/`col` members), `var`, `varpct`, `sum` (a key pattern such as `m*` or `m[0-8]`), `avg`,
  `expr`, `ref`, `rowsum`, `value`, `text`, `drill`. `repeat` makes one line per month or member; months not
  pulled stay empty.
- **Plan the pulls yourself** (this year + last year, FinMax + FinPlan) and add the same period last year to P&L and
  cost questions by default.
- **Input shapes.** `rows`/`columns` are `[{dimension, members}]`.
- The result carries `checks`, `review` (findings), `material`, `outline`, `routes` (HTTP or Smart View per
  source sheet), `timings` (seconds per stage: mention an unusually slow one; a drilling recipe adds `first frame`
  and `drills`), `formula_style` (skill `tn-formelstil`; if `ok` is false, a `value`/`expr` you wrote broke a rule:
  fix it with a follow-up), `key_values` (line label → column header → the value the sheet calculated, at most 40
  lines), `followups` (at most 3 next questions) and `sheets` (every sheet written).
- **Enforced before anything is pulled** (each refusal names the fix; [reference/operations.md](reference/operations.md)
  "Errors"): the gate is re-run on `spec.question` with `spec.gate.answers` and a still-ambiguous question returns
  its questions instead; a departure from the gate's recipe or a recipe default needs `deviation`; sheet text must
  be English.

### 2.3 Follow-ups

`sv_answer(followup=<patch>, q=<n>)` changes the current answer (default: the active `Q<n> Answer`, else the
latest): same sheet, unchanged pulls not fetched again, the user's cells and charts outside the answer block kept.
Ask a new question only when the user does. Address parts by the `outline` of the last result (pull ids, column
headers, section titles, line keys, check labels). Patch semantics: [reference/operations.md](reference/operations.md)
"Recipes and follow-ups".

- Add a month: re-send the pull with the longer member list:
  `{"pulls": [{"id": "m", "columns": [{"dimension": "All Time Periods", "members": ["Jan 25", "…", "Oct 26"]}]}]}`.
- Swap a scenario: `{"pulls": [{"id": "fp", "pov": {"Scenario": "Forecast Q326"}}]}`.
- Text: `{"layout": {"title": "Business EBITDA 2026", "notes": {"append": ["Excludes one-offs."]}}}`.
- Drop a line: `{"layout": {"sections": [{"title": "Per month", "lines": [{"key": "a1", "drop": true}]}]}}`.

### 2.4 Recipes

Instead of `pulls` + `layout`, give `{"recipe": {"name", "params", "overrides"?}}` (with `version`, `question`,
`gate`). The server drafts the spec; the period defaults to the closed month (returned as `period`). `overrides` is a
follow-up-shaped patch on the draft; `params` cannot change in a follow-up (ask a new question).

| Recipe | Params | Delivers |
|---|---|---|
| `pnl` | `area`, `period?`, `compare?`, `detail?` (`compact`/`full`) | Month and YTD vs last year, budget and the latest forecast; materiality; EBITDA bridge by line and sub-unit; KPIs; year-end landing. Read live; check "Live read = source sheet". |
| `actual_vs_plan` | `entity` (FinPlan alias), `month?`, `year?`, `basis?`, `forecast?`, `finmax_entity?` | Actual vs Target and the latest forecast in FinPlan, with a FinMax EBITDA reconciliation |
| `monthly_trend` | `entity?`, `accounts?`, `through?`, `compare?`, `connection?` | Jan→month periodic with the same months last year, a YTD column checked against Accumulated, line chart |
| `cost_base` | `parent?`, `accounts?`, `period?`, `top?`, `connection?` | Top 10 leaf cost centres + Other + total, share %, last year |
| `concept` | `concept`, `period?`, `compare?`, `entity?` | A concept from the library: every block, keys, FinMax basis |
| `lines_ranking` | `accounts`, `entity?`, `period?`, `by` (`supplier`), `projects?`, `top?`, `compare?`, `group_separate?` | Drills the FinLite ledger lines itself (this year + last year) and ranks invoiced amounts per supplier: top 10 + Other, group companies as their own block, a reconciling line for lines without a supplier, an "Unexplained gap" line, tied to the cube; `projects: true` adds `Q<n> By project` |

- **Finance defaults are in each recipe** (same period last year, top 10 + Other + total + share %, group
  companies separate, reconciliation to the cube); turning one off is a departure (`deviation`, §2.1). The gate's
  recipe table and every default: [reference/recipes.md](reference/recipes.md).
- **Line-level questions** ("which suppliers / consultancies invoiced most, on which projects") use
  `lines_ranking` in ONE `sv_answer` call: it draws the first frame, drills (about 60 s per year, with progress
  notes) and builds `Q<n> Answer` (+ `Q<n> By project`, `Q<n> Lines d1` / `d2`). Never `sv_drill` plus your own
  sums. If the drill fails, the answer shows the cube total only, with a warning.

- **`pnl` area:** a business area by name (Wholesale, Business, Technology …) or any FinMax unit below one
  (`un.SY.5OK_W`, or the bare code `5OK_W`). Consumer and IG (abolished 1 June 2026) are refused with their
  successors: tell the user and ask which successor(s). Most units have no FinPlan member, so they get last year
  only (pass the warning on).
- **Product / profitability questions:** check the business map and concept library (below). A matching concept →
  recipe `concept`; several family concepts → one answer per family, plus their sum in the reply. No concept:
  research the members, PROPOSE the blocks (included / estimated with key / left out), save the concept in
  `knowledge\concepts.json` once the user confirms, then run it.

### 2.5 Checks, findings and commentary

- Report every `checks` value that is not OK and every warning. A `Warning`-severity DIFF is a known definitional
  difference (FinPlan vs FinMax), not a wrong number.
- `review` lists findings (Critical / Warning / Info; `knowledge/advisor/findings.md`). Pass Critical ones on first.
- `material` lists the lines that need a driver. Explain them in `commentary.drivers` (`knowledge/advisor/
  variance-commentary.md`): `basis: data` with the pull/drill cell as `evidence`, otherwise `basis: hypothesis`,
  said as "Hypothesis — not verified in data". Never figures from memory.
- **Report in chat, in the user's language, as a full summary:** the key points; a compact table of the headline
  lines; cube · POV · period on one line; the checks and findings above; the sheet names (`sheets`); the
  `followups` (at most 3) as suggestions; that the user's workbook was **not saved** (session and new files are the
  MCP's own and are saved).
- **Numbers in chat:** quote `key_values` (the values the sheet calculated). Never compute numbers yourself (no
  sums, differences or shares of your own); a number not in `key_values` is read with `xl_read`.

## 3. Utilities

- **Members:** `sv_find_member(text)`, `sv_members(connection, member, relation)`, `sv_member_info(connection,
  member)`, `sv_dimensions(connection)` (local atlas of all three cubes, instant). FinMax names are prefix + code:
  `un.SY.2TEC`, `ac.C30`, `pd.SY.P56`. Entities default to the `un.SY` tree; `un.NOR.*` is the management view.
  FinPlan uses aliases as shown (`2TEC Technology`, `FY26`, `Target 2025`).
- **Drill (invoice / journal lines, FinLite only), a workbench utility:** rankings by supplier or project use the
  recipe `lines_ranking` (§2.4). For other line-level reports: `sv_drill(question, cell={'All Accounts': ..,
  'All Entities': .., 'All Time Periods': ..})`. Hidden browser, never Smart View's own drill-through. One drill at
  a time across sessions; 60–90 s with the silent sign-in; big cells split themselves. Read `reconciliation` (must
  be `exact`, or explain the gap) and `profile`, then build the report with `sv_answer` using `{"drill": "d",
  "where": {...}}` cells (a new spec with drill cells answers the pending drill's `Q<n>`). Details:
  [reference/drill.md](reference/drill.md).
- **The xl workbench** (any open workbook, no Smart View): `xl_workbooks`, `xl_overview`, `xl_read` to look;
  `xl_write`, `xl_rows`, `xl_add_sheet`, `xl_format`, `xl_format_sheet` to change (each snapshotted; `xl_changes`,
  `xl_revert` undo); `xl_trace`, `xl_check` to explain and audit formulas. Never saved by the tools.
- **Health:** `sv_status` (engine, atlas, sign-in, drill browser), `sv_restart` (restarts only the hidden Excel),
  `sv_atlas_build` / `sv_atlas_status` / `sv_atlas_changes` (atlas older than 30 days, or after a reorg).

## 4. Sheets it writes

- **`Q<n> Answer`**: the answer in TN report style (key points, table, checks, findings, notes, source, chart).
  All numbers are live formulas on the source sheets.
- **`Q<n> FinMax` / `Q<n> FinPlan` / `Q<n> FinLite`**: one source sheet per pull (a value grid; a follow-up
  re-fetches).
- **`Q<n> <topic>`** (e.g. `Q<n> By project`): a recipe's topic sheet, hyperlinked to and from `Q<n> Answer`.
- **`Q<n> Lines`**: drill lines (Excel table `Lines_Q<n>`); `lines_ranking` writes `Q<n> Lines d1` (this year) and
  `Q<n> Lines d2` (last year), tables `Lines_Q<n>_d1` / `_d2`.
- **`Sources`**: one row per pull (fetched, question, cube, POV, route). A user's own unrelated `Sources` sheet
  pushes ours to `Sources SV`.
- **Order** is fixed after every answer and drill: all `Q<n> Answer`, the user's other sheets, data per question,
  `Sources` last. Do not reorder by hand.

## 5. Errors

Every error names one next step. The table of messages (sign-in, hung Login window, `Empty response from provider`,
busy, rejected members, drill) is in [reference/operations.md](reference/operations.md).

- **Sign-in.** A Login window in the user's Excel needs the user (or the SSO watcher) to click 'Telenor SSO'; the
  engine says so within about 5 s and never closes the user's Excel. A hung Login window in the engine's own hidden
  Excel is restarted automatically.
- **Long calls** send progress notes (`still working: <stage> (n s)`): wait for them, do not call again.

## 6. Rules

### Pick the cube (decided with the user)

- **Actuals only** come from **FinMax** (the ledger, finest detail, FTE `ac.STA20Y`). FinLite is lighter: no FTE and
  no Functions/Projects/Reserves/Systems.
- **Budget or forecast only** comes from **FinPlan**.
- **Actual vs budget/forecast** is done **entirely in FinPlan** (Actual = version `Working`; Target and the latest
  `Forecast Q<q><yy>` = version `Final`), with a reconciliation line to FinMax EBITDA. FinPlan groups accounts
  differently (Technology revenue YTD Sep 26: 59.8 MNOK FinMax, 67.2 MNOK FinPlan).

### Financial format (sheet and chat)

- **Tables:** full NOK `1 000 000`; negatives `(1 000 000)`; zero `–`; FTE and % with 1 decimal.
- **Running text:** a minus sign, `−151 328 997`, never double parentheses.
- **Signs:** keep FinMax signs; costs are negative. Variance = current − comparison (positive = better, for costs
  too). Variance % = variance / |comparison|. Cost per FTE is shown positive.
- **Line labels** in English as FinMax shows them (`C30 REVENUES`, `C70 Total Operating Expenses`).

### P&L chain

FinMax: C30 + C40S = `C48`; C50S + C60 = `C70`; C48 + C70 = `EBITDA_BOI`; + `C73` = **`EBITDA`**. `C75` is
depreciation. FTE is `STA20Y` (end-of-period stock: `Accumulated` = level, `Periodic` = the month's change).
FinPlan: Revenue + COGS = Gross_Profit; + OPEX = EBITDA before Other Items; + Total Other Items = EBITDA. Target
carries no depreciation.

### Traps

- **History is restated to today's organisation.** The 1 June 2026 reorg emptied `2CON`, `2CWK` and `3WHS_C` and
  moved `4REV_W` to `3WHS_W`. Old extracts match on leaves and the `un.SY` total, not on division splits.
- **A `Yr` member before year-end is year-to-date.** The January Dec cut is preliminary.
- **The archive is protected.** Files under `Analyse (Telenor Norge AS) - Produktregnskapet\<year>\` are refused
  unless the user explicitly allows it for that request (`allow_archive=true`). Archive `GL` sheets are pasted values.
- **Rejected members.** A member Smart View rejects is left out and named in `warnings`: report those lines as
  missing. An unknown member: find the right name with `sv_find_member`.
- **Error cells** (`#NoAccess`, `#Invalid`, `#VALUE!`) are reported as warnings: those numbers are missing, never
  zero.
- **FinPlan KPIs** (volumes, ARPU, %) are read `Periodic` per month; `Accumulated` sums the months for many of them.
  Actual volumes are not in FinPlan for `BANORWAY`.

### Business map (forretningskart)

The repo's `knowledge\` folder (plus `packs\<pack>\knowledge\`) holds what the business MEANS in FinMax and FinPlan;
the gate reads it.

| Area | File | Contents |
|---|---|---|
| Accounts and bookings | `accounts_map.json` | 234 P&L groups: meaning, product/reserve tagging, division, pitfalls |
| Products | `products_map.json` | 496 pd.SY groups: P&L, selling division, Produktregnskap codes and family |
| Organisation | `org_map.json` | All un.SY entities: division 2026 vs 2025 (reorg), cost base, FTE, pitfalls |
| FinPlan | `finplan_map.json` | 654 P&L, KPI and ratio lines: actual, Target, forecast, FinMax equivalent |
| Concepts | `concepts.json` | `splitt` (hand-made, with keys) + 22 product families `fam_*` (booked only) |

- Revenue is product-tagged; salaries (C50S) are never; depreciation about 30 %, mostly in Technology. Consumer
  fixed broadband/TV has no fixed Produktregnskap code. The Functions dimension is unused.
- Rebuild the maps with `knowledge\build_*_map.py` after the monthly atlas refresh, then
  `build_family_concepts.py`. Members added since the user's approval are flagged: ask the user about those, and
  record approval with `python review.py approve <accounts|products|org|finplan> "<comment>"`.

The team's recurring extracts (GL extract, revenue/COGS, depreciation control, FinLite special costs) as pull specs:
[reference/recipes.md](reference/recipes.md). Top levels of every dimension: [reference/cubes.md](reference/cubes.md).
Operations, files, logs and errors: [reference/operations.md](reference/operations.md).
