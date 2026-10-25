---
name: smartass
description: Pull Telenor Norge financial data from Oracle EPM Cloud (FinMax/FinLite actuals, FinPlan plan/forecast) through the read-only smartass MCP (sv_* tools) and answer with an analysis in the user's Excel - live Smart View data sheets plus a formatted answer sheet. Use for any number or analysis from FinMax, FinLite, FinPlan, Smart View, Hyperion or EPM - actuals, YTD, P&L, EBITDA, revenue/COGS, opex, depreciation, FTE/årsverk, cost base per cost centre (un.SY), monthly trend, actual vs budget/forecast, GL extract/uttrekk - for finding member names, and for refreshing a workbook's Smart View sheets.
---

# Smart View analyses (FinMax · FinLite · FinPlan)

The `smartass` MCP is read-only towards EPM. It works **inside the user's Excel**.
- **Default target: the session file.** Every session (Claude Code or Copilot) keeps its own analysis workbook,
  `<analyser folder>\<project folder>\<date time> <first question>.xlsx`, where the analyser folder is `paths.analyser` in
  `%LOCALAPPDATA%\smartass\config.toml`. It is created on the session's first question,
  and every later question of that session goes into it as Q2, Q3 and so on. If it was closed, it is reopened.
- **Other targets.** Use `target='active'` only when the user says to put it in *their* open workbook; `target='new'` for an explicitly separate file.
- **Parallel sessions.** Sessions run in parallel. Only writing in the same Excel window takes turns, and a second window opens when that one is busy.

Per question it builds:

- **`Q<n> Svar`**: the answer in TN report style (key points, P&L, checks, notes, source, chart when useful). All
  numbers are live formulas.
- **`Q<n> FinMax` / `Q<n> FinPlan` …**: one live Smart View grid per pull. The user can press Refresh, and the answer
  updates.
- **`Kilder`**: one row per pull (time, question, cube, POV, link).
- **Sheet order** is fixed automatically after every build/pull/answer/drill: all `Q<n> Svar` (Q1, Q2 … Q10), then
  other sheets, then data per question (FinMax, FinLite, FinPlan, `Linjer`), `Kilder` last, hidden sheets after.
  In the user's own workbook only these sheets are reordered among themselves. Do not reorder by hand.

**Call the `mcp__smartass__sv_*` tools directly** (server `smartass`, formerly `smartview`). That is how the user wants Smart View used: never a side script that
imports the package. Code changes reload by themselves on the next call. If `sv_analyse` / `sv_ws_pull` are missing
from the tool list, or `sv_status` lacks `behind_reload_proxy`, the server predates the reload relay. Ask the user to type
`/mcp` → smartass → Reconnect once, then continue with the tools.

## Pick the cube (decided with the user)

- **Actuals only** come from **FinMax**: the ledger, finest detail, FTE (`ac.STA20Y`). FinLite is a lighter actuals cube with no
  FTE and no Functions/Projects/Reserves/Systems.
- **Budget or forecast only** comes from **FinPlan**.
- **Actual vs budget/forecast** is done **entirely in FinPlan**, so line definitions match: Actual = version `Working`;
  Target and the latest `Forecast Q<q><yy>` = version `Final`. Add a reconciliation line to FinMax EBITDA. FinPlan groups
  accounts differently, for example Technology revenue YTD Sep 26 is 59.8 MNOK in FinMax and 67.2 MNOK in FinPlan, and
  EBITDA 2025 is 14.16 bn in FinPlan `BANORWAY` against 12.40 bn in FinMax `un.SY` (COGS alone differs by 2.3 bn).

## Steps

1. **Classify the request.**
   - A *lookup* (a member name, status, one number) stays in chat: `sv_find_member`, `sv_members`, `sv_query`.
   - Anything with a table of numbers is an *analysis*: continue with step 2.
2. **Scope it.** Do this for product, profitability, business-area and "what does X earn or cost" questions. Call
   `sv_scope(question)` FIRST. It searches the business map (forretningskart) and the concept library and returns:
   - matching **concepts**, with their blocks marked booked / estimated (with key) / attributed / not in FinMax;
   - matching account and product groups, with their booking pattern (product-tagged? reserve? which division?) and pitfalls.

   Then act on what it found:
   - **A concept matches:** answer with `sv_concept(concept)`. That is one call: every block, the keys and the FinMax basis. Report
     what was included, what was estimated (and the key used) and what is not in FinMax.
   - **Several family concepts match** (e.g. "mobilabonnement" = Consumer and Business subscriptions): run `sv_concept`
     once per family on the same workbook (`target=<workbook>`; each gets its own Q), and give each family plus their
     sum in the reply. Ask first only if the question clearly means one segment.
   - **No concept matches:** research the members, using the map hits plus `sv_find_member` / `sv_members` and the archive's
     revenue/COGS files. PROPOSE the blocks to the user (included / estimated with key / left out). Once they confirm, add the
     concept to `knowledge\concepts.json` and run `sv_concept`.
   - **Lookups** (one account, one entity) skip scoping.
3. **Resolve members** with `sv_find_member` / `sv_members` (atlas of all three cubes). Done when every member came from them.
   FinMax names are prefix + code: `un.SY.2TEC` (the entity shown as `2TEC Technology (un.SY)`), `ac.C30`, `pd.SY.P56`. Entities default to
   the `un.SY` tree (Telenor Norge AS); `un.NOR.*` is the management view with IFRS16 units. FinPlan uses aliases as
   shown (`2TEC Technology`, `FY26`, `Target 2025`).
4. **Build the analysis.** Use a concept or a template when one fits:
   - `sv_analyse(question, template, params)`. `target` defaults to the session's workbook.

     | Template | Params | Delivers |
     |---|---|---|
     | `division_pnl` | `entity`, `period?` (default: last closed), `basis` ytd/month | P&L vs the same period last year, checks, top-5 sub-unit variance drivers, EBITDA bridge (waterfall) |
     | `monthly_trend` | `entity`, `accounts?`, `through?` | Jan→month periodic, a YTD column checked against Accumulated, line chart |
     | `actual_vs_plan` | `entity` (FinPlan alias), `month?`, `year?` | Actual vs Target and the latest forecast, FinMax reconciliation |
     | `cost_base` | `parent`, `accounts?`, `period?`, `top?` | Top-N leaf cost centres + "other" + total; the full tree is on the data sheet |

   - Otherwise build it free-form in **one call**: `sv_ws_build(question, pulls=[{id, connection, rows, columns,
     pov}], spec)`. Pull cells in the spec use the pull `id`.
     1. Keep the spec short with shorthands:
        - `repeat` for one line per month or member: `{"over": "months:Jan..Dec", "line": {"key": "m{i}",
          "label": "{v}", "cells": [{"pull": "m", "row": [..], "col": ["{v} 26"]}, ...]}}`. Months not pulled yet
          stay empty.
        - `{"sum": "m*"}` and chart `"lines": "m*"`.
     2. Write `key_points` only for interpretation: why it moved, what to check. The server drafts the factual
        points from the numbers (totals vs comparison, largest increase and decrease) and puts them first.
        `auto_points: false` turns that off.
     3. Plan several pulls yourself (this year + last year, FinMax + FinPlan) and explain them afterwards.
     4. Add the same period last year to P&L and cost questions by default. Ask only when it is genuinely ambiguous, for example which forecast.
   - A follow-up ("legg til budsjett", "vis også i fjor") extends the same `q`.
     - For text or layout changes: `sv_ws_answer(q, spec_patch={...}, target=<workbook>)` sends only the change
       (about 4 s).
     - For new data: `sv_ws_build(..., q=<q>, target=<workbook>)`.
   - `sv_ws_pull` + `sv_ws_answer(q, spec)` still work for step-by-step building.
   - Every result carries `timings` (seconds per stage). Mention an unusually slow stage when you report.
   - Every result carries `formula_style` (skill `tn-formelstil`). The builder writes SUM over ranges,
     label-keyed XLOOKUPs and a variance % that is blank when the base is 0. If `formula_style.ok` is false, a
     `value`/`expr` you put in the spec broke a rule: fix it with `spec_patch` before reporting. LONG_FORMULA
     warnings on lookups with long FinMax labels are acceptable.
   - Editing an existing open workbook outside a Smart View answer (fix a formula, add a line, format a sheet,
     explain a number): use the `xl_*` tools on the same server (`xl_read`, `xl_write`, `xl_trace`, `xl_check`,
     `xl_format_sheet`, `xl_revert`). They are snapshotted and never save.
5. **Report in chat.** Done when everything below is written:
   - the returned key points;
   - a compact table of the headline lines in financial format;
   - cube · POV · period on one line;
   - every `checks` value that is not OK, and every warning;
   - the sheet names;
   - that the user's workbook was **not saved**. New analysis files (`target='new'` / `'session'`) are the MCP's own
     and ARE saved, in the analyser folder (`paths.analyser` in `%LOCALAPPDATA%\smartass\config.toml`).

## Drill-through (invoice / journal lines)

For "what is behind this number" (suppliers, invoices, accruals, allocations) on any FinLite cell, P&L or balance
sheet account alike (`un.SY` is a valid entity). FinLite only.
1. `sv_drill(question, cell={'All Accounts': .., 'All Entities': .., 'All Time Periods': ..})` (account and period
   required; without an entity the top member is drilled, with a warning). Parents and
   quarters are fine, and big cells (over 20 000 lines) split themselves into months, entities, accounts. It runs
   in a hidden browser: nothing opens on screen. NEVER use Smart View's own drill execute.
   - **One drill at a time across sessions.** A lock serialises them; a second session gets "try again": wait and retry.
   - **Time.** About 60–90 s per drill including the silent SSO login. The headless browser closes after 5 idle
     minutes, so the next drill pays the login again. Sign-in and page load are capped (about 95 s); the errors then
     name the next step (see the error table in [reference/operations.md](reference/operations.md)).
   - **Re-drill.** The same `q` and `id` refreshes the lines table in place. If the result says columns or grid
     changed, re-run `sv_ws_answer` for that q.
   - `target='none'` = cache only (no sheets in a workbook).
   - In the lines table, `Entity` is the cost centre (`un.SY.<cc>`), not the division; per-division totals are in
     `pieces` when the cell split. Use the column names the drill result returns in `where`.
2. Read `reconciliation` (must be `exact`, or explain the gap) and `profile`. If `no_supplier` is ~100 % (e.g. asset or
   accrual journals, `JE_SOURCE` Assets), say there are no suppliers and what the journals are instead. The profile has the top suppliers,
   accounts, months and journals without a supplier, plus `signs_mixed`.
3. Build the report with `sv_ws_answer(q, spec, target=<workbook>)` using `{drill: 'd', where: {...}}` cells and
   a reconciliation block that reads the live `cube_pull`. Pick the sections that answer the question. If
   `signs_mixed`, use no share column; show gross debit and credit instead.
4. Report `text_repairs.unfixed` (letters lost in EPM) and offer to add them to `%LOCALAPPDATA%\smartass\text_fixes.json`.

Column glossary, report rules, text repair, timings and a worked example: [reference/drill.md](reference/drill.md).

## Financial format (sheet and chat)

- **Tables:** full NOK `1 000 000`; negatives in parentheses `(1 000 000)`; zero `–`; FTE with 1 decimal; % with 1 decimal.
- **Running text** (key points, sentences): a minus sign, `−151 328 997`, never double parentheses.
- **Signs:** keep FinMax signs. Costs are negative. Variance = current − comparison, so positive = better and negative = worse
  for costs too. Variance % = variance / |comparison|.
- **Line labels** are in English as FinMax shows them (`C30 REVENUES`, `C70 Total Operating Expenses`).

## P&L chain (FinMax accounts) for checks

C30 + C40S = `C48`. C50S + C60 = `C70`. C48 + C70 = `EBITDA_BOI`. EBITDA_BOI + `C73` = **`EBITDA`**. `C75` is depreciation. FTE is
`STA20Y`, the end-of-period stock: `Accumulated` gives the level and `Periodic` gives the month's change. Search FTE as "man-years".

FinPlan: Revenue + COGS = Gross_Profit. + OPEX = EBITDA before Other Items. + Total Other Items = EBITDA. Target carries
no depreciation, so stop at EBITDA.

## Traps

- **History is restated to today's organisation.** The 1 June 2026 reorg emptied `2CON`, `2CWK` and `3WHS_C`, and moved
  `4REV_W` to `3WHS_W`. Old extracts match on leaves and on the `un.SY` total, not on division splits. `sv_atlas_changes` lists
  moved, added and removed members between atlas builds.
- **Use the last closed month.** `Current Reporting Month` resolves to it (Sep 26 on 13.10.2026). The month after it is open
  and carries reversals.
- **A `Yr` member before year-end is year-to-date.** The January Dec cut is preliminary.
- **Archive `GL` sheets are pasted values**, so rebuild them instead of refreshing.
- **The archive is protected.** Files in `Analyse (Telenor Norge AS) - Produktregnskapet\<year>\` are refused unless the user
  explicitly allows it for that request (`allow_archive=true`).
- **Wrong layout.** If Smart View returns a layout that differs from the request (moved rows), the pull fails loudly. Never
  work around that with fixed cell references.
- **Rejected members.** One member Smart View rejects (rc=1000) no longer sinks a pull: the engine isolates it, leaves it
  out and says so in `warnings`. Report those lines as missing.
- **Error cells.** Answers are scanned for `#NoAccess`, `#Invalid`, `#VALUE!` and the like, on the answer sheet and its
  data sheets. A warning such as `error cells on 'Q3 FinMax': 12 x #NoAccess (first: B7, ...)` or `formula error on line
  'x'` means those numbers are not usable: say so, never report them as zero. `sv_analyse` checks every template
  member against the atlas (and live for members the atlas lacks) before it pulls, and names a wrong one.
- **Input shapes.** `rows`/`columns` may be `[{dimension, members}]` or `{dimension: [members]}`.
- **Sign-in.** If a tool says the EPM login dialog is open or Smart View is not signed in, the USER has to click
  'Telenor SSO' in their Excel (you cannot): tell them, then retry once. Repeated login prompts usually mean another
  Office app (often PowerPoint) holds Smart View's WebView2: ask them to close it.
- **Member details:** `sv_member_info(connection, member)`.
- **Long calls** send progress notes (`still working: <stage> (n s)`) every 10 s: wait for them, do not call again.
- **FinPlan KPIs.** Read volumes, ARPU and percentages as `Periodic` per month; `Accumulated` sums the months for many
  of them (Mobile Subscriptions YTD Sep = 78 M against a stock of 8.7 M). Actual volumes are not in FinPlan for
  `BANORWAY`, only Target and forecast.

## Business map (forretningskart)

The repo's `knowledge\` folder (plus `packs\<pack>\knowledge\` for enabled packs) holds what the business MEANS in FinMax and FinPlan. Each area is an Excel review
sheet and a JSON copy that `sv_scope` reads:

| Area | File | Contents |
|---|---|---|
| 1. Accounts and bookings | `accounts_map.json` | 234 P&L groups: meaning, product/reserve tagging, division, pitfalls |
| 2. Products | `products_map.json` | 496 pd.SY groups: P&L, selling division, Produktregnskap codes and family |
| 3. Organisation | `org_map.json` | All un.SY entities: division 2026 vs 2025 (reorg), cost base, FTE, pitfalls |
| 4. FinPlan | `finplan_map.json` | 654 P&L, KPI and ratio lines: actual, Target, forecast, FinMax equivalent, reading rules |
| Concepts | `concepts.json` | `splitt` (hand-made, with keys) + 22 product families `fam_*` (direct contribution, booked only) |

- **Review state.** All four maps were approved by the user on 15.10.2026 (`review_status.json`). A rebuild flags only
  members added since then ("Ny siden godkjenning – ikke gjennomgått"); ask the user about those. Record a new approval with
  `python review.py approve <accounts|products|org|finplan> "<comment>"`.
- **Rebuilds.** Regenerate with `knowledge\build_*_map.py` after the monthly atlas refresh, then `build_family_concepts.py`
  (rewrites only `source: auto-familie` concepts). List new or moved members for the user.
- **Family concepts.** Each `fam_*` uses the highest pd.SY nodes whose groups all belong to that family (no double
  counting; 91 % of FY25 revenue covered) and checks revenue + COGS + opex = EBITDA before other items. Re-running
  `sv_concept` with `q=` replaces that question's data sheets.
- **Key facts from the map.**
  - Revenue is product-tagged; salaries (C50S) are never product-tagged.
  - Depreciation is about 30 % product-tagged, mostly in Technology.
  - Consumer fixed broadband/TV has no fixed Produktregnskap code (it is split with FinPlan keys).
  - The Functions dimension is unused in Telenor Norge.

Recipes for the team's recurring extracts (GL extract, revenue/COGS, depreciation control, FinLite special costs), each
reconciled against the 2025 archive, are in [reference/recipes.md](reference/recipes.md). Top levels of every dimension
are in [reference/cubes.md](reference/cubes.md). Errors, SSO, the watchdog, files and logs, and tests are in
[reference/operations.md](reference/operations.md).
