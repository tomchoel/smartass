# Design: smartass as a finance-expert tool

## Context

smartass (`C:\dev\smartass`, 30 MCP tools) works well as a **data pipe**: it drives Oracle Smart View over COM in a
hidden Excel, has a member atlas (SQLite, ~300k members), a spec-driven Excel answer builder, a journal-line drill and
the `xl_*` tools for open workbooks. Two things hold it back from being a proper finance tool for colleagues:

1. **Fragile and slow data access.** Every question is a live Smart View pull: 25–50 s, plus COM/SSO/dialog/WebView2
   failure modes (`docs\troubleshooting.md:104-121`, reliability audit 2026-10-23). There is no REST usage
   (`DESIGN.md:7`).
2. **Thin financial reasoning.** It lives in 4 hardcoded templates (`engine\smartview_mcp\templates.py`, ~300 LOC),
   JSON maps and skill prose. The LLM writes the whole pull+spec JSON itself and does all interpretation. Missing:
   materiality, driver bridges (FTE × cost/FTE, price/volume/mix), LTM/run-rate, year-end landing, KPIs,
   anomalies, structured commentary, review.

Decisions from the grilling session (Tom, 2026-10-25):
- **Users:** finance colleagues too → robust and self-checking product.
- **Output:** Excel analysis (traceable to the cube) **+ written memo** (CFO style).
- **Use cases:** everything a large telecom finance division needs; start with monthly P&L + variance, cost deep dives,
  product profitability, forecast & scenarios.
- **Interface:** a chat pane inside Excel (add-in).
- **Trust:** every number traceable to a cube cell or a workbook formula; the LLM never does arithmetic; automatic
  reconciliation before anything is shown.
- **EPM REST access:** unknown → find out first, design for both paths.
- **AI backend for colleagues:** unknown → engine stays model-neutral (MCP + HTTP); decide before the add-in stage.
- **Data:** local warehouse first, shared later. Closed months snapshotted, open month/forecast refreshed on demand,
  freshness shown on every answer.
- **Approach:** evolve into layers. Keep atlas, drill, Excel builder and `xl_*`; build the new layers in independent
  stages.

## Target architecture

```
 Excel add-in chat  ·  Claude Code / Copilot          (front ends)
          │                    │
   Finance-expert layer: router skill + playbooks + reviewer agent      (Stage 5)
          │  plan → run → memo
   High-level tools: fin_plan · fin_run · fin_memo  (+ existing sv_*/xl_* for power use)   (Stage 4)
          │
   Calc engine (deterministic, Python/DuckDB) + reconciliation          (Stage 3)
          │
   Semantic catalogue (metrics, scenarios, entities, time, drivers, materiality)   (Stage 1)
          │
   Warehouse (DuckDB, long fact table, freshness)                       (Stage 2)
          │
   Sources:  RestSource (exportdataslice, if Stage 0 = go)  |  SmartViewSource (existing COM path, fallback + live sheets)
```

## Stages

**Stage 0: EPM REST access probe** (spike; runs in parallel with Stage 1)
- Ask the EPM/IDCS admin for an OAuth client (device-code flow) for read-only access. Basic auth with SSO users
  will fail.
- Probe script `engine\smartview_mcp\epm_rest.py` (probe only): list applications, read substitution variables, one
  `exportdataslice` on FinMax and FinPlan.
- **Test the key risk:** Oracle documents that exportdataslice only returns cells the user can read *and write*.
  Compare against `sv_query` for the same grid.
- Output: go / no-go note in `docs\specs`. No-go means SmartViewSource stays the only source; nothing else blocks.

**Stage 1: Semantic catalogue**
- `knowledge\catalog\*.yaml`:
  - **metrics:** account member, sign, whether higher is better, P&L chain parent (C30+C40S=C48, …).
  - **scenarios/versions:** Actual = FinMax; Target and the latest `Forecast Q<q><yy>` = FinPlan Final; FinPlan
    Working as actual.
  - **entities:** business areas and aliases (e.g. "Wholesale" → `un.SY.2WHS`), plus the 1 June 2026 reorg.
  - **time:** last closed month, YTD/LTM/LY.
  - **KPIs/drivers:** FTE → cost per FTE, volume × ARPU from FinPlan.
  - **materiality:** default ±10 % and NOK floor, an always-comment list (revenue, personnel, EBITDA).
- Generated parts reuse the atlas (`atlas.py`) and the business-map builders (`knowledge\build_*.py`). Handwritten
  parts are reviewed by Tom.
- A validator (unit tests) checks that every member exists in the atlas. The hardcoded lists in `templates.py:195-202,405`
  move here.

**Stage 2: Warehouse + source abstraction**
- `engine\smartview_mcp\warehouse.py`: DuckDB at `%LOCALAPPDATA%\smartass\warehouse.duckdb`. One long table
  `fact(cube, scenario, version, entity, account, period, view, currency, value, source, fetched_at)`.
- `sources.py`:
  - `SmartViewSource` wraps the existing retrieve + bisect path (`ops.py`, `grid.py`, `host.py`, `modBridge.bas`).
  - `RestSource` is added if Stage 0 = go.
- Extract jobs:
  - **Close snapshot:** business-area entities × catalogue accounts × 36 months, FinMax + FinPlan Target/Forecast.
  - **On-demand top-up** of the open month and the latest forecast.
  - **Restatement invalidation** when `sv_atlas_changes` reports moved members.
- Every result carries a freshness stamp. The path is configurable so a shared copy (SharePoint/server) can come later.
  The shared copy itself is out of scope now.

**Stage 3: Calc engine + reconciliation**
- New package `engine\smartview_mcp\analytics\`, pure functions over warehouse tables. It covers:
  - variance vs LY, Target and Forecast;
  - materiality filtering and ranking;
  - bridges by account, by entity, and by driver (FTE × cost/FTE; price/volume/mix where volumes exist);
  - YTD, LTM, run rate and seasonality;
  - year-end landing (YTD actual + remaining forecast, with a run-rate alternative);
  - KPIs;
  - cost-centre anomalies (z-score).
- `reconcile.py` checks:
  - children = parent;
  - the P&L chain identities (moved from `templates.py`);
  - months sum to YTD;
  - the tie to FinMax EBITDA.
- Output is result tables plus a **facts list**: id, value, unit, formula, source cells.
- Merge the three number formatters (`tools_ws.py fin()`, `specs.py _fmt/_pct`, `templates.py _fmt/_pct/_dec`) into one
  `numfmt.py`.
- Rebuild `sv_analyse` templates on the engine, then retire the old code.

**Stage 4: Plan → run → memo tools**
- `fin_plan(question)`: resolves through the catalogue and returns an editable **analysis plan**: playbook, scope,
  lines, comparisons, KPIs, materiality, drill dimensions, data needed and its freshness.
- `fin_run(plan)`:
  - fetches (warehouse first, then a top-up), computes and reconciles;
  - renders the workbook through the existing builder (`answer.py`, `specs.py`, `styling.py`, `vendor\tn_rapportstil.py`,
    `vendor\formelsjekk.py`), with the data sheet plus formula-driven answer sheets;
  - returns the facts with cell addresses and the check results.
- `fin_memo(run_id, memo)`:
  - **every number in the memo must match a fact**, otherwise it is rejected;
  - writes a `Memo` sheet + .docx/markdown.

  This is what enforces "all numbers traceable".
- The existing `sv_*` and `xl_*` tools stay as power tools. `sv_ws_build` remains for ad hoc work.

**Stage 5: Finance-expert layer**
- `skills\smartass\SKILL.md` becomes a short router and rules file. New playbooks in `skills\smartass\playbooks\`:
  - monthly P&L + variance
  - cost/opex/FTE deep dive
  - product profitability (uses `sv_concept`)
  - forecast landing & scenarios
  - later: capex, balance sheet, KPI packs
- Each playbook states:
  - which questions it answers and the data it needs;
  - its steps and materiality;
  - its commentary structure: *what / why / what next*, "a driver explains why, not what", and
    "driver unclear → flag for controller";
  - its checklist.
- The structure is borrowed from `anthropics/financial-services-plugins` (variance-commentary, audit-xls) and the
  installed `finance:variance-analysis`.
- Telecom KPI glossary: ARPU, churn, EBITDA margin, capex/sales, cost per FTE, opex per subscriber.
- `agents\fin-reviewer.md`: a reviewer subagent that checks the memo against the facts and the playbook checklist
  before delivery.

**Stage 6: Excel add-in**
- An Office.js task pane, sideloaded per user, talking to an extended `httpapi.py` (`/plan`, `/run`, `/memo`,
  `/status`, progress streaming).
- **Decision point at stage start:** which model drives the chat (M365 Copilot, GitHub Copilot or Claude
  Enterprise/API). The engine needs no change for any of them.
- First version: question → shows the plan for approval → run with progress → workbook + memo.

**Order:** 0 ∥ 1 → 2 ∥ 3 (3 is developed against SmartViewSource/fixtures) → 4 → 5 → 6. Each stage gets its own
spec → implementation plan → build cycle.

## First steps

1. Draft the email/request to the EPM admin for the OAuth client (Stage 0). Tom sends it.
2. Run the writing-plans skill for **Stage 1 (catalogue) + the Stage 0 probe script**.

## Verification (per stage, and end to end)

- `python -m pytest -q` in `engine\`: the existing 295 tests stay green, plus new unit tests per module (catalogue
  validator, warehouse, analytics, reconcile, memo validator).
- Golden tests on a **synthetic** fixture extract (no real data in the repo; `scripts\check_staged.py` guards this) →
  expected variance tables, bridges and landing.
- Live matrix (`engine\tests\live\`) gets a new scenario, "P&L Wholesale Sep 26":
  - `fin_plan` → `fin_run` → `fin_memo`;
  - totals checked against `sv_query`;
  - a memo with an invented number must be rejected;
  - warm-warehouse answer time < 10 s.
- Stage 0: the exportdataslice grid equals `sv_query` cell for cell, or the gap is documented.

## Risks / open questions

- REST access and exportdataslice read-only gaps (Stage 0 settles this).
- Policy for keeping cube data on local disk (sensitivity label "Internal") → confirm with IT before Stage 2 ships to
  colleagues.
- Which model the add-in uses (Stage 6 decision).

## Sources (research 2026-10-25)

- Planning REST Export Data Slice: https://docs.oracle.com/en/cloud/saas/epm-cloud/prest/export_dataslices.html
- Data slices overview: https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/prest/working_with_dataslices.html
- Export Data job: https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/prest/export_data.html
- REST authentication (Basic / OAuth 2, device code): https://docs.oracle.com/en/cloud/saas/epm-cloud/prest/authentication.html
- OAuth for EPM Cloud: https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/cgsad/3_info_security_epm_cloud_oauth.html
- EPM AI features 26.04: https://docs.oracle.com/en/cloud/saas/readiness/epm/2026/epm-apr26/26apr-epm-wn-f44078.htm
- Essbase MCP server (standalone Essbase only, not EPM Cloud; safety design worth copying): https://blogs.oracle.com/autonomous-ai-database/essbase-mcp-server
- Community EPM MCP server (reference only): https://github.com/fmepm/oracle-epm-mcp-server
- Anthropic finance agents and skills (variance-commentary, audit-xls): https://github.com/anthropics/financial-services-plugins
- Semantic layer vs text-to-SQL accuracy: https://docs.getdbt.com/blog/semantic-layer-vs-text-to-sql-2026
- Pigment Analyst Agent (materiality ±10 %, supervisor agent): https://www.pigment.com/blog/analyst-agent-launch
- Workday Adaptive Planning Agent (top variances, drill, narrative): https://doc.workday.com/adaptive-planning/en-us/what-s-new/releases/2025r2-release-notes/planning-agent--data-exploration.html

Not verified: exportdataslice size and rate limits, whether read-only users get data back, Telenor's OAuth setup.
