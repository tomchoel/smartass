# Design: Stage 2+3 - warehouse and calc engine (first slice: business-area P&L)

Part of `2026-10-25-finance-expert-design.md`. It builds on:
- the Smart View HTTP route (`svhttp`, about 0.2 s per grid once signed in);
- the semantic catalogue (`knowledge\catalog.toml`).

## Decisions (Tom, 2026-10-25)

- **Warehouse content:** all business areas plus the Telenor Norge total, × the catalogue's P&L lines (incl. FTE),
  × 36 months. Actuals from FinMax; actual, budget and the latest forecast from FinPlan. Lower levels (cost centres,
  accounts) are pulled on demand and stored as they come.
- **Refresh:** automatic when a new month closes, and when the atlas shows a reorg or restatement. The open month and
  the latest forecast are read fresh at question time, and every answer carries a freshness stamp.
- **Database:** DuckDB (a new dependency).
- **First slice:** "P&L for <business area>":
  - month and YTD, compared with last year, budget and the latest forecast;
  - materiality flags (catalogue rule);
  - an EBITDA bridge per line and per sub-unit;
  - KPIs: margins, cost per FTE;
  - year-end landing.

  Everything is delivered as formulas in Excel.

## Principle

**The engine thinks fast; the workbook proves it.**

The warehouse and the calc engine decide *what* the answer needs: which lines are material, which sub-units drive the
bridge, the KPIs and the landing. They decide in milliseconds, without Excel.

The delivered workbook is still built by the existing pipeline (`sv_ws_build`):
- live Smart View source sheets ("kilder") that anyone can refresh;
- answer cells that are formulas over them;
- a reconciliation check per value, warehouse = source sheet.

The LLM never computes a figure. Every number in the memo (Stage 4) comes from a fact that points at a cell.

## Components

### 1. `warehouse.py`: storage (DuckDB)

- **File:** `%LOCALAPPDATA%\smartass\warehouse.duckdb`. The path is configurable (`paths.warehouse`), so a shared copy
  can come later. One writer per machine, with a file lock like the drill's.
- **Table `fact`:** `cube, scenario, version, entity, account, period, balance, value, source ('http'|'com'),
  fetched_at, closed_month, atlas_build`. The primary key is the member tuple; an upsert replaces the row.
- **Table `fill`:** one row per extract job (what, when, cells, seconds, result), for freshness and the doctor.
- **API:** `put(rows)`, `get(cube, entities, accounts, periods, scenario, version, balance) -> {tuple: value}`,
  `freshness(...)`, `status()`. No business logic here.

### 2. `extract.py`: fill and refresh jobs (over `svhttp`, COM fallback)

The **base fill** is built from the catalogue:
- **FinMax Actual:**
  - flow lines read *Periodic* (YTD is the sum of the months);
  - stock lines (`measure = "stock"`, i.e. FTE) read *Accumulated* (the month-end level), plus the opening December;
  - each grid is entities × lines on rows and months on columns.
- **FinPlan:**
  - Actual/Working, Target/Final and the latest `Forecast Q<q><yy>`/Final;
  - Period months for FY-1, FY and FY+1;
  - FinPlan entity names come from the catalogue.

That is about 5,000 + 1,500 cells: a few dozen grids, under a minute.

**Closed month:** `Current Reporting Month`.
- First try it over HTTP. That means probing the provider member-information request and adding it to the svxml
  allowlist (read-only).
- Otherwise use the COM route, cached per day as today.

**Refresh triggers**, checked at the start of every `fin_*` call. The check is cheap and makes no EPM call when
nothing changed:
- **A new closed month:** fill the new month, and re-read the previous two (late postings).
- **The atlas shows moved members for a catalogue area** (`atlas_history` diff, i.e. a restatement): re-read that
  area's full 36 months.
- **The latest forecast name changes:** re-read FinPlan forecast.
- **The open month or the latest forecast:** read live for the question and stored with `closed_month` = open. It is
  never mixed into closed history.

**Reconciliation after each fill** (results go into `fill`):
- the catalogue identity checks per entity and month (C30 + C40S = C48, ...);
- the sum of Periodic months = Accumulated for one spot column per entity;
- areas sum to the company total where the catalogue says they should, as information only (eliminations exist).

### 3. `analytics/`: deterministic calc engine (pure functions, no I/O)

The input is a `Frame`: values keyed by (entity, line, period, scenario) with the catalogue attached.
- `periods.py`: month, YTD, LTM, same period last year, and the remaining months of the year (the closed month as
  reference).
- `measures.py`: a flow sums over months, a stock takes the period-end level, and the stock average is the opening
  level + each month end (n+1 points).
- `variance.py`: current − comparison, and % of |comparison|. The sign convention is in the catalogue: a higher value
  is better.
- `materiality.py`: the catalogue rule (10 % from NOK 500k OR NOK 5m, `always` lines). It returns flags and a ranking
  by |variance|.
- `bridge.py`: an EBITDA walk by line (FinMax chain) and by sub-unit (children from the atlas). The top-N sub-units
  plus "other" sum exactly to the total.
- `kpis.py`: the catalogue KPIs, including `average_stock` and annualise.
- `landing.py`: year-end = YTD actual + remaining months of the latest forecast. The run-rate alternative is
  YTD / n × 12, adjusted by last year's seasonality profile. Both are shown with their gap to target.
- `facts.py`: every number the engine produces becomes a `Fact(id, value, unit, definition, inputs)`. `inputs` names
  the warehouse tuples. This makes every number traceable.

### 4. First slice tool: `fin_pnl(area, period=None, compare=None)`

1. Resolve `area` with `catalog.resolve_entity`. For an abolished area, explain and offer its successors. For several
   areas, ask.
2. Run the refresh check, then pull from the warehouse.
3. Run the calc engine: variances against LY, budget and forecast; materiality; line bridge and sub-unit bridge; KPIs;
   landing.
4. Generate a `sv_ws_build` request:
   - **pulls:** the live Smart View grids that cover exactly the cells the answer shows (the source sheets);
   - **spec:** sections for the P&L table, KPIs, bridge and landing; `checks` = catalogue identities plus one
     reconciliation check per source sheet (warehouse value = sheet value, within NOK 1); a waterfall chart; the
     materiality flag as a column;
   - **key_points:** drafted from the facts (factual only); interpretation is Stage 5.
5. Return a summary for the agent to write the answer from: facts, flags, the workbook path and freshness.

`templates.division_pnl` becomes this slice's predecessor. Retire it once `fin_pnl` covers it.

### 5. Ops

- **`smartass doctor`:** a `warehouse` check (file present, last fill, freshness, reconciliation result).
- **New tool `sv_warehouse(action='status'|'refresh'|'rebuild', area=None)`** for manual control.
- **Live matrix:** a scenario "fin_pnl Wholesale": every answer value equals a direct `sv_query`, and the
  reconciliation checks are OK.

## Error handling

- **Warehouse missing or locked:** the slice answers from live pulls (slower) and says so.
- **Any fill error:** recorded in `fill`; stale data is used with a visible freshness warning; never a silent gap.
- **Calc engine:**
  - a missing input is an explicit `Fact` with value None and a reason;
  - a zero comparison value has no % (materiality uses the NOK rule);
  - no division by zero.

## Testing

- **Unit tests** for every analytics function on synthetic frames:
  - the n+1 average;
  - annualising;
  - bridge closure;
  - landing;
  - materiality with the floor;
  - sign conventions.
- **Warehouse** tests on a temporary DuckDB.
- **Extract** tests with the svhttp fake transport.
- **Live:** the base fill on the real cubes, reconciliation all OK, and `fin_pnl` for Wholesale matches `sv_query`.
- **Never** commit real figures (`scripts\check_staged.py`).

## Out of scope (later)

- The shared warehouse.
- Balance sheet and capex.
- Price/volume/mix bridges (they need volume KPIs from FinPlan).
- The memo validator and the plan → run → memo tools in full (Stage 4).
- Playbooks (Stage 5).
- The add-in (Stage 6).

## Stage 2+3 live check (Task 22, 2026-10-25)

`engine\tests\live\warehouse_live.py`, run against the real cubes. No figures are recorded.

| Check | Result |
|---|---|
| Closed month over HTTP (`EnumMembers` children of `CurrRepMonth`) | PASS, Sep 26 |
| Base fill of a scratch warehouse | PASS: 12 grids, 8,955 cells, 7.8 s, route http, 3,654 reconciliation checks, 0 problems, 0 errors |
| Warm refresh check | PASS: 0 jobs, 0.06 s |
| Live top-up (forecast, open months, sub-units) | PASS: 41 cells, 1.2 s |
| Engine = direct retrieve, Wholesale | PASS 7/7: EBITDA YTD, EBITDA YTD LY, revenue for the month, FTE month end, FinPlan actual / Target / Forecast EBITDA YTD |
| Bridges | PASS: they close (10 sub-units), with no unexplained residual |
| Calc engine | 469 ms |

The missing facts left after the fill are all percentages on a zero comparison (Other items with no budget). They are
correct by design.

**What the first live run caught.** `grid.parse` reported real FinPlan members as unknown (`2WHS Wholesale`,
`Period: Oct`, ...), because FinPlan names equal their labels and some rows or columns were empty. The extract
dropped those cells, so the warehouse had no FinPlan data for 5 areas and the FY27 budget (not yet entered) failed.
Since `2730879`, flagged members are checked against the atlas, and only members the atlas does not have are
dropped.

## Matrix scenario S11 (Task 23, 2026-10-25)

`S11.fin_pnl.wholesale`: `fin_pnl("Wholesale")` builds a workbook in the user's Excel. Report:
`docs\audit\2026-10-25-stage2-3-s11.md`.

- **All assertions pass.** 48 checks are OK, including the 5 "Warehouse = source sheet" checks; 4/4 values equal
  `sv_query`; the warm call only does the live top-up.
- **Verdict SLOW:** 273 s against a 240 s budget. The cold call took 193 s and the warm call 80 s, and a Smart View
  Login dialog appeared once in Excel.
- **Where the time goes:** in the existing Excel pipeline, not in the new engine.

| Step | Cold | Warm |
|---|---:|---:|
| Warehouse refresh | 7.8 s | - |
| Live top-up | 1.3 s | 1.2 s |
| Closed month (incl. browser sign-in) | 10.2 s | cached |
| New workbook | 68.8 s (incl. the Login dialog) | - |
| Smart View pulls, 5 sheets | 67.4 s (connect 31 s) | 29.1 s |
| Answer sheet | 26.2 s (styling 13.2 s) | 32.0 s (styling 16.6 s) |

**Follow-up:** speed up the Excel build (connect, the check-colour styling pass, cell writes). This is a separate
change from Stage 2+3.
