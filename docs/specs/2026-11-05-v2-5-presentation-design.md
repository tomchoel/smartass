# smartass v2 · Sub-spec 5: presentation (answer, data and Sources sheets, charts)

Status: design approved by Tom on 05.11.2026, with web research (IBCS, FAST, ICAEW, FT Visual Vocabulary,
Zelazny/SWD, Oracle Smart View, IPE guidance). Builds on `2026-11-04-v2-1-engine-design.md`.

## Why

A side-by-side render of a v1 answer (03.11) and a v2 answer (04.11) showed these problems:

1. **The order puts method before conclusion.** v2 answer sheets open with the agent's "Plan" and hypotheses, and
   the conclusion sits at the bottom.
2. **The data sheets lost their provenance.** `Q<n> Data` sheets have no provenance (v1 had source, POV, fetched,
   route), use the default blue table style, show `1,61E+08`, and cut entity names short.
3. **The Sources sheet is empty.** v2 never writes it, because `log_source` is v1 code tied to v1's `Pull`.
4. **The tables lost their structure.** Rows show no account codes, no highlighted P&L chain and no materiality
   marker. Column headers are generic ("Actual", "Comparison").
5. **Charts are rare.** Only trend and bridge blocks get one.

The styling code itself (`xl/styling.apply_report`, TN styles) is v1's and stays. What changes is structure,
headers, number formats and the parts that are missing.

**Principle: layout belongs to the engine, content to the agent.**
- The agent chooses blocks, metrics, periods, charts and cause sentences.
- The engine fixes the order, the formats and the provenance, so they cannot drift.

## Decisions (Tom, 05.11)

| # | Decision |
|---|---|
| D1 | Data sheets: a provenance header block on top, with table `Data_Q<n>` below it (formulas unchanged). |
| D2 | Answer sheet order is fixed: title → Key points → charts → main table → supporting tables → checks → notes/caveats → source. The plan and hypotheses are not drawn on the sheet. |
| D3 | Lines: a narrow code column, then the name. Subtotals are bold with a top border. The full P&L chain is shown when EBITDA is shown. |
| D4 | Default columns for P&L and cost: month and YTD side by side (AC · comparison · Δ · Δ %), plus a Material marker. |
| D5 | Key points: the engine writes the number sentence as a live formula, and the agent adds one cause sentence. |
| D6 | Checking is fast and cheap: an in-engine layout check in milliseconds. No image goldens and no long runs. |
| D7 | Every answer has at least one chart, and may have several. The agent picks the chart type and data that carry the message. |
| D8 | Column headers use scenario codes: `AC Oct 26`, `PY Oct 25`, `ΔPY`, `ΔPY %`, `BU`, `FC Q326`. Actual comes first, not IBCS PY→AC order. |
| D9 | Units: full NOK in tables. MNOK with 1 decimal in charts and Key points. |
| D10 | Variance columns are marked favourable or unfavourable by effect: green ▲ or red ▼. Other numbers stay black, with parentheses for negatives. |

## 1. Answer sheet (`Q<n> Answer`)

### 1.1 Fixed order

| Zone | Content | Source |
|---|---|---|
| Title | `spec.plan.purpose` or the question, as a message title where the agent gives one (`plan.title`, new and optional) | agent |
| Subtitle | `NOK · <cube> · <entity name> · <period label>` | engine |
| Key points | 3–5 points (§1.4) | engine numbers + agent cause |
| Charts | Every `chart` block (§3), in spec order, laid out in a grid of two charts per row | agent choice, engine draws |
| Main table | The first table-like block in spec order | agent |
| Supporting tables | The other blocks in spec order (bridge, trend, matrix, ranking, drivers, lines, concept) | agent |
| Checks | Sheet = model, reconciliations, bridge "Other / unexplained", lines vs cube; each shown as OK or DIFF | engine |
| Notes | The conventions (§1.5), the thresholds (variance % floor, materiality), findings and caveats | engine |
| Source | `Source: Oracle EPM Cloud (<cube>) · table Data_Q<n> (sheet Q<n> Data) · Sources row <n>` | engine |

- `plan` (purpose, depth, hypotheses) stays in the spec, the run log and answer memory. It is not drawn on the
  sheet.
- `CommentaryBlock.title` and its free `key_points` are dropped from the sheet. Free key points from the agent
  become extra Key points (§1.4), capped at 5 in total.
- The commentary drivers become the cause sentences.

### 1.2 Columns

- Column A is empty and narrow, as now.
- B is the **code**: width 9, grey, left-aligned. It holds the FinMax/FinPlan member without its prefix (`C30`,
  `C48`, `EBITDA`) for metric rows, and the bare entity or product code for entity and product rows.
- C is the **label**, with the width rule of plan 2D task 2.
- D onward hold the **numbers**, 16 wide.
- `LABEL_COL` becomes 3 and `FIRST_COL` becomes 4. Formulas are placed by `p.addr`, so the move is mechanical.
- Titles, sections and notes start in column B, so they stay left-aligned with the code.

### 1.3 Headers, column groups and the Material marker

**Headers.** Headers are built from the query periods and the scenario:
- value → `AC <period>`
- ly → `PY <period LY>`
- budget → `BU <period>`
- forecast → `FC <forecast label>`
- var → `ΔPY` / `ΔBU` / `ΔFC`, depending on the base on its left
- var_pct → `ΔPY %` and so on
- share → `Share %`

A YTD period shows as `YTD Oct 26`.

**Column groups.** `TableBlock.periods: list[str] | None` (new) gives one column group per period, for example
`["closed-month", "closed-ytd"]`. Each group repeats `columns`, and a blank narrow spacer column separates the
groups.
- The query must hold every period, and the guard names a missing one.
- The skill's default P&L and cost table uses `periods: ["closed-month", "closed-ytd"]`. It is a skill default,
  so the method stays soft.

**Material.** `TableBlock.material: bool = True` adds a last narrow column headed `Material`.
- It holds a **live formula**: `●` when |Δ| ≥ the metric's absolute threshold, or when |Δ %| ≥ its percentage
  threshold and |Δ| ≥ a tenth of the absolute threshold. This is the rule of `model/findings.material`, applied
  to the YTD group (the last group).
- The thresholds sit in labelled cells in the Notes zone, beside the variance % floors. They are never constants
  inside the formula (tn-formelstil).
- The model computes the same marker, so readback compares it.

### 1.4 Key points (D5)

1. The engine takes the material lines of the main table in the YTD group, ordered by |Δ| descending, up to 4.
   With no material lines, it takes the table's top line (EBITDA, or the first line).
2. For each point, the engine writes **one formula** that builds the number sentence from cells on the sheet:
   `="EBITDA YTD Oct 26 "&FIXED(D20/Unit_MNOK,1)&" MNOK, "&FIXED(F20/Unit_MNOK,1)&" MNOK vs PY ("&FIXED(G20*100,1)&" %)"`.
   - `FIXED` follows the user's locale. `TEXT` format codes do not, so `TEXT` is not used.
   - `Unit_MNOK` is a labelled input cell (1 000 000) in the Notes zone, not a constant in the formula.
   - The sign: Δ is current − comparison, and positive is better. The sentence says "vs PY" and keeps the sign
     with a minus sign.
3. The cause sentence goes on the row below, indented. It is the commentary driver whose `line` matches the
   point's line (or metric name), as `<text> (data: <evidence>)` or `<text> (hypothesis)`.
   - A material line without a driver gets `Cause not yet explained`. This is the critic's existing gap, made
     visible.
4. The agent's free `key_points` (static text) follow, up to 5 points in total.

The key points refresh with `fin_refresh`. The cause sentences are static until the agent patches them.

### 1.5 Number formats and conventions

| Cell | Format |
|---|---|
| NOK amount | `#,##0;(#,##0);–` |
| Δ NOK | `[Color10]▲ #,##0;[Red]▼ (#,##0);–` (D10) |
| Δ % | `[Color10]▲ 0.0 %;[Red]▼ (0.0 %);–`, and blank below the floor as now |
| FTE / stock | `#,##0.0;(#,##0.0);–` |
| Share % | `0.0 %` |

- Colour is used only in Δ columns. Black and white print still reads through ▲/▼.
- The conventions go in a fixed note at the top of the Notes zone: "NOK. FinMax signs: costs are negative. Δ =
  current − comparison; ▲ better, ▼ worse. Variance % blank when |base| < floor. ● = material (thresholds
  below)."
- Fonts, bold and borders follow `styling.apply_report`: no gridlines, bold only for titles, headers and
  subtotals, and a top border on sum and grand rows.
- Print setup: landscape, 1 page wide, header row repeated, footer `Source · <fetched date> · page n of m`.

### 1.6 P&L chain (D3)

When a table's metrics include EBITDA from FinMax, the engine completes the chain to C30, C40S, C48, C70,
EBITDA_BOI, C73 and EBITDA, in that order, adding any steps that are missing. In FinPlan the chain is Revenue,
COGS, Gross_Profit, OPEX, EBITDA before Other Items, Total Other Items and EBITDA.
- The kinds are: subtotals (`C48`, `EBITDA_BOI`) → `sum`, and `EBITDA` → `grand`.
- `TableBlock.chain: bool = True`. The agent turns it off only for a deliberate subset, and that is a departure
  noted on the sheet.

## 2. Data sheets (`Q<n> Data`, `Q<n> Lines <b>`)

### 2.1 Provenance header (D1)

Rows 1–N form a label/value block in a fixed order, one block per query on the sheet:

| Label | Value |
|---|---|
| Title | `Q<n> Data — <question>` |
| Question | `Q<n>`: the question as asked |
| Query | `<query id>`: cube, scenario and version (`FinPlan · Actual/Working, Target/Final`) |
| POV | Every dimension outside the grid, as names: entity name and code, currency, balance type |
| Periods | The concrete members (`Jan 26 – Oct 26`), and the relative spec (`closed-ytd`) |
| Route | `HTTP` or `Smart View ad hoc` |
| Fetched | A static date and time with timezone, and the Windows user |
| Rows | The row count of this query in the table, as a live `COUNTIFS` on the table |
| Control total | The static sum returned by EPM, beside a live `SUMIFS` of the table; status `OK` / `DIFF` |
| Sources | A hyperlink to this pull's row on the Sources sheet |
| Refresh | "Refresh: fin_refresh, or ask a follow-up. Do not edit this table by hand." |

Below the block there is one empty row, then the table.

### 2.2 The table

- It keeps its long format, its name `Data_Q<n>` and its columns. Formulas do not change.
- **Style**: a light style (`TableStyleLight1`, no banding) with a bold header and a bottom border.
- Value uses `#,##0;(#,##0);–` in **blue** font (FAST convention for imported numbers) and is right-aligned.
- Label columns are text and fitted to their content (entity names in full), with a maximum width of 40.
- Freeze panes sit below the table header. The table filters stay on.
- `render/excel.rewrite_grid` and `read_data` work from the table's own top-left cell (`lo.Range.Row`), never
  from A1.
- **Old answers** (table at A1): a redraw inserts the header rows above the table. Inserting whole rows keeps
  the table and its structured references.

### 2.3 Lines sheets

Lines sheets get the same header. The Query row holds the drill cell, and the Control total row holds the lines
total against the cube cell, with the tie-out status already computed by the lines block.

## 3. Charts (D7)

### 3.1 Chart block

```
ChartBlock {type "chart", kind, title, source, series?, categories?, highlight?, sort?}
  kind:   waterfall | line | column | bar | stacked | stacked100 | combo | variance
  title:  the message in one sentence, with the number ("Revenue explains 115 % of the EBITDA gap to Target")
  source: {block: <index or id>}  - draw from a table already on the sheet
        | {query, metrics, periods, rows?}  - own data: the engine draws a labelled "Chart data" table under
          Notes on the answer sheet, with live formulas, and the chart reads it
  series: column codes of the source (value, ly, budget, forecast, var); default by kind
  highlight: line labels or series to colour; everything else is grey
  sort:   desc | asc | none (bar and ranking default desc, with Other last)
```

- The existing `chart: bool` on trend and bridge blocks becomes sugar for a chart block on that block.
- **Guard (hard):** an answer without a chart block is refused with one fix, "add a chart block (e.g. kind
  waterfall on the bridge)". The single exception is `depth: quick`, where one chart is added automatically from
  the main table: a column chart of AC against the comparison.

### 3.2 Choosing the chart (skill text, the agent's choice)

| Message | Kind |
|---|---|
| Bridge from comparison to actual (EBITDA, cost) | `waterfall` |
| Development over time (6 months or more) | `line`, 1–2 series highlighted |
| Few periods, or AC against PY/BU | `column`, at most 3 series |
| Ranking (cost centres, suppliers, products) | `bar`, sorted, top N plus Other, at most 15 bars |
| Variance by line | `variance`, diverging bars from zero, green or red by effect |
| Mix or share | `stacked100`, at most 5 parts, never a pie |
| Two units (MNOK and %) | `combo`, line and column |

### 3.3 Drawing (IBCS styling, built for COM)

- **Scenarios** (IBCS):
  - AC: solid dark navy.
  - PY: solid grey.
  - BU: outlined (white fill, navy border).
  - FC: hatched (`msoPatternWideUpwardDiagonal`).
  - Variances: green or red by effect.
  - Highlight: the TN accent. The rest is grey.
- **Labels instead of axes:**
  - Data labels in MNOK use `#,##0.0,,`, the unit sits once in the subtitle, and there are no gridlines or tick
    marks.
  - There is no legend for 3 or fewer series, which are labelled directly. With more, the legend goes on top.
- **Waterfall:** a stacked column of base (no fill), increase, decrease and total series, computed in the Chart
  data table. Excel's native waterfall is not used, because its COM object model does not reliably set totals.
  A bridge that crosses zero splits the base into a positive and a negative part.
- **Variance:** two series, favourable and unfavourable, with 100 % overlap. `InvertIfNegative` is not used.
- **Bar:** `ReversePlotOrder` on the category axis keeps the largest bar at the top.
- **Order of calls:** set the series and type first, then the formats, because changing the type resets the
  formatting.
- **Size and place:** 480 × 260 pt, in the Charts zone (two per row), named `smartass chart` (owned, so a
  follow-up can drop it).
- A chart that fails is a **warning** plus a one-line fix, never a failed answer. The guard still requires at
  least one chart to be drawn.

## 4. Sources sheet

- The sheet holds one row per query pull. It is **append-only**, so a refresh adds a row.
- Columns:
  - Fetched (static)
  - Q
  - Query id
  - Question
  - Cube
  - Scenario/Version
  - POV
  - Periods
  - Route
  - Rows
  - Control total
  - Status (OK/DIFF)
  - Data sheet (hyperlink)
  - Supersedes (the row it replaces, for refreshes)
- A **Known limits** block sits above the log. It lists the workbook's distinct `warn` and `block` caveats, one
  line each, rewritten on every answer.
- The sheet is the same TN style as the data sheets: header, frozen panes, filters, Sources last
  (`sheetorder`).
- It is written by a new v2 `xl/sources.py`, called from `pipeline._draw`, `redraw_job` and `fin_refresh`. The v1
  `log_source` (bound to v1's `Pull`) is deleted.

## 5. The layout check (D6)

After every draw, `render/check.py` runs on the `SheetPlan` and a few COM reads (under 100 ms). Each finding is
one line with its fix, returned in the reply as `layout` (a list):

1. The zones are in §1.1 order, and there is no Plan zone.
2. At least one chart was drawn.
3. Every numeric cell has a number format from §1.5, and no cell shows `E+`.
4. No text is cut off (the existing lint) and no `#` shows (column too narrow).
5. The data sheet has its header and its control total says OK.
6. The Sources rows of this Q are present.

A finding is a `warning`, not a stop. The agent fixes it by patch. Unit tests lock the plan side.

## 6. Spec, skill and pane changes

- **Spec:**
  - `plan.title` is new and optional.
  - `TableBlock.periods`, `.material` and `.chain` are new.
  - `ChartBlock` is new.
  - `CommentaryBlock.key_points` is capped at 5.
  - The trend and bridge `chart: bool` stays as sugar.
- **Reply:** `layout` findings. `key_values` keys keep the **model line labels and the column codes' display
  headers**. Headers change (`Actual` → `AC Oct 26`), so the test-set goldens change keys (§7).
- **Skill:**
  - The order and the "plan is not on the sheet" rule.
  - The chart table in §3.2, with one worked example per chart kind in `reference/charts.md`.
  - The default P&L and cost table with month and YTD.
  - SKILL.md stays under 5500 characters, with detail in reference files.
- **Pane prompt:** "every answer has a chart; choose it by the message".

## 7. Test and rollout

- **Unit tests**, all offline on fakes:
  - zone order
  - headers by scenario and period
  - column groups and the spacer
  - the Material formula and its thresholds
  - the Key point formula (`FIXED`, `Unit_MNOK`), and driver matching
  - P&L chain completion
  - number formats
  - the data header fields
  - `rewrite_grid` with the table below a header, and inserting the header above an old A1 table
  - the Sources rows (append, Supersedes)
  - the chart block guard and the source variants
  - the stacked-waterfall series maths, including a bridge that crosses zero
  - the layout check findings
- **Test set:**
  - Run `testset candidate --offline` for all 9 cases.
  - Allowed differences: renamed headers, added chain lines, Material and key point lines.
  - The **values of existing lines must be identical**. Any changed value means STOP and report.
  - Tom approves the new goldens. The approve step lists the header renames.
- **Live (controller):** one answer each of P&L month and YTD vs PY (FinMax), actual vs Target (FinPlan), a
  bridge, a ranking and a lines block. Render each to PDF once and look at it (a one-off controller check, not a
  test).
- **Rollout:** bump the version. New answers, patches and `fin_refresh` draw the new layout. Old workbooks change
  only when they are redrawn.

## Out of scope

- Pivoted v1-style data grids (D1 chose the header and long table).
- NOK thousand or million units in tables (D9).
- Image goldens (D6).
- PowerPoint export.
- IBCS PY→AC column order (D8).
