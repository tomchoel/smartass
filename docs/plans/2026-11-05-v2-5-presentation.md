# smartass v2 · Plan 5: presentation (answer, data, Sources, charts)

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans. Steps use checkbox (`- [ ]`) syntax.

**Goal:** implement sub-spec 5 (`docs/specs/2026-11-05-v2-5-presentation-design.md`, decisions D1–D10). Answer
sheets get:
- a fixed order: Key points, then charts, then tables, then checks, notes and source;
- scenario headers, month and YTD column groups, the Material marker, the P&L chain, code columns, and Δ
  formatting.

Data sheets get a provenance header. The Sources sheet is written again. Every answer gets at least one chart.

**Architecture:**
- **Model (`model/`):** decides what the columns, lines and parameters are. This covers column groups,
  headers, the chain, the Material expression, and the chart blocks with their data.
- **Layout (`render/layout.py`):** decides where each part sits. It produces the zones in a fixed order.
- **Writers (`render/excel.py`, `render/charts.py`, `render/datasheet.py`, `xl/sources.py`):** draw the sheets.
- **Check (`render/check.py`):** validates the layout in milliseconds.

The formulas on `Data_Q<n>` and `Sheet = model` stay as they are.

**Tech stack:** Python 3.12, pydantic 2, Excel COM via `xl/jobs.com`, styling `xl/styling.apply_report` (TN
styles), pytest + ruff, test set CLI.

**Conventions** (as in plans 2B–2D):
- Branch `v2/5-presentation` (it exists, with the spec commit). TDD.
- Gates after every task: `uv run pytest -q` and `uv run ruff check .` in `engine/`, and
  `uv run python -m smartass.testset run` = 9/9. key_values keys stay stable (Task 1), so the test set stays a hard
  gate throughout.
- Never run python through a stdin heredoc: write a `.py` file and run it with `< /dev/null`.
- Stage files explicitly.
- Commits end with
  `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>` and
  `Claude-Session: https://claude.ai/code/session_01MW9ANq493fHqddMk5ipVTq`.
- Tests that touch Excel use the existing fakes (see `tests/v2/test_v2_excel*.py` and `tests/v2/fakes*`). No real
  Excel in unit tests.

**Key facts about the code** (read before starting):
- `model/blocks/common.py`:
  - `HEAD` and `col_headers()` build headers today ("Actual", "Last year", "Variance").
  - `Line(key, label, cells, kind, level, metric, entity)` describes one line.
  - `BlockOut(key, type, title, headers, cols, query, col_metrics, lines, chart, drills)` describes one block.
- `Ctx.pick(q, metric, series, entity, base_months, product)` builds the SUMIFS. `base_months` already selects a
  subset of months, so it is the hook for column groups.
- `model/series.months()` shifts `ly` by 12 months.
- `model/findings.material()` is the materiality rule. `Resolver.materiality(metric)` returns `(abs, pct)`, and
  `Resolver.pct_floor(metric)` returns the variance % floor.
- `model/analyze.apply_floors` puts `Param(FLOOR_KEY + metric, value)` into Ratio cells. `Result.floors` carries
  them to the layout's "Thresholds" cells (`p.params`). The Material thresholds and `Unit_MNOK` reuse this
  mechanism.
- `render/layout.layout()`:
  - currently draws Title, Subtitle, Plan, the blocks, Thresholds, Findings, Caveats and Source;
  - `LABEL_COL = 2`, `FIRST_COL = 3`;
  - `p.rows[{row, kind, num}]` feeds `styling.apply_report`, whose number kinds are `heltall`, `tall1`,
    `prosent1` and so on;
  - pass 2 writes the formulas through `e.to_excel(table, addr, const)`.
- `render/readback.compare(plan.addr, cells, sheet_block, fractions)` is Sheet = model. Every model cell must keep
  an address.
- `pipeline._draw()`:
  - writes data → lines sheets → layout → `write_answer` → `write_charts` → calculate → readback → save spec →
    arrange;
  - `reply()` builds the tool reply.
- `render/excel.rewrite_grid` writes from `Cells(1,1)`. That must change to the table origin (Task 9).
- `xl/drill_xl.write_lines` writes the title in A1, the source in A2, and the table from `hdr`. `model/lines.CUBE_CELL`
  is the cube value that `CubeRef` reads (`pipeline._write_lines` writes C2/D2). **CUBE_CELL must not move.**
- `xl/sheets.log_source` is v1 code and is never called in v2, so it is removed in Task 11.

---

### Task 1: Scenario headers and column groups (model)

**Files:**
- Modify: `model/spec.py` (TableBlock), `model/blocks/common.py`, `model/blocks/table.py`. Also update the
  callers of `col_headers`: kpi, ranking, matrix, concept and drivers (grep `col_headers(`).
- Test: `tests/v2/test_v5_headers.py`.

**Header text.** Replace `col_headers(cols)` with `col_headers(cols, *, period: str, ly_period: str, fc: str)`. Map
the column codes as follows:
- value → `AC <period>`
- ly → `PY <ly_period>`
- budget → `BU <period>`
- forecast → `FC <fc>`
- var → `Δ<PY|BU|FC>`, using the base to its left
- var_pct → `Δ<..> %`
- share → `Share %`
- amount → `Amount` (bridge, unchanged)

**Period labels.** Add `period_label(months)`:
- one month gives `Oct 26`;
- Jan..closed of one year gives `YTD Oct 26`;
- any other span gives `Jan–Sep 26`, or `Nov 25–Oct 26` across years.

For ly, take the label of the months shifted by −12. `fc` is the short forecast name
(`resolver.latest_forecast()`, e.g. `Forecast Q326`, shown as `Q326`). A cube without plan has `fc=""`.

**Column groups.** Add `TableBlock.periods: list[str] | None = None`, with each entry a period expression.
- When it is set, `build_table` makes one group per entry. Each group repeats `columns`.
- A spacer column (code `gap`, header `""`, cells `None`) sits between groups.
- Line keys stay the same. Column indexes run across the groups.
- Each value, ly, budget or forecast pick passes `base_months=resolve(expr, closed)`. For ly, `months()` shifts
  by −12 as it does now.
- **Guard** (spec validation or `guard_spec`): every group's months must be inside the query's months. Otherwise
  raise `GuardError("table periods 'X' fall outside query 'q' periods 'Y'", "widen the query periods or drop the
  group")`.

**Derived cells.** `derived()` and `base_of()` must find the base within the **same group**: search left only up
to the previous `gap`.

**Materiality.** `findings.material()` reads only the **last** group (YTD), so a material line is not listed
twice.

**Stable keys (display headers are not keys).** `BlockOut` gains `keys: list[str]`, parallel to `headers`.
- Keys are today's names from the current `col_headers`: `Actual`, `Last year`, `Budget`, `Forecast`,
  `Variance`, `Variance %`, `Share %`, and the `… vs …` forms.
- `findings.key_values` keys by `o.keys`, never by `o.headers`.
- The **last** group keeps the plain keys, so goldens and answer memory stay stable month to month. Earlier
  groups get the key prefix `<expr> · `, e.g. `closed · Actual`.
- `gap` and `material` columns are left out of `key_values`.

As a result, the test set stays a hard gate through every task (`testset run` = 9/9).

**Tests:**
- [ ] `key_values` of a single-group table equals today's keys exactly. A month+YTD table has plain keys for YTD
  and `closed · …` for the month group.
- [ ] Headers for value/ly/var/var_pct over `closed-ytd` with closed Oct 26 are `AC YTD Oct 26`, `PY YTD Oct 25`,
  `ΔPY`, `ΔPY %`.
- [ ] Headers for budget and forecast (FinPlan) are `BU YTD Oct 26` and `FC Q326`.
- [ ] `periods=["closed","closed-ytd"]` gives 4 + 1 + 4 columns, with the month group's values equal to the Oct
  26 facts only, the YTD group summing Jan–Oct, and the gap cells `None`.
- [ ] The var in group 2 references the base in group 2, not group 1.
- [ ] A group outside the query's months raises a GuardError that names the fix.
- [ ] `material()` lists a line once, judged on the YTD group.
- [ ] Run, watch it fail, implement, then run pytest. **Expect** existing tests that assert the old header texts
  to fail: update them to the new texts in the same commit and list them in the commit message.
- [ ] Commit: `feat(v2): scenario headers and month/YTD column groups`.

### Task 2: P&L chain and code column (model)

**Files:**
- Modify: `model/blocks/table.py`, `model/blocks/common.py` (`Line.code: str = ""`), `model/spec.py`
  (`TableBlock.chain: bool = True`).
- Test: `tests/v2/test_v5_chain.py`.

**Chain.** Add `CHAIN = ["Revenue", "COGS", "Gross margin", "Opex", "EBITDA before other items", "Other items",
"EBITDA"]` (semantic metric names), with kinds `Gross margin` → `sum`, `EBITDA before other items` → `sum`, and
`EBITDA` → `grand`.
- When a metric-row table (`rows == "metrics"`, `chain=True`) includes a metric that resolves to `EBITDA`:
  1. insert the missing chain metrics in chain order;
  2. keep every non-chain metric after the chain, in the order given;
  3. skip any chain metric that `resolver.metric(name)` cannot resolve in the query's cube, and do not raise.
- Line keys: the metrics given keep their `m{i}` key **by their original index**, so existing goldens keep their
  keys. Inserted chain lines get the keys `c{n}`.

**Code column.** `Line.code` is the first member of the metric in that cube, without its prefix (`ac.C30` → `C30`,
`ac.EBITDA_BOI` → `EBITDA_BOI`; FinPlan aliases unchanged). Entity and product lines get the member code without
its `un.SY.` or `pd.SY.` prefix.

**key_values cap.** `Line.added: bool` marks the inserted chain lines. `findings.key_values` fills the 40-line cap
with the non-added lines of every block first, then the added ones. Added lines can then never push a golden line
out. Test this with a spec near the cap.

**Departure.** `chain=False` is allowed. It shows up as a departure note (Task 6, Notes zone: "P&L chain shortened
on request").

**Tests:**
- [ ] `metrics=["EBITDA"]` on FinMax gives the 7 chain lines in order, with the kinds above and the codes C30, C40S,
  C48, C70, EBITDA_BOI, C73, EBITDA. Check the codes against the resolver fake, and adapt them if the fake's
  members differ.
- [ ] `metrics=["EBITDA","Revenue"]` keeps the keys `m0` (EBITDA) and `m1` (Revenue). The inserted lines are
  `c*`.
- [ ] `chain=False` leaves the metrics as given.
- [ ] A FinPlan cube without "Other items" skips it.
- [ ] Commit: `feat(v2): complete P&L chain and code column in metric tables`.

### Task 3: The Material column and threshold params (model)

**Files:**
- Modify: `model/expr.py` (new node), `model/analyze.py` (params), `model/spec.py`
  (`TableBlock.material: bool = True`), `model/blocks/table.py`.
- Test: `tests/v2/test_v5_material.py`.

**The expression.** Add a new Expr `Material(var: Expr, pct: Expr | None, a: Param, pct_t: Param, a10: Param)`.
- `evaluate` returns 1.0 or 0.0 under the same rule as `findings.material`. A `None` var gives 0.0.
- `to_excel` gives `IF(OR(ABS(v)>=a,AND(ABS(p)>=pct,ABS(v)>=a10)),1,0)`. Without `pct`,
  `IF(ABS(v)>=a,1,0)`.
- All three thresholds are Param cells. There are no constants in the formula (tn-formelstil).

**Params.** Generalise `Result.floors` into `Result.params: dict[str, tuple[str, float]]` (param key → (label,
value)), keeping `floors` as an alias, or migrate its single use in the layout. Keys:
- `floor.<metric>` → `Variance % floor: <metric>`
- `mat_abs.<metric>` → `Materiality: <metric> (NOK)`
- `mat_pct.<metric>` → `Materiality %: <metric>`
- `mat_small.<metric>` → `Materiality with %: <metric> (NOK)`, with value abs/10
- `unit_mnok` → `Unit: MNOK`, 1 000 000 (added by the layout in Task 5; the model does not need it)

**The column.** When a table has `material=True` and a var column, append one last column, code `material`,
header `Material`. Its cell per line is `Material(...)` over the **last group's** var and var_pct against the
first base. A line without a metric, or a check line, gets `None`.

**Readback.** `material` cells are numbers 0/1, so readback compares them as amounts. Add `material` to a new
`MARK_CODES` set so the layout gives them the marker number format (Task 5).

**Tests:**
- [ ] `Material.evaluate` matches `findings.material` on a grid of var and pct values: above abs, pct only above
  pct with var above a/10, below both.
- [ ] `to_excel` with a param address map gives the exact string above.
- [ ] A table with var gets the Material column, and its values equal the `material()` hits (1 exactly for the
  lines listed).
- [ ] `material=False` leaves it out.
- [ ] Commit: `feat(v2): live Material marker column with labelled thresholds`.

### Task 4: Chart blocks (spec, guard, data)

**Files:**
- Modify: `model/spec.py` (ChartBlock, Block union, `Plan.title`), `model/validate.py` (guard),
  `model/blocks/__init__.py`.
- Create: `model/blocks/chart.py`.
- Test: `tests/v2/test_v5_chart_block.py`.

**The spec.** Add `Plan.title: str = ""`, the message title of the sheet. Then:

```python
class ChartSource(_M):
    block: int | None = None          # index into spec.blocks (a table-like block already on the sheet)
    query: str | None = None          # own data: query id + metrics/periods/rows like a TableBlock
    metrics: list[str] = []
    periods: list[str] | None = None
    rows: Literal["metrics", "entities", "products"] = "metrics"
    metric: str | None = None
    columns: list[Column] = ["value", "ly"]
    # validator: exactly one of block / query

class ChartBlock(_M):
    type: Literal["chart"]
    kind: Literal["waterfall", "line", "column", "bar", "stacked", "stacked100", "combo", "variance"]
    title: str                        # the message, required, non-empty
    source: ChartSource
    series: list[Series | Literal["var"]] = []     # default by kind (see chart.py DEFAULT_SERIES)
    highlight: list[str] = []         # line labels to colour; the rest grey
    sort: Literal["desc", "asc", "none"] = "none"  # bar default desc (validator)
```

**The builder** (`build_chart` in `model/blocks/chart.py`):
- With `source.block`, return a BlockOut of type `chart` with **no lines**. It points at that block:
  `BlockOut.chart_of = "b<idx>"`, a new optional field.
- With `source.query`, build the data as a table (reuse `build_table` on a synthetic TableBlock with
  `material=False` and `chain=False`). Return it as type `chart_data` with title `Chart data: <title>`, so it is
  evaluated and read back like any table.

`DEFAULT_SERIES` gives the default series per kind:
- waterfall: amount
- line: value and ly
- column: value and ly, or value and budget
- bar: value
- stacked: value
- stacked100: value
- combo: value plus var_pct
- variance: var

**Guard (hard, NEW questions only).** A **new** `fin_analyze(spec)` (server path `_new_question`, not patch,
refresh or redraw), with no chart block and `plan.depth != "quick"`, raises
`GuardError("the answer has no chart", "add a chart block, e.g. {type:'chart', kind:'waterfall', title:'<message>',
source:{block:<bridge index>}}")`.
- Implement it as a separate `guard_chart(spec)` called only from that path.
- Every other path, including quick answers, stored pre-v5 specs, patches, `fin_refresh`, redraw and the test-set
  replay, gets an **implicit chart** when there is none. `compute` appends a `column` chart on the first table
  block: a waterfall on the first bridge, or a line on the first trend, when one exists. The spec stored stays as
  given.
- Test that a pre-v5 stored spec without a chart (one of the test-set case specs, as stored) prepares, refreshes
  and redraws without an error, and gets the implicit chart.
- The existing `chart: bool` on trend and bridge blocks counts as a chart block when true: line for trend,
  waterfall for bridge. The guard accepts it.
- A `source.block` must point at an existing non-chart and non-commentary block. A chart of kind `waterfall` must
  point at a bridge, or have its own data with an `amount` series.

**Tests:**
- [ ] A spec with no chart is refused (standard), while quick gets an implicit chart.
- [ ] A bridge with `chart: true` passes.
- [ ] `source.block` out of range is refused.
- [ ] Own-data chart: a BlockOut `chart_data` whose values equal a table with the same query and metrics.
- [ ] `title` empty is refused.
- [ ] `bar` defaults to sort desc.
- [ ] The test-set case specs stay unchanged: they replay through the implicit-chart path. **Never** give a case an
  own-data chart (`source.query`), because the offline replay has no recording for its grids.
- [ ] Update only the unit-test fixtures that call the new-question path.
- [ ] Commit: `feat(v2): chart blocks, chart guard and own chart data`.

### Task 5: Answer layout: zones, columns, formats (render)

**Files:**
- Modify: `render/layout.py` (substantial), `render/excel.py` (`write_answer` widths), `xl/styling.py` (new
  number kinds).
- Test: `tests/v2/test_v5_layout.py`.

**Columns.**
- `CODE_COL = 2`, `LABEL_COL = 3`, `FIRST_COL = 4`.
- Titles, subtitles, sections, notes and text are written in **column B** (`TEXT_COL = 2`), so they line up with
  the code.
- Widths: A 2, B 9, C = `label_width` (now measured on column C), numbers 16, gap columns 2, Material column 9.

**Zones, in this order:**
1. Title: `plan.title` or `plan.purpose` or the question.
2. Subtitle: `NOK · <cubes> · <entity label of the first query> · <period label of the main table's last group>`.
   Add the closed month as now: `· closed Oct 26`.
3. `Key points` section: rows reserved here and filled in Task 6.
4. `Charts` zone: reserve `ceil(n_charts/2) * CHART_ROWS` empty rows (`CHART_ROWS = 18`). Record each chart's
   anchor (col B, or the col after 480 pt for the second chart in a row; compute it from the widths) in
   `ChartSpec`.
5. The table blocks in spec order: chart and commentary blocks are skipped in this pass. The first one is the main
   table.
6. `Checks` section: a row `Sheet = model`, with its value cell left empty and written after readback (Task 8).
   Then one row per `check` line of every block, holding
   `=IF(ABS(<addr of the check value>)<=<Tolerance param>,"OK","DIFF")`. Add the param `tolerance` → `Tolerance
   (NOK)`, 1.
7. `Notes` section:
   - the conventions note (spec §1.5, exact text);
   - the departure notes (`chain=False`; the existing deviation text if any);
   - the `Thresholds` sub-block: every param in `res.params`, plus `unit_mnok` and `tolerance`, as labelled input
     cells;
   - Findings;
   - Caveats.
8. `Chart data` blocks (`chart_data` outs), each as a normal table section titled `Chart data: <title>`.
9. The Source note:
   `Source: Oracle EPM Cloud (<cubes>) · table Data_Q<n> (sheet Q<n> Data) · Sources sheet`.

**Remove** the Plan zone (`Plan`, `Question:`, hypotheses) and the old commentary rendering. Commentary moves to
the Key points in Task 6.

**Number kinds.** Add these to `styling.NUM_ALIAS`, through `tn` or as a local override applied after the styles,
mirroring how `TALL1_THOUSANDS` is handled:
- `avvik` → `[Color10]▲ #,##0;[Red]▼ (#,##0);"–"`
- `avvikpst` → `[Color10]▲ 0.0 %;[Red]▼ (0.0 %);"–"`
- `markor` → `"●";"●";""`

`_num(code, stock)`:
- var → `avvik`, or `tall1` for stock
- var_pct → `avvikpst`
- material → `markor`
- share / key → `prosent1`
- else `heltall` / `tall1`

`NumberFormat` (not Local) codes are en-US, so check that `apply_report` sets these through `NumberFormat`. It
uses `NumberFormatLocal` at line ~446: add a branch, and test it through the fake.

**Headers.** Header row cells come from `o.headers`. Gap columns have an empty header.

**Lines.**
- Each line writes `ln.code` in column B (grey; `apply_report` role `code_cols: ["B"]`, a new role that sets the
  font colour grey and left alignment).
- It writes the label in C, indented by level as now.

**Tests** (on SheetPlan only, no COM):
- [ ] The zone rows come in the order above. Title in B1.
- [ ] No cell contains "Plan" or "Hypothesis".
- [ ] A month+YTD table puts headers in D..G, a gap in H, then I..L, then Material in M.
- [ ] The `num` kinds per column match the rules above.
- [ ] The Checks rows hold the IF formula with the tolerance param address.
- [ ] The Thresholds rows hold every param, and `p.params` has their addresses.
- [ ] Every model cell key has an address (readback coverage).
- [ ] The Chart data section comes after Notes.
- [ ] `label_width` measures column C.
- [ ] Update `test_v2_layout*`, `test_v2_excel*` and the 2D label-width test for the moved columns in the same
  commit.
- [ ] Commit: `feat(v2): fixed answer zones, code column, scenario number formats`.

### Task 6: Key points (render)

**Files:**
- Modify: `render/layout.py`, `model/spec.py` (`CommentaryBlock.key_points` max 5, validator).
- Create: `render/keypoints.py`.
- Test: `tests/v2/test_v5_keypoints.py`.

**Choosing the points.** `key_points(main: BlockOut, res, drivers, addr) -> list[KeyPoint]`.
- Take the material lines of the main table (`res.material` with `key` starting with `<main.key>.`), ordered by
  |variance| descending, at most 4.
- If none, take the main table's `grand` line, else its first line.
- `KeyPoint = (formula: str, cause: str | None)`.

**The formula.** One live formula per point, built from cell addresses in the **last group**:

```
="<label> <period header> "&FIXED(<AC addr>/<unit_mnok addr>,1)&" MNOK; "&FIXED(<Δ addr>/<unit_mnok addr>,1)&" MNOK vs <PY|BU|FC> ("&IF(<Δ% addr>="","n/a",FIXED(<Δ% addr>*<hundred addr>,1)&" %")&")"
```

- `<hundred addr>` is a param `hundred` → `Percent factor`, 100. Constants stay out of formulas.
- `<label>` and `<period header>` are text literals: labels are allowed in formulas.
- When the main table has no var column, use the short form `="<label> <period> "&FIXED(<AC>/<unit>,1)&" MNOK"`.

**The cause.** Match the commentary driver whose `line` equals the point's line label (case-insensitive), or
whose `line` equals the metric name.
- With `basis=data`: `"   <text> (data: <evidence>)"`.
- With `basis=hypothesis`: `"   <text> (hypothesis — not verified in data)"`.
- No match for a material line: `"   Cause not yet explained"`.
- No match for a non-material fallback point: no cause row.

Then come the agent's free `key_points` (static text), until there are 5 points in total. The layout writes
`•  ` + formula as the formula
`="•  "&<rest>`, and the cause as a `text` row.

**Readback.** Key point cells are text, so they are not in `p.addr`. Add a layout check (Task 13) that the formula
cells do not evaluate to an error.

**Tests:**
- [ ] For a fake result with 2 material lines, two points come out in |var| order, and the formulas equal the
  exact strings, with addresses from a fake addr map.
- [ ] A driver match gives the data form, a hypothesis gives its form, and a missing driver gives
  `Cause not yet explained`.
- [ ] No material lines gives the grand line.
- [ ] Free key points cap the total at 5.
- [ ] More than 5 `key_points` in the spec are refused by the validator.
- [ ] Commit: `feat(v2): live key points with cause sentences`.

### Task 7: Chart drawing (render)

**Files:**
- Rewrite: `render/charts.py`. Modify: `render/layout.py` (ChartSpec v2).
- Create: `render/chartdata.py` (the waterfall helper maths).
- Test: `tests/v2/test_v5_charts.py`.

**ChartSpec v2:**
- `kind`
- `title`
- `cat_range` (an address range of category labels, column C)
- `series: list[(name, range, role)]`, where role is `ac`, `py`, `bu`, `fc`, `var` or `pct`
- `anchor` (cell)
- `width=480`
- `height=260`
- `millions`
- `sort`
- `highlight: list[int]` (point indexes)

The layout builds it from the referenced BlockOut's rows and columns, using the **last group**.

**Waterfall.** `chartdata.waterfall_helpers(n_steps)` returns, for the bridge rows start, steps and end, the four
helper columns base, up, down and total as formulas over the amount column. Each formula references the previous
row's running end. Use `MAX`/`MIN` so a running total that crosses zero stays correct: a `neg_base` series
handles below-zero bases.
- The helper columns are written to the right of the bridge table, under a small header `Chart helper` in grey,
  and are not part of the model.
- Test the maths in Python, through a function that evaluates the same rules: steps +5, −8, +2 from start 10 give
  bars that rebuild 10 → 15 → 7 → 9, including a zero crossing (start 3, step −8).

**Drawing** (`write_charts(ws, plan)`). Every chart:
- `ws.Shapes.AddChart2(-1, xlType, left, top, 480, 260)`;
- name it `smartass chart`;
- add the series explicitly with `SeriesCollection.NewSeries`, `.Values` and `.XValues`;
- **then** set the formatting.

Per kind:
- column (51), bar (57, with `ReversePlotOrder`), line (4), stacked (52), stacked100 (53);
- combo: column plus a line on the secondary axis for pct;
- variance: two series, fav and unfav, with `Overlap = 100`;
- waterfall: stacked column (52) with base and neg_base invisible (no fill).

**IBCS fills:**
- ac: solid navy `0x5E2C00` (match tn navy from the `tn.rgb` palette);
- py: solid grey;
- bu: white fill with a navy border;
- fc: pattern `msoPatternWideUpwardDiagonal`;
- var: green or red by sign;
- highlight points: the TN accent;
- all other points grey when `highlight` is set.

**Format rules:**
- No gridlines, legend only when there are more than 3 series (on top).
- Data labels on, with number format `#,##0.0,,` when `millions` is true, else `#,##0.0`.
- Gap width 50.
- Title = the message, with the font from the TN title style.

A failing chart is caught. It yields `"chart '<title>' failed: <err>; fix: <hint>"` in the lint list, and the
others continue.

**Tests:**
- [ ] Use a COM fake recording the calls: per kind, the chart type, the series count and roles, the fills
  (navy/grey/outlined/hatched), no gridlines, the label format, and the order "series before format".
- [ ] A chart that raises gives one lint line and does not stop the others.
- [ ] The waterfall helper maths, including the zero crossing.
- [ ] Two charts sit side by side in one zone row, and a third on the next row.
- [ ] Commit: `feat(v2): IBCS charts by message kind (stacked waterfall, variance, bar, combo)`.

### Task 8: Answer writer, readback status, print setup (render + pipeline)

**Files:**
- Modify: `render/excel.py`, `pipeline.py` (`_draw`, `reply`), `xl/styling.py`.
- Test: `tests/v2/test_v5_writer.py`.

**Steps:**
- `write_answer` sets the widths from Task 5 and passes `code_cols` to `apply_report`.
- After `compare()`, `_draw` writes `OK` or `DIFF in <n> cells` into the Sheet = model status cell.
  `plan.status_cell` gives its address.
- **Print setup**, through a new `styling.print_setup(ws, last_col, last_row, footer)`:
  - landscape, `Zoom=False`, `FitToPagesWide=1`, `FitToPagesTall=False`;
  - `PrintTitleRows` = none (zones vary);
  - `CenterFooter = "Source: Oracle EPM Cloud · fetched <date> · page &P of &N"`;
  - gridlines off (`ActiveWindow` is not touched: use `ws.PageSetup.PrintGridlines=False`, and the display
    gridlines already set by `apply_report`).

**Tests** (with fakes):
- [ ] The status cell gets OK with no mismatches and the DIFF text otherwise.
- [ ] The print setup properties are set.
- [ ] The widths dict has B=9 and the gap columns 2.
- [ ] Commit: `feat(v2): sheet=model status cell and print setup`.

### Task 9: Data sheet provenance header and table style

**Files:**
- Create: `render/datasheet.py`. Modify: `render/excel.py` (`write_data`, `rewrite_grid`, `read_data`),
  `pipeline.py`.
- Test: `tests/v2/test_v5_datasheet.py`.

**The header block.** `header_rows(q, spec, pre, fetched_at, user) -> list[tuple[label, value]]` builds one block
per query, in this order:
1. Title `Q<n> Data — <question>`.
2. Question.
3. Query `<id>: <cube> · <scenario/version per series>`. Use `model.series.scenario` for each series in the query:
   value plus its compare list.
4. POV: entity label and code, `NOK`, the balance types used.
5. Periods: concrete `Jan 26 – Oct 26` and the relative expression.
6. Route: `HTTP` / `Smart View ad hoc`, from `pre.routes` and failed.
7. Fetched: `YYYY-MM-DD HH:MM <tz>` (static) plus `os.getlogin()`.
8. Rows: `=COUNTIFS(Data_Q<n>[Query],"<id>")`.
9. Control total: the static sum of `pre.frame` facts for the query, in a cell, beside
   `=SUMIFS(Data_Q<n>[Value],Data_Q<n>[Query],"<id>")` and a status
   `=IF(ABS(<static>-<live>)<=<tol cell>,"OK","DIFF")`. The tolerance cell is a labelled `Tolerance (NOK)` 1 in
   the header.
10. Sources: a hyperlink, added by Task 11.
11. Refresh: the fixed text from spec §2.1.

After the header there is one blank row, then the table header row.

**`write_data(ws, frame, q, header)`:**
1. write the header (TN styles: title, then label grey and value black);
2. create the table at `A<h+2>`;
3. apply the light style: `TableStyle = "TableStyleLight1"`, `ShowTableStyleRowStripes=False`;
4. set Value to `NumberFormat '#,##0;(#,##0);"–"'` with blue font `0xC00000` (BGR for blue) and right alignment;
5. set label columns as text, AutoFit capped at 40;
6. freeze panes below the table header, via `ws.Activate` only if the sheet is in our own workbook. Otherwise use
   the `SplitRow`/`FreezePanes` on its window, wrapped in try. Freezing is cosmetic, so a failure is ignored.

**`rewrite_grid` / `read_data`** use `lo.HeaderRowRange.Row` and `.Column` as the origin, never `Cells(1,1)`.
- **Old answers:** if the table's header row is 1 and the sheet has no header block, insert `len(header)+1` whole
  rows above row 1 first (`ws.Rows("1:k").Insert()`), then write the header.
- Then refresh the header values (fetched, control total) on every redraw.

**Tests** (with fakes):
- [ ] The header rows come in order with the exact labels.
- [ ] Control total static equals the sum of the query's facts, and the live formula strings are exact.
- [ ] The table is created at row h+2.
- [ ] Rewrite with the table at row 14 writes from row 14.
- [ ] An old A1 table gets rows inserted and the header written, and the table name is unchanged.
- [ ] The Value number format and the style are set.
- [ ] The existing `rewrite_grid` tests (the "never Value2 header+body as one rectangle" rule) still pass with the
  new origin.
- [ ] Commit: `feat(v2): provenance header and calm style on Q<n> Data`.

### Task 10: Lines sheets header

**Files:**
- Modify: `pipeline._write_lines`, and `xl/drill_xl.write_lines` only if it is needed.
- Test: `tests/v2/test_v5_lines_header.py`.

**Do not move** A1 (title), A2 (source), the `CUBE_CELL` cells or the table origin. Add the provenance as values
to the right of the existing header, in columns F:G of rows 1–4 (label and value):
- `Drill cell` (the cell dims)
- `Fetched` (static date and time plus user)
- `Lines total` (`=SUM(Lines_Q<n>_<alias>[<amount col>])`)
- `Tie-out` (`=IF(ABS(<lines total>-<CUBE_CELL>)<=tol,"OK","DIFF")`)

The amount column is the one `LinePick` sums. Reuse its name constant from `model/lines.py`.

**Tests:**
- [ ] The fake records values in F1:G4 with the exact formulas.
- [ ] CUBE_CELL is untouched.
- [ ] Commit: `feat(v2): provenance and tie-out on lines sheets`.

### Task 11: The Sources sheet (v2)

**Files:**
- Create: `xl/sources.py`. Modify: `pipeline._draw` and `redraw_job` (call it), `ops/refresh` (whichever module
  runs `fin_refresh`; grep it).
- Delete: v1 `log_source` / `_log_source` / `_sources_name` in `xl/sheets.py` if they are unused (grep first;
  keep `_sources_name` if sheetorder uses it).
- Test: `tests/v2/test_v5_sources.py`.

**`write_sources(t, q, spec, pre, data_sheet, header_addr) -> int`** returns the Sources row:
- Sheet name `Sources` (or `Sources SV` when the user's own `Sources` is not ours: keep the existing
  `_sources_name` rule).
- The log header is at row `K+2`, where K is the Known limits block height.
- Columns:
  - Fetched
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
  - Status
  - Data sheet
  - Supersedes
- One row per query. **Append only.** When a row for the same (Q, query id) already exists, the new row's
  Supersedes column says `row <r>`. Earlier rows are never edited.
- Data sheet is a hyperlink to `'Q<n> Data'!A1`. The Data header's `Sources` row gets a hyperlink back to
  `'Sources'!A<row>`.
- **Known limits** (rows 1..K): the title `Known limits`, then one line per distinct `warn` or `block` caveat text
  over the answers stored in this workbook (`answers.db` for this workbook, or `pre.caveats` merged with the lines
  already there). It is rewritten in place on every call, and the log is moved down with `Rows.Insert` when K
  grows.
- TN style: header bold navy, frozen below the log header, AutoFilter on, and the Question column capped at 60.
  The sheet goes last (sheetorder).

**Tests:**
- [ ] The first call creates the sheet with Known limits and one row per query.
- [ ] A second call for the same Q appends, with Supersedes pointing at the first row.
- [ ] The hyperlinks are set both ways.
- [ ] Known limits are deduplicated.
- [ ] A user-owned `Sources` sheet leads to `Sources SV`.
- [ ] Commit: `feat(v2): Sources sheet log with known limits`.

### Task 12: Reply and the pipeline wiring

**Files:** `pipeline.py`, `server.py` (INSTRUCTIONS, `fin_analyze` docstring).
- `_draw` order:
  1. data sheet with its header;
  2. lines sheets;
  3. layout;
  4. write_answer;
  5. charts (now from `plan.charts`, including the waterfall helpers);
  6. calculate;
  7. readback;
  8. status cell;
  9. Sources;
  10. save the spec;
  11. arrange;
  12. the layout check (Task 13).
- `reply()` adds `layout` (the findings from Task 13) and `charts` (titles drawn).
- `key_values` keeps its shape and keys: block title → line label → **stable key** (Task 1) → value.
- `INSTRUCTIONS` gets one line: "every answer has at least one chart block (kind by message); the sheet shows Key
  points first, the plan is not drawn".

**Test:** the existing pipeline/excel_job fake tests still pass. Add one test that the reply contains `layout` and
`charts`. Commit with Task 13.

### Task 13: Layout check

**Files:**
- Create: `render/check.py`. Test: `tests/v2/test_v5_check.py`.

`check(plan, ws, data_ws, sources_ok) -> list[str]` returns one line per finding with its fix. It runs on the plan
plus at most 4 COM reads:
1. The zone order follows `plan.zones` (recorded by the layout), and there is no Plan zone.
2. `len(plan.charts) >= 1` and at least one chart was drawn (`ws.ChartObjects().Count`).
3. Every `p.rows` line row has a `num` for each numeric column.
4. `ws.UsedRange.Text` scan: any cell text with `#` only or `E+` → "column too narrow / number shown as E+".
5. Key point cells do not hold an Excel error.
6. The data header status cell reads `OK`.
7. `sources_ok`.

A finding is reported in `reply.layout`. It never raises.

**Tests:**
- [ ] Each rule fires on a crafted fake and stays silent on a clean one.
- [ ] Timing: under 100 ms on a 200-row fake. Assert loosely, under 0.5 s.
- [ ] Commit: `feat(v2): fast layout check in every reply`.

### Task 14: Skill, references, pane prompt, docs

**Files:**
- `skills/smartass/SKILL.md`, which must stay under 5500 characters (`test_skill_tools` size test).
- Create `skills/smartass/reference/charts.md`, with one worked chart-block example per kind and the message →
  kind table from spec §3.2.
- `skills/smartass/reference/spec.md`: `periods`, `chain`, `material`, `plan.title`, and the ChartBlock schema.
- `knowledge/agent/plan.md`: the default P&L/cost table is
  `{"type":"table","query":"q","periods":["closed","closed-ytd"]}`, with a chart chosen by message.
- `addin/tn-claude/server.mjs` PANE_PROMPT: "every answer needs a chart block chosen by the message (see skill
  reference charts.md)". Update `npm test` if it asserts the prompt.
- `CHANGELOG.md` and `docs/DESIGN.md`: a presentation section.

**Tests:** pytest (skill tests) and `npm test`.

Commit: `docs(v2): presentation rules, chart reference, default month+YTD table`.

### Task 15: Test set check (STOP for Tom only if needed)

1. Run `uv run python -m smartass.testset run`. It must be 9/9 with the goldens unchanged: the stable keys come
   from Task 1, and the implicit charts from Task 4.
2. Run `testset candidate --offline` for all cases and list the **added** key_values lines (chain lines, month
   groups) per case.
3. Ask Tom whether the goldens should also lock the added lines. If he says yes, run `approve`. The existing
   values must be identical in either case, and any change to them means STOP.

### Task 16: Live check and finish (controller)

1. Run `fin_status(restart=True)`.
2. In a new scratch workbook (`target="new"`), run six answers:
   - FinMax 2BUS P&L month + YTD vs PY (chain, Material, waterfall chart);
   - FinPlan 2WHS actual vs Target + forecast (BU/FC headers, column chart);
   - a bridge (waterfall);
   - an un.SY cost-centre ranking (bar);
   - a monthly trend (line);
   - a FinLite lines block (Lines header, tie-out).
3. Expect Sheet = model OK, the layout check empty, the Sources rows present and the data header OK.
4. Render each Answer and Data sheet to PDF once, through a hidden Excel `DispatchEx` copy in scratch, and look at
   the images. Fix what looks wrong: patch the code, rerun the tests and render again.
5. Run the `fin_critic` loop on one standard answer.
6. Run the test set (after Tom approves the goldens), merge `v2/5-presentation` → `main`, run `scripts\bump.ps1`,
   and push.
7. Restart the pane server.
8. Append a "Live result" section to this plan, and update memory (`smartass-v2-design`, `excel-addin-pane` if the
   prompt changed).

---

## Live result (2026-11-05)

Three live rounds against EPM, all in new scratch workbooks. Sheet = model was OK every time, and the numbers equal
the earlier v1/v2 answers (Business EBITDA YTD Oct 26: 3 868 633 901 against 3 811 374 124 PY).
- **Answers run:**
  - FinMax P&L month+YTD with chain, bridge and waterfall (plus a patch);
  - FinPlan actual vs Target and forecast (column chart, BU/FC headers);
  - a Technology opex ranking (bar) and trend (line);
  - a FinLite lines drill (bar, tie-out OK, gap 0).
- **Found live and fixed:**
  - number formats must be in the user's locale (pywin32 late-bound COM): `styling.local_format`;
  - chart labels need U+00A0 for the MNOK scale: `chart_format`;
  - `General` is `Standard` under nb-NO, so formats are reset with `ClearFormats`;
  - multi-area series are refused, so they become a union reference text;
  - a bridge over steps missing from its query: block metrics are added to the query;
  - FinPlan has no other items in Target, so the chain stops at EBITDA there;
  - duplicate chart sugar;
  - waterfall axis and labels;
  - cost charts' axis and labels;
  - units by magnitude (NOK under 1 m);
  - key points for lines answers;
  - print area including charts;
  - code repeated in labels;
  - Lines print setup.
- **Test set:** 9/9 throughout, goldens unchanged (stable key_values keys). The chain is completed only on new
  questions and patches, so no case gained lines and no approval was needed.
- **Gates at the end:** pytest 1192 passed, ruff clean, npm test 2/2.

## Carried forward
1. Line-chart data labels can touch the lines when two series are close.
2. The standalone `fin_drill` lines path (`xl/drill_xl.py`) sets no page setup of its own.
3. Known limits on Sources (column P) is outside the printed area.
4. The accent colour (0,160,220) for chart highlights is a placeholder. Tom to confirm the TN accent.
5. A trend chart with several metrics draws every AC line in navy, so the metrics cannot be told apart. Give each metric its own shade, keep AC solid and PY grey, and show the plan dashed.
6. The redraw of a pre-v5 workbook was checked on a scratch copy of the 04.11 Wholesale book (5 answers, 5 data and 4 lines tables). The header was inserted above every A1 table, the tables survived, 0 formulas show an error, and Sheet = model was OK for all 5. A workbook with many queries (Q4: 7) gets a 60-row header above its table.


## 5c follow-up (05.11.2026, v0.2.9-0.2.11)

Tom: "fiks alt". Superseded above: a line without a driver gets no cause row (no "Cause not yet explained").
- One unit per set of key points (`shared_unit`: MNOK when any amount >= 1 m, two decimals under 50 000); a ledger
  gap to the cube keeps its own unit (1 NOK, never 0,00 MNOK).
- A base column (PY/BU/FC) empty on every line is dropped with its variance, variance % and the columns that read it
  (concept helpers); Material is rebuilt over the remaining variances; `BlockOut.no_base` makes the key points say
  `no PY (new)`; `BlockOut.dropped` keeps the chart lint quiet about it.
- A trend's AC column empty on every month is dropped (FinPlan KPI with only Target); the points say `no AC`.
- A trend starts at its first month with data (rows before launch go; Total and cumulative sums pruned; the period
  label and subtitle follow).
- A grid without one data cell is no data ("FinPlan: no data for ... in Actual"), not "Smart View did not know".
- Font lint reads the drawn block only (cells a redraw cleared fall back to the Normal font).
