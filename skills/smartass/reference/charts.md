# Charts: choose the kind by the message

Every answer needs a chart. Write the message first, then pick the kind that shows it. A new `standard` or `deep`
question without a chart is refused. On every other path (quick, patch, refresh), a spec without a chart gets an
implicit one: on the first bridge (waterfall), else the first trend (line), else the first table (column), else a
bar or column on the first block with data. A `trend` or `bridge` with `chart: true` counts as a chart.

## Message → kind

| Message | Kind |
|---|---|
| Bridge from comparison to actual (EBITDA, cost) | `waterfall` |
| Development over time (6 months or more) | `line` |
| Few periods, or AC against PY/BU | `column`, at most 3 series |
| Ranking (cost centres, suppliers, products) | `bar`, top N plus Other, at most 15 bars |
| Variance by line | `variance`, bars from zero, green or red by effect |
| Parts of a total, in MNOK | `stacked`, at most 5 parts |
| Mix or share | `stacked100`, at most 5 parts, never a pie |
| Two units (MNOK and %) | `combo`, columns and a % line |

## The chart block

```
{type "chart", kind, title, source, series?, highlight?, sort?}
```

- **kind:** `waterfall`, `line`, `column`, `bar`, `stacked`, `stacked100`, `combo`, `variance`.
- **title:** the message in one sentence. Put the number in from `key_values` by patch after the first answer
  ("Opex explains 60 % of the EBITDA fall").
- **source:** one of two.
  - `{"block": i}`: a block already in `blocks` (its index, from 0). Not a chart or commentary block.
  - `{"query", "metrics", "periods", "rows", "metric", "columns"}`: own data. The engine lays it out as a
    "Chart data" table under Notes, with live formulas, and the chart reads it. The fields work as in a `table`
    block; `columns` defaults to `["value", "ly"]`.
- **series:** column codes of the source: `value`, `ly`, `budget`, `forecast`, `var`, `var_pct`, `amount`. Each must
  be a column of the source, or the chart warns. Defaults: waterfall `amount`; line and column `value` plus the
  base the data has (ly, else budget, else forecast); bar, stacked, stacked100 `value`; combo `value`, `var_pct`;
  variance `var`.
- **highlight:** line labels as shown in column C, drawn in the accent; the rest is grey. Column and bar only.
- **sort:** `desc`, `asc`, `none`. Bar defaults to `desc`. Bars are only re-ordered when the source is a ranking or
  lines block; otherwise the bars keep the table's order and the reply warns.

Rules the engine applies:

- A waterfall needs a bridge block: `source {"block": <bridge index>}`. It has no own data.
- A chart on a table with column groups (month and YTD) draws the **last** group (YTD).
- A chart on a table leaves out its total and subtotal rows (EBITDA, Gross margin, EBITDA before other items) and
  its checks; only a waterfall shows them. To chart EBITDA as a bar, use own data (no chain there).
- `stacked` and `stacked100`: each row is a part, each series column is a stack. Give two series (e.g. `value`,
  `ly`) to compare two stacks. More than 5 parts gets a warning.
- A `trend` or `bridge` with `chart: true` (the default) already draws a line or a waterfall. When you add a chart
  block on that block, set its `chart: false`, or it draws twice.
- A chart that fails is a warning with a fix, never a failed answer. The `layout` check reports an answer with no
  chart drawn.

## The answer the examples use

Each example below is one chart block added to the `blocks` of this spec. `source.block` points at its blocks:
0 = the P&L table (month + YTD), 1 = the bridge, 2 = the trend, 3 = the ranking.

```json
{"version": 2, "question": "Business P&L for the month and year to date vs last year",
 "plan": {"depth": "standard", "title": "Business EBITDA is below last year on higher opex",
          "purpose": "Explain Business EBITDA month and YTD vs LY and name the cost centres behind it"},
 "queries": [{"id": "q", "cube": "FinMax", "entity": "un.SY.2BUS", "periods": "closed-ytd", "compare": ["ly"],
              "metrics": ["Revenue", "COGS", "Opex", "Other items", "EBITDA", "Personnel cost", "Other opex"]},
             {"id": "m", "cube": "FinMax", "metrics": ["Opex"], "entity": "un.SY.2BUS", "periods": "closed-12m",
              "compare": ["ly"]}],
 "blocks": [{"type": "table", "query": "q", "periods": ["closed", "closed-ytd"],
             "metrics": ["Revenue", "COGS", "Opex", "Other items", "EBITDA"]},
            {"type": "bridge", "query": "q", "total": "EBITDA", "steps": ["Revenue", "COGS", "Opex", "Other items"],
             "chart": false},
            {"type": "trend", "query": "m", "metrics": ["Opex"], "chart": false},
            {"type": "ranking", "query": "q", "metric": "Opex", "top": 10}]}
```

## One example per kind

**waterfall**: the bridge from last year to this year.

```json
{"type": "chart", "kind": "waterfall", "title": "Opex explains most of the EBITDA fall YTD", "source": {"block": 1}}
```

**line**: twelve months of opex against last year.

```json
{"type": "chart", "kind": "line", "title": "Opex has run above last year every month since May",
 "source": {"block": 2}, "series": ["value", "ly"]}
```

**column**: AC against PY for Revenue and EBITDA (own data, so EBITDA is a plain line), EBITDA in the accent.

```json
{"type": "chart", "kind": "column", "title": "Revenue holds, EBITDA falls against last year",
 "source": {"query": "q", "metrics": ["Revenue", "EBITDA"], "columns": ["value", "ly"]},
 "series": ["value", "ly"], "highlight": ["EBITDA"]}
```

**bar**: the cost centres ranked by opex.

```json
{"type": "chart", "kind": "bar", "title": "Three cost centres carry most of the Business opex",
 "source": {"block": 3}, "series": ["value"], "sort": "desc"}
```

**variance**: the variance by P&L line, green or red by effect.

```json
{"type": "chart", "kind": "variance", "title": "Opex and COGS pull EBITDA below last year",
 "source": {"block": 0}, "series": ["var"]}
```

**stacked**: the parts of opex in MNOK, this year and last year.

```json
{"type": "chart", "kind": "stacked", "title": "Personnel cost is the larger part of opex in both years",
 "source": {"query": "q", "metrics": ["Personnel cost", "Other opex"], "columns": ["value", "ly"]},
 "series": ["value", "ly"]}
```

**stacked100**: the mix of opex, this year and last year.

```json
{"type": "chart", "kind": "stacked100", "title": "Personnel cost takes a larger share of opex than last year",
 "source": {"query": "q", "metrics": ["Personnel cost", "Other opex"], "columns": ["value", "ly"]},
 "series": ["value", "ly"]}
```

**combo**: MNOK in columns and the variance % as a line.

```json
{"type": "chart", "kind": "combo", "title": "Revenue grows, but EBITDA falls in % against last year",
 "source": {"query": "q", "metrics": ["Revenue", "EBITDA"], "columns": ["value", "ly", "var_pct"]},
 "series": ["value", "var_pct"]}
```
