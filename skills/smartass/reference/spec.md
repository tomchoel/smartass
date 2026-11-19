# AnalysisSpec v2: details and block options

## A full example

```json
{"version": 2, "question": "Wholesale P&L year to date vs last year",
 "plan": {"depth": "standard", "purpose": "Explain the Wholesale EBITDA bridge YTD vs LY"},
 "queries": [{"id": "q1", "cube": "FinMax", "metrics": ["Revenue", "COGS", "Opex", "Other items", "EBITDA"],
              "entity": "un.SY.2WHS", "periods": "closed-ytd", "compare": ["ly"]}],
 "blocks": [{"type": "kpi_tiles", "query": "q1", "metrics": ["Revenue", "EBITDA"]},
            {"type": "table", "query": "q1", "title": "P&L", "periods": ["closed", "closed-ytd"]},
            {"type": "bridge", "query": "q1", "total": "EBITDA", "steps": ["Revenue", "COGS", "Opex", "Other items"]},
            {"type": "chart", "kind": "variance", "title": "Opex pulls EBITDA below last year", "source": {"block": 1}}]}
```

## Rules

- **Strict:** unknown fields are refused; a block's metrics and series must be in its query.
- **plan:** `depth` (`quick` | `standard` | `deep`), `title` (the message title of the answer sheet; empty = the
  purpose, else the question), `purpose` (required unless quick), `hypotheses` you will test. The plan is not drawn
  on the sheet.
- **periods:** `closed`, `closed-ytd`, `closed-12m`, `fy`, `Sep 26`, `Jan 26..Sep 26`. Relative periods roll
  forward in `fin_refresh`. An open month needs `preliminary: true`.
- **compare:** `ly`; `budget` and `forecast` (the latest) **only in FinPlan**: actual vs plan is done entirely in
  FinPlan (entity alias e.g. `1TEC Technology`; columns `["value", "budget", "var", "forecast", "var"]`).
- **products:** a query may list products; see [concepts.md](concepts.md).

## Blocks

| Block | Options |
|---|---|
| `table` | `rows`: `metrics` (default), `entities` + `metric`, or `products` + `metric`; `columns` from value, ly, budget, forecast, var, var_pct, share; `top`, `other`, `sort` (none, desc, asc, abs); `periods`, `chain`, `material` (below) |
| `kpi_tiles` | 1–6 `metrics` (2–6 read best), one `compare` |
| `trend` | `metrics`, `series` (value, ly, budget, forecast), `cumulative`, `chart` |
| `bridge` | `total`, `steps`, `base`. An EBITDA bridge uses every step: Revenue, COGS, Opex, Other items (FinMax), otherwise `Other / unexplained` is large |
| `ranking` | `metric`, `top`, `compare`; `rows: "products"` ranks the query's products |
| `matrix` | `metric`: children of the entity by month, with a children = parent check |
| `drivers` | `metric`, `base`, `depth` 1–4, `drill` (the leaves become ledger lines, see [drill.md](drill.md)) |
| `lines` | `cell`, `group_by` (supplier, project, account), `top`; needs no query |
| `concept` | `concept`, a FinMax/FinLite `query`; see [concepts.md](concepts.md) |
| `commentary` | `title`, `key_points` (at most 5), `drivers` |
| `chart` | `kind`, `title`, `source`, `series`, `highlight`, `sort`; see [charts.md](charts.md) |

**Commentary drivers:** `{"line", "text", "basis": "data", "evidence": "<cell key>"}` or `"basis": "hypothesis"`,
written from the reply's evidence (the material lines, drivers and lines blocks), never from memory.

## Table columns: periods, chain, material

- **periods:** column groups, one per period expression, side by side. The default P&L or cost table is
  `{"type":"table","query":"q","periods":["closed","closed-ytd"]}`: the closed month and YTD, each with the
  table's `columns`. Every group must fall inside the query's `periods` (a `closed-ytd` query holds both); a group
  outside it is refused. Leave `periods` out for one group over the query's periods. Charts and key points read the
  last group.
- **chain** (default `true`): a metric-row table that shows EBITDA gets the whole P&L chain (Revenue, COGS, Gross
  margin, Opex, EBITDA before other items, Other items, EBITDA), in chain order. The engine adds the missing chain
  metrics to the table and its query on a new question and on a patch. Set `false` to keep only your metrics.
- **material** (default `true`): a table with a `var` column gets a last Material column (● on a material line, a
  live formula on the thresholds in Notes).

## Headers and key points

- Column headers name the scenario and period: AC, PY, BU, FC, and Δ for a variance (e.g. `AC Oct 26`).
- Key points: at most 5. The engine writes one live number point per material line of the main table (at most 4),
  with the cause sentence from your commentary driver on that line under it. Your `key_points` fill the rest.

## Chart block

`{"type": "chart", "kind", "title", "source", "series", "highlight", "sort"}`

- `kind`: waterfall, line, column, bar, stacked, stacked100, combo, variance.
- `title`: the message (required, not blank).
- `source`: `{"block": <index>}` (a block in `blocks`, not a chart or commentary), or own data `{"query", "metrics",
  "periods", "rows", "metric", "columns"}` (laid out as a Chart data table). Give exactly one.
- `series`: column codes (value, ly, budget, forecast, var, var_pct, amount); empty = the kind's default.
- `highlight`: line labels in the accent (column and bar). `sort`: desc, asc, none (bar: desc).
- A waterfall must point at a bridge block. Kinds, defaults and one example per kind: [charts.md](charts.md).

## Layout findings

The reply's `layout` lists what the layout check found (zone order, no chart drawn, a number format, a cut-off
text or `###`, the Data header, the Sources rows). Each line ends with its fix. Apply the fix by patch.

## Variance and variance %

Variance = current − comparison (FinMax signs, so positive is better). Variance % = variance / |base|; it is
empty when the base is zero, and also when |base| is below the metric's floor: a tenth of its absolute
materiality for a NOK metric (500 000 NOK by default); a stock or ratio metric (FTE) only with its own floor. The
floors used sit in labelled cells in the answer's **Thresholds** section, and the formulas reference them.

## Follow-ups

`fin_analyze(patch=..., q=n)` applies a JSON merge patch to the stored spec: objects merge, lists such as
`blocks` or `metrics` are replaced whole, `null` removes a field. Default q is the active `Q<n> Answer`, else the
latest. The question's sheets are replaced (same names); anything typed on them goes with them.

## Members

Entities `un.SY.<code>` (2WHS, 2BUS, 2SSM, 2EI, 2TEC …), accounts `ac.<code>`, periods `Sep 26`. History is
restated to today's organisation: the 1 June 2026 reorg emptied 2CON into 2SSM + 2EI. Top levels per cube:
[cubes.md](cubes.md).
