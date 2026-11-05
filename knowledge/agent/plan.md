# Writing the plan

The `plan` tells the engine and the critic what a good answer is. Write it before choosing blocks.

## Fields

- **title**: the message of the answer in one sentence, shown as the sheet title (empty = the purpose). The
  plan itself is not drawn on the sheet.
- **purpose** (required unless `quick`): one sentence naming the decision or question, e.g. "Explain Business opex
  YTD vs last year and name the cost centres behind the increase". The critic checks the answer against it.
- **hypotheses**: what you expect and will confirm or reject from data, e.g. "Consultants up on the BSS
  migration". Each one ends as a data driver or is reported as rejected.
- **audience**: `controller` (default: full detail, reconciliations), `management` (key points and the bridge),
  `analyst` (detail and ledger lines).
- **depth**: `quick` = one number or a short table, no critic; `standard` = table, comparison, drivers,
  commentary, critic; `deep` = standard plus ledger lines and a second level of drivers.

## Comparison and cube

- P&L and cost: the same period last year (`compare: ["ly"]`) unless the user asks otherwise.
- Actuals in FinMax. Actual vs budget or forecast entirely in FinPlan, with a FinMax EBITDA line as the
  reconciliation to actuals.
- Variance = current − comparison (positive = better); costs are negative.

## Which block

- `table`: the P&L or cost table by default shows the closed month and YTD:
  `{"type":"table","query":"q","periods":["closed","closed-ytd"]}` (the query's periods `closed-ytd`). A table
  with EBITDA gets the whole P&L chain. A metric by entity: `rows: "entities"`, `top`, `other`.
- `chart`: every answer has one, chosen by the message (skill reference charts.md): a bridge `waterfall`, a
  trend `line`, AC against PY/BU `column`, a ranking `bar`, a variance by line `variance`, a mix `stacked100`.
- `kpi_tiles`: 2–6 headline figures first. `trend`: monthly path. `ranking`: top movers. `matrix`: metric by
  entity by month.
- `bridge`: an EBITDA bridge uses the full chain Revenue, COGS, Opex, Other items (FinMax); a missing step leaves
  a large "Other / unexplained".
- `drivers`: one per material metric (`depth` 2, `drill: true` at `deep`), so every material line has evidence.
- `lines`: for supplier, project or account questions, a cell with All Accounts, All Entities and All Time
  Periods (Periodic, FinLite); the lines must tie to the cube value.
- `commentary`: at most 5 `key_points` (the engine writes the material lines first), and one driver per
  material line with `basis: "data"` and the evidence cell key, or `basis: "hypothesis"`. The driver becomes the
  cause sentence under that line's key point.

## Before you build

- Look up past answers (`fin_lookup(kind="answer")`) for continuity: last month's driver, last score.
- Read the caveats from `fin_lookup` and the reply. Severity `block` stops or changes the answer; `warn` must be
  named in the commentary or notes (a reorg, an abolished unit, an account trap); `info` is context.
