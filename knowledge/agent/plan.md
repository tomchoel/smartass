# Writing the plan

The `plan` tells the engine and the critic what a good answer is. Write it before choosing blocks.

## Fields

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

- `kpi_tiles`: 2–6 headline figures first. `table`: the P&L or a metric by entity (`rows: "entities"`, `top`,
  `other`). `trend`: monthly path. `ranking`: top movers. `matrix`: metric by entity by month.
- `bridge`: an EBITDA bridge uses the full chain Revenue, COGS, Opex, Other items (FinMax); a missing step leaves
  a large "Other / unexplained".
- `drivers`: one per material metric (`depth` 2, `drill: true` at `deep`), so every material line has evidence.
- `lines`: for supplier, project or account questions, a cell with All Accounts, All Entities and All Time
  Periods (Periodic, FinLite); the lines must tie to the cube value.
- `commentary`: 3–5 `key_points`, and one driver per material line with `basis: "data"` and the evidence cell
  key, or `basis: "hypothesis"`.

## Before you build

- Look up past answers (`fin_lookup(kind="answer")`) for continuity: last month's driver, last score.
- Read the caveats from `fin_lookup` and the reply. Severity `block` stops or changes the answer; `warn` must be
  named in the commentary or notes (a reorg, an abolished unit, an account trap); `info` is context.
