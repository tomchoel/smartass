# Writing the plan

The `plan` tells the engine and the critic what a good answer is. Write it before choosing blocks.

## Fields

- **title**: the message of the answer in one sentence, shown as the sheet title (empty = the purpose). The
  plan itself is not drawn on the sheet.
- **mode**: `adhoc` (default: investigative/ad-hoc workbook without forced P&L chains or mandatory charts) or
  `pack` (executive monthly pack with P&L chain completion, mandatory charts, and critic pass).
- **purpose** (required unless `quick`): one sentence naming the decision or question, e.g. "Explain Business opex
  YTD vs last year and name the cost centres behind the increase".
- **hypotheses**: what you expect and will confirm or reject from data, e.g. "Consultants up on the BSS
  migration". Each one ends as a data driver or is reported as rejected.
- **audience**: `controller` (default: full detail, reconciliations), `management` (key points and the bridge),
  `analyst` (detail and ledger lines).
- **depth**: `quick` = one number or a short table, no critic; `standard` = table, comparison, drivers,
  commentary (critic for `pack`); `deep` = standard plus ledger lines, deeper drivers, and critic.

## Explore first (Tier 1, zero Excel)

- Before committing to an Excel layout, use `fin_tree` (browse dimensions like Counterparts, Projects, Systems, or
  account children) and `fin_slice` (multi-dimensional EPM queries), plus `fin_drill` -> `fin_lines` (filter/group
  or read-only SQL over FinLite ledger lines).

## Comparison and cube

- P&L and cost: the same period last year (`compare: ["ly"]`) unless the user asks otherwise.
- Actuals in FinMax. Actual vs budget or forecast entirely in FinPlan, with a FinMax EBITDA line as the
  reconciliation to actuals.
- Variance = current − comparison (positive = better); costs are negative.

## Which block

- `table`: e.g. `{"type":"table","query":"q","periods":["closed","closed-ytd"]}` shows month and YTD side-by-side.
  `rows`: `"metrics"` (set `chain: true` or `mode: "pack"` to inject the 7-line P&L chain), `"entities"`,
  `"products"`, `"accounts"` (sub-accounts under a parent), or `"entities_x_metrics"` (side-by-side units). Supports
  `Query.dimensions` for Counterparts/Projects/Systems.
- `chart`: required in `mode: "pack"` (`standard`/`deep`), optional in `mode: "adhoc"`. Chosen by message: a bridge
  `waterfall`, trend `line`, AC vs PY/BU `column`, ranking `bar`, variance `variance`, mix `stacked100`.
- `kpi_tiles`: 2–6 headline figures first. `trend`: monthly path. `ranking`: top movers. `matrix`: metric by
  entity by month.
- `bridge`: an EBITDA bridge uses the full chain Revenue, COGS, Opex, Other items (FinMax); a missing step leaves
  a large "Other / unexplained".
- `drivers`: `by: "entities"` (default), `"accounts"`, or `"products"` with configurable `cover`, `max_per_level`, and
  `include_other`; `drill: true` at `deep`.
- `lines`: for supplier, project, account, or multi-column FinLite ledger views (`where`, `exclude_reversals`,
  `drill_id` reuse); raw drills must tie to the cube value.
- `commentary`: at most 5 `key_points` (the engine writes the material lines first), and one driver per
  material line with `basis: "data"` and the evidence cell key, or `basis: "hypothesis"`. The driver becomes the
  cause sentence under that line's key point.

## Before you build

- Look up past answers (`fin_lookup(kind="answer")`) for continuity: last month's driver, last score.
- Read the caveats from `fin_lookup` and the reply. Severity `block` stops or changes the answer; `warn` must be
  named in the commentary or notes (a reorg, an abolished unit, an account trap); `info` is context.
