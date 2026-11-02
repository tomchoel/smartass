# Variance commentary

Adapted from Anthropic financial-services `variance-commentary` (Apache-2.0; see NOTICE.md).

## Which lines get a comment

`sv_answer` returns `material`: the lines the engine flagged from the first variance column of the answer.

- **Material** if either is true (catalog `[materiality]`):
  - the variance is at least 10 % of the absolute comparison value **and** at least NOK 500k, or
  - the variance is at least NOK 5 million.
  - FTE and other levels are tested on the percentage only.
- **Always commented**: revenue, personnel cost and EBITDA, material or not.

A material line without a driver is a **Warning** in the review; an always-commented line without one is **Info**.

## Writing a driver

Write drivers in the answer spec: `commentary.drivers: [{line, text, basis, evidence?}]`, where `line` is the layout
line key it explains.

- A driver explains **why**, not **what**. "Consultancy up NOK 12m on the BSS migration in Q3" — not "Consultancy
  increased NOK 12m (35 %)".
- **Data first.** Look at the activity behind the line: `sv_drill` to journal lines or suppliers, a pull by
  sub-account or cost centre, FTE delta, volume × price. Then `basis = "data"` and `evidence` names the pull or
  drill it rests on.
- **Hypothesis.** An informed reason not shown in the data is allowed only as `basis = "hypothesis"`. The sheet
  shows it as "Hypothesis — not verified in data: …".
- **Unclear.** If neither works, write the driver text "driver unclear — flag for controller". Never invent one.

## Output

The drivers appear as key points on the answer sheet. In the chat, add a short narrative (3–5 sentences) on the
biggest movers.
