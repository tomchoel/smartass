# Playbook: variance commentary

Use for the commentary on any standard or deep answer and on every pack question: why each material line moved,
from data, in controller language.

## Our rules (smartass v2)

1. **Which lines.** The reply's `material` list and the `critic_brief` say which lines need a driver; revenue,
   personnel cost and EBITDA are always commented. Quote numbers only from verified tool outputs (`fin_slice`,
   `fin_lines`, `fin_drill`, `fin_analyze` `key_values`).
2. **Comparison.** P&L and cost against the same period last year (FinMax) unless the question asks for plan.
   Actual vs budget or forecast is done entirely in FinPlan, with a FinMax EBITDA line as the reconciliation to
   actuals; never mix FinMax actuals with FinPlan plan in one variance.
3. **Signs and format.** Costs are negative; variance = current − comparison, so a positive variance is better;
   variance % = variance / |base|, empty when the base is zero. Tables in full NOK with negatives in parentheses
   `(1 000 000)`; text with a minus sign `−1 000 000`.
4. **Drivers basis.** Each driver in the `commentary` block is
   - `basis: "data"` with `evidence` = the cell key it rests on (a `drivers` block line, a `lines` block group,
     a table cell such as `b1.m1.2`), or
   - `basis: "hypothesis"`, shown on the sheet as not verified in data. Prefer data: investigate first with
     `fin_slice` / `fin_drill` -> `fin_lines`, or add a `drivers` block (`by: "entities"` or `"accounts"`,
     `drill: true` at deep) or a `lines` block before writing a hypothesis.
5. **Caveats.** Every warn or block caveat in the reply (reorg moves, abolished units, account traps) is named in
   a key point or a driver, by its id or its wording. A block caveat changes the conclusion; say how.
6. **Key points.** 3–5, biggest mover first, each one sentence with the amount and the reason.
7. **Close with the critic.** `fin_critic(q)`, fix the gaps by patch, store the score.

Where the vendored text below names v1 tools, read `sv_answer` as the `fin_analyze` reply and `sv_drill` as
`fin_drill` or a `lines` block. Its materiality is the knowledge layer's per metric (`fin_lookup`).

## Vendored method

The section below is quoted unchanged from `knowledge/advisor/variance-commentary.md`, adapted from the Anthropic
financial-services-plugins marketplace (<https://github.com/anthropics/financial-services-plugins>,
`month-end-closer/skills/variance-commentary`). Copyright Anthropic, PBC; licensed under the Apache License,
Version 2.0. See `knowledge/advisor/NOTICE.md` for the notice and the changes made.

> # Variance commentary
>
> Adapted from Anthropic financial-services `variance-commentary` (Apache-2.0; see NOTICE.md).
>
> ## Which lines get a comment
>
> `sv_answer` returns `material`: the lines the engine flagged from the first variance column of the answer.
>
> - **Material** if either is true (catalog `[materiality]`):
>   - the variance is at least 10 % of the absolute comparison value **and** at least NOK 500k, or
>   - the variance is at least NOK 5 million.
>   - FTE and other levels are tested on the percentage only.
> - **Always commented**: revenue, personnel cost and EBITDA, material or not.
>
> A material line without a driver is a **Warning** in the review; an always-commented line without one is **Info**.
>
> ## Writing a driver
>
> Write drivers in the answer spec: `commentary.drivers: [{line, text, basis, evidence?}]`, where `line` is the layout
> line key it explains.
>
> - A driver explains **why**, not **what**. "Consultancy up NOK 12m on the BSS migration in Q3" — not "Consultancy
>   increased NOK 12m (35 %)".
> - **Data first.** Look at the activity behind the line: `sv_drill` to journal lines or suppliers, a pull by
>   sub-account or cost centre, FTE delta, volume × price. Then `basis = "data"` and `evidence` names the pull or
>   drill it rests on.
> - **Hypothesis.** An informed reason not shown in the data is allowed only as `basis = "hypothesis"`. The sheet
>   shows it as "Hypothesis — not verified in data: …".
> - **Unclear.** If neither works, write the driver text "driver unclear — flag for controller". Never invent one.
>
> ## Output
>
> The drivers appear as key points on the answer sheet. In the chat, add a short narrative (3–5 sentences) on the
> biggest movers.
