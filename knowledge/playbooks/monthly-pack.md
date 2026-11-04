# Playbook: monthly pack

Use when the same set of questions is answered every month (a division's P&L, opex by cost centre, FTE, actual vs
plan). The deterministic part is the stored specs and `fin_refresh`; judgement goes into the commentary.

## 1. Build the pack once

- One workbook per pack (`target: "new"` for the first question, then `"session"`). One question per recurring
  view, numbered Q1, Q2 … in reading order: headline P&L first, then cost, FTE, plan.
- Use **relative periods only**, so the pack rolls forward: `closed` (the latest closed month), `closed-ytd`,
  `closed-12m`. Never a fixed month such as `Sep 26` in a pack question.
- Comparison: the same period last year (`compare: ["ly"]`) for P&L and cost in FinMax. Actual vs budget or
  forecast as its own question entirely in FinPlan, with a FinMax EBITDA line as the reconciliation.
- Depth `standard` for every pack question, with `purpose` and `hypotheses`, so each gets a critic brief.
- Blocks: `kpi_tiles` for the headline, `table`, a full EBITDA `bridge` (Revenue, COGS, Opex, Other items),
  `drivers` on each metric that is usually material, and a `commentary` block.
- Run the critic on each question once (`fin_critic(q)`) and store the score: the first month sets the bar.

## 2. Each month

1. Check `fin_status()`: the closed month moved (FinMax and FinPlan), no doctor problems, knowledge refresh not
   pending (otherwise the user runs `smartass knowledge refresh` and approves first).
2. `fin_refresh()` re-runs every question to the latest closed month. Read the result per question: `ok`, the
   closed month, and its `critic_brief`. A failed question does not stop the rest; fix it by patch and re-run.
3. Continuity: `fin_lookup(query, kind="answer")` shows last month's answers, scores and key values, so the
   commentary can say "the consultancy driver seen in August continued" or "reversed".

## 3. Critic per question

For each question, `fin_critic(q)` returns the brief and the rubric. Score the six criteria, then:
- material lines with `backed: false` → add or deepen a `drivers` block, or a `lines` block, then a data driver;
- caveats in `gaps` → name them in a key point;
- a missing comparison → add `ly` (or move the plan view to FinPlan).
Store the score with `fin_critic(q, score, notes)`. Below 4 after 2 rounds: tell the user what is still open.

## 4. The commentary patch

Rewrite the `commentary` block of each question by `fin_analyze(patch={"blocks": [...]}, q=n)` — `blocks` is a
list, so the patch carries the whole list, commentary included. Follow the variance-commentary playbook:
3–5 key points, biggest mover first, each driver with `basis: "data"` and its evidence cell key, or labelled
`hypothesis`. Quote numbers only from `key_values` of this month's reply; last month's numbers come from the
answer memory, named as such.

## Done when

Every question refreshed to the same closed month, every brief without unbacked material lines, every critic
score at least 4 and stored, and the user told which sheets changed. The user saves the workbook.
