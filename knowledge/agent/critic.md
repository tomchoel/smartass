# Critic rubric

Apply to every `standard` or `deep` answer and to every pack question, after the commentary is written. Read the
`critic_brief` (from the reply or `fin_critic`): plan, material lines with `backed`, comparisons, caveats, checks,
findings, commentary drivers and `gaps`. Score each criterion 1–5 as a controller would.

| # | Criterion | 2 = weak | 5 = controller-ready |
|---|-----------|----------|----------------------|
| 1 | **Right question** | Unasked analysis dominates, or the plan's purpose is only partly answered. | Every part of the purpose is answered, nothing unasked crowds it. |
| 2 | **Numbers reconciled** | A check is not OK, Sheet = model shows a DIFF, or a gap (bridge Other, lines vs cube) is unexplained. | All checks OK, Sheet = model OK, every remaining gap named and explained. |
| 3 | **Comparison basis** | P&L or cost without last year; plan figures outside FinPlan; periods not like-for-like. | Same period last year for P&L and cost; budget or forecast in FinPlan when plan is asked (with the FinMax EBITDA reconciliation line); like-for-like months. |
| 4 | **Drivers backed by data** | Material lines without a driver, or drivers stated as fact with no evidence. | Every material line has a driver; each rests on `data` evidence (drivers block, lines block, cell key); hypotheses are labelled as such. |
| 5 | **Caveats respected** | A warn or block caveat is ignored. | Every warn and block caveat is reflected in the commentary or the notes. |
| 6 | **Clear for a controller** | Long, jargon, numbers not in NOK format, conclusion buried. | 3–5 key points first; full NOK, negatives in parentheses in tables, minus sign in text; plain words. |

A 3 is between: the criterion is mostly met with one clear miss. A 1 means the criterion is not addressed at all.

## Score and fix

- The score is the rounded mean of the six. It must be **at least 4** for standard or deep answers and packs.
- Every `gaps` item in the brief costs at least one point on its criterion until it is fixed.
- Below 4: list the fixes (one line each, the criterion number first), apply them by `fin_analyze(patch=...)`, and
  score again. At most **2 rounds**; after that, store the score and tell the user what is still open.
- Store the final score with `fin_critic(q, score, notes)`; the notes name the weakest criterion and why.
- Never raise a score by removing a material line, a caveat or a comparison: fix the answer instead.
