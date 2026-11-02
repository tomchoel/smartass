# smartass benchmarks (2026-11-02)

Acceptance set for the pipeline rework (`2026-11-02-pipeline-architecture-design.md` §8). Expected numbers are not
stored here: each run checks the answer against an independent `sv_query` oracle at run time, so the set stays valid
as months close. Fixed values from the live matrix are noted where they exist.

Each run records: correct (oracle match ±1 NOK, all checks OK, no formula errors), gate outcome, time to first
visible output, total time.

## Answer questions

| ID | Question | Exercises | Gate | Origin |
|----|----------|-----------|------|--------|
| B01 | Wholesale EBITDA YTD Sep 26 vs last year and budget | P&L recipe, FinMax + FinPlan | clear | S11 |
| B02 | Business Norway EBITDA per month 2026 vs 2025, reconciled to accumulated YTD | free-form layout, checks, chart | clear | S2.build (YTD 3 496 062 985, var 89 346 421) |
| B03 | Follow-up to B02: add FinPlan Actual EBITDA YTD | follow-up adds a pull | — | S2.plan |
| B04 | Follow-up to B02: add a note and change the title | patch without data change | — | S2.patch |
| B05 | Technology actual vs plan, Sep 26 | actual-vs-plan recipe | clear | S3 |
| B06 | Top 15 cost centres by opex under Technology, Sep 26 | cost-base recipe | clear | S3 |
| B07 | Telenor Norge monthly revenue trend through Sep 26 | monthly-trend recipe, un.SY | clear | S3 |
| B08 | Follow-up to B07: split by business area | follow-up adds pulls and lines | — | new |
| B09 | "What do we earn on mobile subscriptions?" | concept recipe | clear | S4 |
| B10 | "EBITDA per product for Mobile Consumer" | infeasible → explanation + contribution-margin proxy | infeasible | new |
| B11 | "How are we doing on costs?" | ambiguous → 1–3 questions, no pull | ambiguous | new |
| B12 | Lines behind account 12881, un.SY, Sep 26 | drill, exact reconciliation | clear | S5 |
| B13 | EBITDA YTD for unit un.SY.5OK_W | sub-unit without FinPlan member → last-year-only warning | clear | S12 |
| B14 | Business full-year EBITDA: latest forecast vs budget | FinPlan scenarios/versions | clear | new |
| B15 | Wholesale full P&L (revenue, COGS, opex, EBITDA) actual vs budget vs last year, YTD | full P&L, many lines | clear | S11 full |

## Trap questions (gate quality)

Measured as false-go (pulled when it should have asked or refused) and needless-ask (asked when the question was
clear).

| ID | Question | Expected gate | Why |
|----|----------|---------------|-----|
| T01 | "What's our budget for Business?" | ambiguous | which version (Final/Working), which measure, which year |
| T02 | "Opex for Wholesale this year" | ambiguous | YTD to last closed month or full-year forecast |
| T03 | "EBITDA margin per customer" | infeasible | no customer dimension in the cubes; offer per segment |
| T04 | "Revenue in EUR for Technology" | infeasible or clear with warning | FinMax currency members; EUR may not be loaded |
| T05 | "Profit per product" | infeasible | opex not allocated to product; offer contribution margin |
| T06 | "Business EBITDA Sep 26 vs Sep 25" | clear | fully specified — asking would be a needless ask |
| T07 | "Q3 numbers for Consumer" | ambiguous | which measures |
| T08 | "Cost per FTE in IT" | clear with stated definition | FTE measure exists; cost definition must be stated |
| T09 | "How much did we spend on consultants?" | clear or ambiguous | account mapping must come from the catalogue, not a guess |
| T10 | "Forecast for October 26 actuals" | infeasible | actuals for an open month do not exist; offer forecast |
