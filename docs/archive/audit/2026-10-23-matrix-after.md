# Live matrix 2026-10-23 (matrix-after)

- Time: 2026-10-23T15:07:14
- Engine: 0.2.0
- Atlas age: 9.7 days
- Options: {'only': '', 'repeat': 1, 'fresh': False, 'keep': False}
- Before: 2026-10-23-matrix-baseline.json
- Totals: {'PASS': 32, 'SLOW': 7, 'FAIL': 1, 'LIMIT': 3}

| ID | Tags | Result | Before | s | Error class |
|---|---|---|---|---:|---|
| S0.status | smoke | PASS | PASS | 0.7 |  |
| S1.find | lookup | PASS | PASS | 1.9 |  |
| S1.query.finmax | lookup | SLOW | SLOW | 30.4 |  |
| S1.query.finlite | lookup | SLOW | PASS | 6.1 |  |
| S1.query.finplan | lookup | PASS | PASS | 6.3 |  |
| S1.bad.balance | lookup, input | PASS | PASS | 0.1 |  |
| S1.bad.quoted | lookup, input | PASS | PASS | 0.1 |  |
| S1.bad.dictrows | lookup, input | PASS | FAIL | 1.9 |  |
| S1.bad.member | lookup, input | PASS | PASS | 2.6 |  |
| S2.build | build | SLOW | SLOW | 38.5 |  |
| S2.patch | build | PASS | PASS | 10.0 |  |
| S2.plan | build | PASS | PASS | 25.1 |  |
| S3.division_pnl.tec | template | PASS | PASS | 52.6 |  |
| S3.division_pnl.sy | template | PASS | PASS | 46.3 |  |
| S3.monthly_trend.tec | template | PASS | PASS | 32.2 |  |
| S3.monthly_trend.sy | template | PASS | PASS | 45.1 |  |
| S3.actual_vs_plan.tec | template | PASS | PASS | 39.7 |  |
| S3.actual_vs_plan.sy | template | PASS | FAIL | 40.1 |  |
| S3.cost_base.tec | template | PASS | PASS | 26.8 |  |
| S3.cost_base.sy | template | PASS | PASS | 41.4 |  |
| S4.scope | concept | SLOW | SLOW | 80.7 |  |
| S5.drill.cold | drill | PASS | PASS | 71.6 |  |
| S5.drill.warm | drill | PASS | PASS | 58.1 |  |
| S5.drill.quarter | drill | PASS | PASS | 59.0 |  |
| S5.drill.report | drill | SLOW | PASS | 98.6 |  |
| S6.xl | xl | FAIL | FAIL | 90.5 | ToolFailed: xl_revert: Error executing tool xl_revert: Sheet 'XL test' has content now; pass force=true to delete it anyway. |
| S7.refresh | refresh | PASS | PASS | 26.2 |  |
| S8.sso | inject | PASS | PASS | 80.5 |  |
| S8.dialog | inject | SLOW | FAIL | 80.6 |  |
| S8.editmode | inject | SLOW | FAIL | 60.3 |  |
| S8.parallel | inject | PASS | PASS | 57.6 |  |
| S8.kill.excel | inject | PASS | PASS | 25.4 |  |
| S8.kill.edge | inject | PASS | PASS | 88.0 |  |
| S8.cold | inject | PASS | PASS | 104.8 |  |
| S8.sleep | inject | PASS | PASS | 92.3 |  |
| S10.1.bundle | feature | PASS | PASS | 1.8 |  |
| S10.3.reload | feature | PASS | PASS | 10.7 |  |
| S10.5.textfixes | feature | LIMIT | LIMIT | 0.0 | Limit: user fix + override of a shipped key verified in-process (repair_text); the drill part is not tested: the # lines hit no shipped text-fix key |
| S10.6.pack | feature | LIMIT | LIMIT | 0.4 | Limit: pack has no runtime consumer; wiring only verified (pack enabled, source_product_list_#.json -> C:\dev\smartass\packs\produktregnskap\knowledge\source_pr |
| S10.7.style | feature | PASS | PASS | 28.3 |  |
| S10.8.coldexcel | feature | LIMIT | LIMIT | 0.1 | Limit: the user's own Excel is open (pid [#]): a cold start needs none |
| S10.9.options | feature | PASS | PASS | 26.9 |  |
| S10.10.http | feature | PASS | PASS | 0.2 |  |

## Error classes

| Count | Error class | Scenarios |
|---:|---|---|
| 1 | ToolFailed: xl_revert: Error executing tool xl_revert: Sheet 'XL test' has content now; pass force=true to delete it anyway. | S6.xl |
| 1 | Limit: user fix + override of a shipped key verified in-process (repair_text); the drill part is not tested: the # lines hit no shipped text-fix key | S10.5.textfixes |
| 1 | Limit: pack has no runtime consumer; wiring only verified (pack enabled, source_product_list_#.json -> C:\dev\smartass\packs\produktregnskap\knowledge\source_pr | S10.6.pack |
| 1 | Limit: the user's own Excel is open (pid [#]): a cold start needs none | S10.8.coldexcel |

## Dialogs seen

| Scenario | s into scenario | PID | Title | Text | Clicked | Killed |
|---|---:|---:|---|---|---|---|
| S1.query.finmax | 22.2 | 122256 | Login |  |  |  |
| S10.9.options | 15.8 | 112412 | Login |  |  |  |
