# Live matrix 2026-10-23 (matrix-baseline)

- Time: 2026-10-23T13:03:24
- Engine: 0.2.0
- Atlas age: 9.7 days
- Options: {'only': '', 'repeat': 1, 'fresh': False, 'keep': False}
- Before: none
- Totals: {'PASS': 32, 'SLOW': 3, 'FAIL': 5, 'LIMIT': 3}
- Fresh profile: `2026-10-23-matrix-baseline-fresh.md`. S9.install, S10.2 and S10.4 pass. S9.lookup is LIMIT (it needs a real signed-in Excel).

## Fix order

The score is count × severity. Severity runs from 5 down to 1: wrong number, hang, failure with no next step, failure with a next step, slow.

| # | Problem | Seen in | Score | Task |
|---|---|---|---:|---|
| 1 | The Smart View Login dialog (WebView2 → IE → broken IDCS page) blocks Excel. The engine can't see it. | S1.query.finmax, S10.9 ×2 | 12 | 12 |
| 2 | "Excel cannot access the file": the engine writes to a workbook that is open in another Excel. The same cause makes SaveAs `target="new"` fail with 0x800A03EC. This hides the dialog/edit-mode tests. | S8.dialog, S8.editmode, earlier run | 9 | 12 |
| 3 | Error codes go unnoticed in answers. Template pulls skip the member check: `actual_vs_plan` asks for `un.SY.BANORWAY` instead of `un.SY`, giving `#NoAccess`, then `#VALUE!`, then "within n/a". Nothing scans the finished sheets. | S3.actual_vs_plan.sy (found by the owner) | 5 | new (C0) |
| 4 | `xl_revert` does not restore the values. | S6.xl | 5 | new |
| 5 | The drill passes 130 s after SSO cookie loss; it succeeds on retry. | S8.sso (flaky) | 4 | 11 |
| 6 | Dict rows are rejected with a raw pydantic error. | S1.bad.dictrows | 3 | 9 |
| 7 | Slow: build 39 s, scope 45 s, drills 46–69 s, first FinMax query 18 s. | S2.build, S4.scope, S5.*, S1.query.finmax | 4 | 13 |

**Then small fixes, all known from earlier runs:**
- `xl_read` returns formats for the first cell only;
- the `xl_changes` path filter (local path vs OneDrive URL);
- S7 refresh copies are left in `output`;
- empty leftover Excels after new-workbook routes.

**Dropped (not reproduced):**
- Task 8: quoted or short dimension names are already rejected with a message naming the right dimension.
- Task 10: shape validation produced no failure.

Task 12's dialog/edit-mode naming is re-tested once #2 is fixed.

| ID | Tags | Result | Before | s | Error class |
|---|---|---|---|---:|---|
| S0.status | smoke | PASS |  | 0.5 |  |
| S1.find | lookup | PASS |  | 1.2 |  |
| S1.query.finmax | lookup | SLOW |  | 18.5 |  |
| S1.query.finlite | lookup | PASS |  | 2.6 |  |
| S1.query.finplan | lookup | PASS |  | 2.6 |  |
| S1.bad.balance | lookup, input | PASS |  | 0.1 |  |
| S1.bad.quoted | lookup, input | PASS |  | 0.1 |  |
| S1.bad.dictrows | lookup, input | FAIL |  | 0.0 | ToolFailed: sv_query: Error executing tool sv_query: # validation error for sv_queryArguments |
| S1.bad.member | lookup, input | PASS |  | 1.4 |  |
| S2.build | build | SLOW |  | 38.6 |  |
| S2.patch | build | PASS |  | 3.9 |  |
| S2.plan | build | PASS |  | 9.2 |  |
| S3.division_pnl.tec | template | PASS |  | 16.3 |  |
| S3.division_pnl.sy | template | PASS |  | 15.0 |  |
| S3.monthly_trend.tec | template | PASS |  | 15.4 |  |
| S3.monthly_trend.sy | template | PASS |  | 16.7 |  |
| S3.actual_vs_plan.tec | template | PASS |  | 20.4 |  |
| S3.actual_vs_plan.sy | template | FAIL |  | 24.1 | AssertionError: actual_vs_plan.sy: ["formula error on line 'fm_diff'"] |
| S3.cost_base.tec | template | PASS |  | 16.8 |  |
| S3.cost_base.sy | template | PASS |  | 23.6 |  |
| S4.scope | concept | SLOW |  | 44.7 |  |
| S5.drill.cold | drill | PASS |  | 59.3 |  |
| S5.drill.warm | drill | PASS |  | 46.6 |  |
| S5.drill.quarter | drill | PASS |  | 56.0 |  |
| S5.drill.report | drill | PASS |  | 69.4 |  |
| S6.xl | xl | FAIL |  | 68.7 | AssertionError: after revert differs from the snapshot in ['values'] |
| S7.refresh | refresh | PASS |  | 19.5 |  |
| S8.sso | inject | PASS (flaky) |  | 100.2 | HANG: sv_drill still running at #s |
| S8.dialog | inject | FAIL |  | 45.9 | AssertionError: sv_ws_answer (#s) does not name ('dialog',): Error executing tool sv_ws_answer: COM error during sv_ws_answer: (-#, 'Exception occurred.', (#, ' |
| S8.editmode | inject | FAIL |  | 46.9 | AssertionError: sv_ws_answer (#s) does not name ('dialog', 'edit'): Error executing tool sv_ws_answer: COM error during sv_ws_answer: (-#, 'Exception occurred.' |
| S8.parallel | inject | PASS |  | 44.7 |  |
| S8.kill.excel | inject | PASS |  | 19.7 |  |
| S8.kill.edge | inject | PASS |  | 86.0 |  |
| S8.cold | inject | PASS |  | 119.8 |  |
| S8.sleep | inject | PASS |  | 92.9 |  |
| S10.1.bundle | feature | PASS |  | 1.7 |  |
| S10.3.reload | feature | PASS |  | 10.5 |  |
| S10.5.textfixes | feature | LIMIT |  | 0.0 | Limit: user fix + override of a shipped key verified in-process (repair_text); the drill part is not tested: the # lines hit no shipped text-fix key |
| S10.6.pack | feature | LIMIT |  | 0.4 | Limit: pack has no runtime consumer; wiring only verified (pack enabled, source_product_list_#.json -> C:\dev\smartass\packs\produktregnskap\knowledge\source_pr |
| S10.7.style | feature | PASS |  | 47.5 |  |
| S10.8.coldexcel | feature | LIMIT |  | 0.1 | Limit: the user's own Excel is open (pid [#]): a cold start needs none |
| S10.9.options | feature | PASS (flaky) |  | 27.9 | AssertionError: follow-up sv_query failed: Error executing tool sv_query: Retrieve failed (rc=# Unsupported Error). Server says: H&yperion. Next step: Call sv_r |
| S10.10.http | feature | PASS |  | 0.2 |  |

## Error classes

| Count | Error class | Scenarios |
|---:|---|---|
| 1 | ToolFailed: sv_query: Error executing tool sv_query: # validation error for sv_queryArguments | S1.bad.dictrows |
| 1 | AssertionError: actual_vs_plan.sy: ["formula error on line 'fm_diff'"] | S3.actual_vs_plan.sy |
| 1 | AssertionError: after revert differs from the snapshot in ['values'] | S6.xl |
| 1 | HANG: sv_drill still running at #s | S8.sso |
| 1 | AssertionError: sv_ws_answer (#s) does not name ('dialog',): Error executing tool sv_ws_answer: COM error during sv_ws_answer: (-#, 'Exception occurred.', (#, ' | S8.dialog |
| 1 | AssertionError: sv_ws_answer (#s) does not name ('dialog', 'edit'): Error executing tool sv_ws_answer: COM error during sv_ws_answer: (-#, 'Exception occurred.' | S8.editmode |
| 1 | Limit: user fix + override of a shipped key verified in-process (repair_text); the drill part is not tested: the # lines hit no shipped text-fix key | S10.5.textfixes |
| 1 | Limit: pack has no runtime consumer; wiring only verified (pack enabled, source_product_list_#.json -> C:\dev\smartass\packs\produktregnskap\knowledge\source_pr | S10.6.pack |
| 1 | Limit: the user's own Excel is open (pid [#]): a cold start needs none | S10.8.coldexcel |
| 1 | AssertionError: follow-up sv_query failed: Error executing tool sv_query: Retrieve failed (rc=# Unsupported Error). Server says: H&yperion. Next step: Call sv_r | S10.9.options |

## Dialogs seen

| Scenario | s into scenario | PID | Title | Text | Clicked | Killed |
|---|---:|---:|---|---|---|---|
| S1.query.finmax | 13.3 | 99968 | Login |  |  |  |
| S5.drill.quarter | 41.1 | 101440 |  |  |  |  |
| S5.drill.report | 3.1 | 101440 |  |  |  |  |
| S8.dialog | 244.6 | 101440 | Microsoft Excel |  | WM_CLOSE |  |
| S8.editmode | 26.2 | 101440 | Microsoft Excel |  | WM_CLOSE |  |
| S8.editmode | 73.5 | 101440 | Microsoft Excel |  | WM_CLOSE |  |
| S10.9.options | 16.1 | 81268 | Login |  |  |  |
| S10.9.options | 38.4 | 97732 | Login |  |  |  |
