# Live matrix 2026-10-23 (matrix-drill5)

- Time: 2026-10-23T15:37:46
- Engine: 0.2.0
- Atlas age: 9.8 days
- Options: {'only': 'drill', 'repeat': 5, 'fresh': False, 'keep': False}
- Before: 2026-10-23-matrix-baseline.json
- Totals: {'PASS': 15, 'FAIL': 5}

| ID | Tags | Result | Before | s | Error class |
|---|---|---|---|---:|---|
| S5.drill.cold #1 | drill | PASS | PASS | 84.8 |  |
| S5.drill.warm #1 | drill | PASS | PASS | 49.4 |  |
| S5.drill.quarter #1 | drill | PASS | PASS | 55.8 |  |
| S5.drill.report #1 | drill | FAIL | PASS | 0.0 | AssertionError: needs S#.build (no wb in this run: run it in the same matrix run, and it must pass) |
| S5.drill.cold #2 | drill | PASS | PASS | 67.6 |  |
| S5.drill.warm #2 | drill | PASS | PASS | 51.9 |  |
| S5.drill.quarter #2 | drill | PASS | PASS | 56.1 |  |
| S5.drill.report #2 | drill | FAIL | PASS | 0.0 | AssertionError: needs S#.build (no wb in this run: run it in the same matrix run, and it must pass) |
| S5.drill.cold #3 | drill | PASS | PASS | 63.8 |  |
| S5.drill.warm #3 | drill | PASS | PASS | 59.9 |  |
| S5.drill.quarter #3 | drill | PASS | PASS | 58.0 |  |
| S5.drill.report #3 | drill | FAIL | PASS | 0.0 | AssertionError: needs S#.build (no wb in this run: run it in the same matrix run, and it must pass) |
| S5.drill.cold #4 | drill | PASS | PASS | 68.2 |  |
| S5.drill.warm #4 | drill | PASS | PASS | 54.8 |  |
| S5.drill.quarter #4 | drill | PASS | PASS | 61.8 |  |
| S5.drill.report #4 | drill | FAIL | PASS | 0.0 | AssertionError: needs S#.build (no wb in this run: run it in the same matrix run, and it must pass) |
| S5.drill.cold #5 | drill | PASS | PASS | 64.2 |  |
| S5.drill.warm #5 | drill | PASS | PASS | 55.9 |  |
| S5.drill.quarter #5 | drill | PASS | PASS | 58.7 |  |
| S5.drill.report #5 | drill | FAIL | PASS | 0.0 | AssertionError: needs S#.build (no wb in this run: run it in the same matrix run, and it must pass) |

## Error classes

| Count | Error class | Scenarios |
|---:|---|---|
| 5 | AssertionError: needs S#.build (no wb in this run: run it in the same matrix run, and it must pass) | S5.drill.report |

## Dialogs seen

None.
