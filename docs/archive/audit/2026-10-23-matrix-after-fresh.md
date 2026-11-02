# Live matrix 2026-10-23 (matrix-after-fresh)

- Time: 2026-10-23T15:53:22
- Engine: 0.2.0
- Atlas age: 9.8 days
- Options: {'only': 'fresh', 'repeat': 1, 'fresh': True, 'keep': False}
- Before: 2026-10-23-matrix-baseline-fresh.json
- Totals: {'PASS': 3, 'LIMIT': 1}

| ID | Tags | Result | Before | s | Error class |
|---|---|---|---|---:|---|
| S9.install | fresh | PASS | PASS | 36.0 |  |
| S9.lookup | fresh | LIMIT | LIMIT | 15.6 | Limit: fresh-profile lookup needs a real signed-in Excel; COM Excel inherits the simulated env — verify manually on a new machine |
| S10.2.setup | feature, fresh | PASS | PASS | 2.6 |  |
| S10.4.update | feature, fresh | PASS | PASS | 28.1 |  |

## Error classes

| Count | Error class | Scenarios |
|---:|---|---|
| 1 | Limit: fresh-profile lookup needs a real signed-in Excel; COM Excel inherits the simulated env — verify manually on a new machine | S9.lookup |

## Dialogs seen

None.
