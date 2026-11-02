# Benchmarks — route `core` (2026-11-02 17:47)

| ID | Result | First visible s | Total s | Oracle | Notes |
|---|---|---|---|---|---|
| B10 | PASS | 0.03 | 0.03 | — | gate infeasible: EBITDA is not available per product |
| B11 | PASS | 0.02 | 0.02 | — | gate ambiguous / asks 1 |
| T01 | PASS | 0.01 | 0.01 | — | gate ambiguous / asks 1 |
| T02 | PASS | 0.02 | 0.02 | — | gate ambiguous / asks 1 |
| T03 | PASS | 0.0 | 0.0 | — | gate infeasible: EBITDA margin is not available per customer |
| T04 | PASS | 0.02 | 0.02 | — | gate infeasible: the cubes hold no EUR values (FinMax reports NOK and local currency only) |
| T05 | PASS | 0.01 | 0.01 | — | gate infeasible: Profit is not available per product |
| T06 | PASS | 0.02 | 0.02 | — | gate clear |
| T07 | PASS | 0.0 | 0.0 | — | gate ambiguous / asks 1 |
| T08 | PASS | 0.01 | 0.01 | — | gate clear |
| T09 | PASS | 0.02 | 0.02 | — | gate clear |
| T10 | PASS | 0.02 | 0.02 | — | gate infeasible: actuals for Dec 27 do not exist yet (last closed month Oct 26) |

Gate: 12/12 right; false-go 0 (should have asked or refused, went ahead); needless-ask 0 (clear, but asked).
