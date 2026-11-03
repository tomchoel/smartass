# Benchmarks — route `core` (2026-11-02 20:54)

| ID | Result | First visible s | Total s | Oracle | Notes |
|---|---|---|---|---|---|
| B01 | PASS | 3.45 | 94.62 | 3/3 | bench pins recipe pnl (gate named actual_vs_plan); period Oct 26 |
| B02 | ERROR | — | — | — | ToolFailed: sv_answer: Error executing tool sv_answer: sv_answer failed - AttributeError: 'str' object has no attribute 'get'. Next step: Retry once; if it repeats, call sv_status (in a terminal: `smartass doctor`) and send the file from `smartass doctor --bundle` to the maintainer. |
| B03 | SKIP | — | — | — | needs B02 to pass in the same run |
| B04 | SKIP | — | — | — | needs B02 to pass in the same run |
| B05 | PASS | 3.95 | 14.25 | 2/2 | period Oct 26 |
| B06 | PASS | 2.91 | 14.06 | 1/1 | bench pins recipe cost_base (gate named pnl); period Oct 26 |
| B07 | PASS | 2.98 | 15.5 | 3/3 | period Oct 26 |
| B08 | PASS | 2.73 | 8.44 | 2/2 | reused ['tr', 'ytd', 'ly'] |
| B09 | PASS | 3.36 | 12.66 | — | concept fam_mobil_own_subscriptions_mobile_consumer; period Oct 26 |
| B10 | PASS | 0.02 | 0.02 | — | gate infeasible: EBITDA is not available per product |
| B11 | PASS | 0.0 | 0.0 | — | gate ambiguous / asks 1 |
| B12 | ERROR | — | — | — | ToolFailed: sv_drill: Error executing tool sv_drill: Smart View in this Excel is not signed in correctly: EPM answers 'Empty response from provider' to every connect, and Smart View shows no Login to click. Ask the user to sign in once in that Excel (Smart View ribbon -> Panel -> connect FinMax, or close that Excel so the next call starts a fresh one), then ask again. If every Excel does this and  |
| B13 | PASS | 3.45 | 20.33 | 2/2 | period Oct 26 |
| B14 | PASS | 3.69 | 17.09 | 2/2 | period Dec 26 |
| B15 | PASS | 3.72 | 25.7 | 6/6 | bench pins recipe pnl (gate named actual_vs_plan); period Oct 26 |
| B16 | FAIL | 5.23 | 81.95 | 1/1 | first frame 5.23 s > 2.5 s |
| T01 | PASS | 0.02 | 0.02 | — | gate ambiguous / asks 1 |
| T02 | PASS | 0.01 | 0.01 | — | gate ambiguous / asks 1 |
| T03 | PASS | 0.02 | 0.02 | — | gate infeasible: EBITDA margin is not available per customer |
| T04 | PASS | 0.0 | 0.0 | — | gate infeasible: the cubes hold no EUR values (FinMax reports NOK and local currency only) |
| T05 | PASS | 0.01 | 0.01 | — | gate infeasible: Profit is not available per product |
| T06 | PASS | 0.0 | 0.0 | — | gate clear |
| T07 | PASS | 0.02 | 0.02 | — | gate ambiguous / asks 1 |
| T08 | PASS | 0.0 | 0.0 | — | gate clear |
| T09 | PASS | 0.02 | 0.02 | — | gate clear |
| T10 | PASS | 0.0 | 0.0 | — | gate infeasible: actuals for Dec 27 do not exist yet (last closed month Oct 26) |
| T11 | PASS | 0.01 | 0.01 | — | gate clear |
| T12 | PASS | 0.0 | 0.0 | — | gate clear |

Gate: 14/14 right; false-go 0 (should have asked or refused, went ahead); needless-ask 0 (clear, but asked).
