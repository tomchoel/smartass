# Benchmarks — route `core` (2026-11-02 17:33)

## Verdict (phase 6, retire + harden; engine at fe4a5af, bench fix b726dd9)

- **Answers: 13/13 PASS.** B01-B09, B13-B15 PASS in this run with every oracle cell matched; B12 hit "Smart View in
  this Excel is not signed in correctly" (see below) and PASSed on its own re-run
  (`2026-11-02-bench-core-final-B12.md`: first visible 61.97 s, total 81.62 s, 1/1 oracle).
- **Gate: 12/12**, false-go 0, needless-ask 0. The FAIL rows for B10, B11 and T01-T10 below are a bench bug: the
  oracle pass re-scored gate cases with the answer verdict (fixed in b726dd9, unit test first). Re-run with the fix:
  `2026-11-02-bench-core-final-gate.md`, 12/12 PASS.
- **Live P&L facts (D5) against the widen run** (total s): B01 74.27 vs 53.14, B13 19.06 vs 25.83, B15 28.67 vs
  33.39. B01 is the first case of the run and pays the cold start: the HTTP sign-in renewed itself (10 s) and the
  new workbook waited for a newly started Excel (about 23 s). The 36-month live read itself is 9.9 s + 6.2 s top-up.
- **Login notes in server.log during the run:** one, B12: Excel 76392 (started by this run) answered two connects
  with "Empty response from provider" and showed no Login; the note "connecting in your Excel: EPM gave an empty
  response; trying again" was sent, then the call failed fast with the NOT_SIGNED_IN advice; the user's Excel was
  not closed. No ProviderEmpty restart, no hung Login.
- **Other live checks:** `S8.kill.excel` PASS (`2026-11-02-matrix-s8-kill-final.md`: the hidden Excel killed during
  an sv_answer; the answer still came back in 9 s over HTTP; follow-up 8 s). `smartass serve` HTTP API on port
  47899: GET /health 200 (401 without the secret), POST /gate clear in 2.0 s and infeasible for "per product in
  EUR", 400 on an unknown field, POST /answer (bench B02 spec) 53.2 s cold / 7.4 s warm with 0 Critical findings,
  409 for a /gate sent during an /answer.
- **Line count** (`engine/smartview_mcp`, Python): origin/main 16 718, HEAD 16 718 (D14 gate held); owner target
  10 031, gap 6 687. Unit tests: 813 passed; ruff clean.


| ID | Result | First visible s | Total s | Oracle | Notes |
|---|---|---|---|---|---|
| B01 | PASS | 49.91 | 74.27 | 3/3 | period Oct 26 |
| B02 | PASS | 0.94 | 7.91 | 3/3 | routes {'Q1 FinMax': 'http', 'Q1 FinMax ytd': 'http'} |
| B03 | PASS | 1.44 | 5.61 | 4/4 | reused ['m', 'ytd'] |
| B04 | PASS | 1.34 | 4.42 | 3/3 | reused ['m', 'ytd', 'fp'] |
| B05 | PASS | 2.53 | 10.97 | 2/2 | period Oct 26 |
| B06 | PASS | 4.83 | 9.75 | 1/1 | period Oct 26 |
| B07 | PASS | 3.36 | 10.7 | 3/3 | period Oct 26 |
| B08 | PASS | 1.24 | 7.69 | 2/2 | reused ['tr', 'ytd'] |
| B09 | PASS | 3.62 | 9.92 | — | concept fam_mobil_own_subscriptions_mobile_consumer; period Oct 26 |
| B10 | FAIL | 0.01 | 0.01 | — | gate infeasible: EBITDA is not available per product |
| B11 | FAIL | 0.0 | 0.0 | — | gate ambiguous / asks 1 |
| B12 | ERROR | — | — | — | ToolFailed: sv_drill: Error executing tool sv_drill: Smart View in this Excel is not signed in correctly: EPM answers 'Empty response from provider' to every connect, and Smart View shows no Login to click. Ask the user to sign in once in that Excel (Smart View ribbon -> Panel -> connect FinMax, or close that Excel so the next call starts a fresh one), then ask again. If every Excel does this and  |
| B13 | PASS | 4.17 | 19.06 | 2/2 | period Oct 26 |
| B14 | PASS | 3.86 | 18.55 | 2/2 | period Dec 26 |
| B15 | PASS | 4.66 | 28.67 | 6/6 | period Oct 26 |
| T01 | FAIL | 0.01 | 0.01 | — | gate ambiguous / asks 1 |
| T02 | FAIL | 0.0 | 0.0 | — | gate ambiguous / asks 1 |
| T03 | FAIL | 0.02 | 0.02 | — | gate infeasible: EBITDA margin is not available per customer |
| T04 | FAIL | 0.0 | 0.0 | — | gate infeasible: the cubes hold no EUR values (FinMax reports NOK and local currency only) |
| T05 | FAIL | 0.02 | 0.02 | — | gate infeasible: Profit is not available per product |
| T06 | FAIL | 0.01 | 0.01 | — | gate clear |
| T07 | FAIL | 0.0 | 0.0 | — | gate ambiguous / asks 1 |
| T08 | FAIL | 0.02 | 0.02 | — | gate clear |
| T09 | FAIL | 0.0 | 0.0 | — | gate clear |
| T10 | FAIL | 0.01 | 0.01 | — | gate infeasible: actuals for Dec 27 do not exist yet (last closed month Oct 26) |

Gate: 12/12 right; false-go 0 (should have asked or refused, went ahead); needless-ask 0 (clear, but asked).
