# Benchmarks — route `core` (2026-11-02 14:39)

| ID | Result | First visible s | Total s | Oracle | Notes |
|---|---|---|---|---|---|
| B01 | PASS | 27.74 | 53.14 | 3/3 | period Oct 26 |
| B02 | PASS | 0.92 | 9.06 | 3/3 | routes {'Q1 FinMax': 'http', 'Q1 FinMax ytd': 'http'} |
| B03 | PASS | 0.81 | 5.69 | 4/4 | reused ['m', 'ytd'] |
| B04 | PASS | 1.53 | 4.97 | 3/3 | reused ['m', 'ytd', 'fp'] |
| B05 | PASS | 2.66 | 11.56 | 2/2 | period Oct 26 |
| B06 | PASS | 4.34 | 9.48 | 1/1 | period Oct 26 |
| B07 | PASS | 3.83 | 10.8 | 3/3 | period Oct 26 |
| B08 | PASS | 1.38 | 7.02 | 2/2 | reused ['tr', 'ytd'] |
| B09 | PASS | 4.03 | 9.88 | — | concept fam_mobil_own_subscriptions_mobile_consumer; period Oct 26 |
| B12 | PASS | 61.59 | 68.39 | 1/1 | None |
| B13 | PASS | 5.38 | 25.83 | 2/2 | period Oct 26 |
| B14 | PASS | 3.17 | 18.47 | 2/2 | period Dec 26 |
| B15 | PASS | 5.39 | 33.39 | 6/6 | period Oct 26 |

## Fixes found live

- B04: `save_spec` overwrote the wrong `_svmeta` row (first meta row is 2); fixed in 1129fd0.
- B05/B14: the FinPlan-vs-FinMax EBITDA recon (~0.86 MNOK grouping difference) was graded Critical; it is now a Warning (`Check.severity`, 96bcee1).
- A first run hit a transient Smart View sign-in failure on B12; it cleared without action and this run is 13/13.
- Gate cases: `docs/audit/2026-11-02-bench-core-gate-p5.md` (12/12).
