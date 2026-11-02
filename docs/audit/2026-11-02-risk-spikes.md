# Risk spikes R1–R4 (2026-11-02)

Go/no-go spikes from [`specs/2026-11-02-pipeline-architecture-design.md`](../specs/2026-11-02-pipeline-architecture-design.md) §7.
Scripts are in `engine/tests/live/spike_r*.py`; raw results are in the JSON files next to this one. No figures are stored.

## Summary

| Spike | Question | Result | Verdict |
|---|---|---|---|
| R1 | Can a source sheet be a grid of HsGetValue formulas in the user's Excel? | 490 cells written in 1.0 s, refreshed in 1.3 s (0.7 s warm), equal to HTTP in every cell, values survive save and reopen | **Go, with a condition:** only when Excel's own Smart View is signed in and connected (see findings) |
| R2 | Can the Smart View HTTP provider fetch B15 fast, in parallel, after idle? | B15 (4 pulls, 174 cells) warm in 0.4 s; a 1 610-cell grid in 0.19 s; parallel = sequential; cold start 13–26 s (browser + SSO) | **Go** (idle result: see below) |
| R3 | Is a one-block write into the open workbook safe while the user works? | 630 cells in 0.02–0.36 s with the user's book active; waits ~7 s while Excel is busy or a cell is edited, then lands; user's book stays active and untouched | **Go** |
| R4 | Is there a supported alternative to the undocumented provider? | Official REST needs an OAuth client registered by an IDCS admin; the SSO session gets 401; `exportdataslice` exports only read-write cells | **Keep svhttp**; ask the EPM admin for an OAuth client as the supported fallback |

## R1 — HsGetValue source grid

Setup: 7 business areas × 10 P&L lines × 7 months (490 cells), FinMax, Actual, NOK, Periodic. Each formula references
its row/column labels (`"All Entities#"&$A5`, …) and names the other 11 dimensions explicitly.

| Step | Result |
|---|---|
| Write 490 formulas (one `Formula2` block) | 1.0 s |
| Shown right after the write | `#NEED_REFRESH` in every cell (no connection yet: `#No Connection`) |
| Smart View Refresh without a connection | Returns OK but changes nothing: every cell stays `#No Connection` |
| Connect FinMax on a helper sheet (`HypUIConnect`), then Refresh | 1.3 s; 490 numbers, **0 cells differ from HTTP** |
| Second Refresh | 0.7 s |
| Save, close, reopen without Refresh | 490 numbers, 0 differ (Excel keeps the last values) |

Findings:
1. **Refresh does not connect.** HsGetValue resolves only after something connects the private connection `FinMax`
   in that Excel. The engine can do it (`B_Connect` on a helper sheet), so no manual step — when the login works.
2. **The login is the weak point.** In this session, one Excel answered every connect with "Empty response from
   provider" (the known stale-session case), and another showed a Smart View Login window that did not react to
   clicks. Both had to be killed. The HTTP route (its own Edge SSO) had no such problem in any run today.
3. **Values cannot be pre-filled in live Excel.** A formula entered through COM is evaluated at once and shows
   `#NEED_REFRESH`; there is no way to give it the HTTP value as its cached result. (Only a closed-file writer can do
   that — the reason the headless spike used openpyxl.) So a formula grid shows numbers only after a successful
   Smart View Refresh.
4. Bridge: added `B_RefreshSheet` (Smart View Refresh on one sheet, read-only like the Refresh button).

## R2 — Smart View HTTP provider

| Measure | Result |
|---|---|
| B15 pulls (pnl 55, fpa 63, fpp 35, fte 21 cells) warm, sequential | 0.38–0.44 s total (0.06–0.14 s per pull) |
| Stress grid 7 BA × 10 lines × 21 months = 1 610 cells | 0.17–0.19 s |
| Same five pulls on five threads | 0.28–0.49 s, every cell equal to sequential (the client serialises on one lock) |
| Cold (hidden Edge start + SSO + three cube sessions) | 12.8 s and 26.0 s in two runs |
| Fetch after 1 h idle | _pending — `2026-11-02-spike-r2-idle.json`_ |

## R3 — Live write

Block of 70 × 9 synthetic numbers written with `Range.Value2` into the session workbook, never activating it. The
"user's" workbook is a second book in the same Excel and is the active one throughout.

| Case | Write time | Landed | User's book active / untouched |
|---|---|---|---|
| Baseline | 0.36 s | yes | yes / yes |
| User's book active | 0.02 s | yes | yes / yes |
| Excel busy (8 s VBA loop) | 7.0 s (waited) | yes | yes / yes |
| Cell in edit mode, Esc after 6 s | 7.1 s (waited) | yes | yes / yes |

The existing `com()` retry (60 s) covers busy and edit mode. Longer than that, the tool reports "press Enter or Esc".

## R4 — Supported path

From the Stage 0 probe (`specs/2026-10-25-stage0-epm-access-probe.md`) and `DESIGN.md`: Planning REST and
`/interop/rest` reject the SSO session (401, bearer tokens only); a client needs one-time IDCS admin registration plus
a device-code login (refresh token 7 days per pod); `exportdataslice` exports only cells the user may write. The
provider protocol is undocumented and may change with a monthly update. Mitigation: a daily canary query in the live
checks, the COM route as fallback, and a request to the EPM admin for an OAuth client so REST can replace the provider
if it ever breaks.

## Consequence for the design

The original decision was "source sheets = HsGetValue grids pre-filled from HTTP". R1 shows the pre-fill is not
possible in live Excel and that a formula grid puts Excel's Smart View login — today's least reliable part — on the
critical path of every answer.

**Decision (owner, 2026-11-02):** source sheets are HTTP value grids with a provenance header and the stored pull
spec (form C) by default; "make Smart View-refreshable" converts a sheet to HsGetValue formulas (form B) on demand
when Excel's Smart View is signed in; ad-hoc grids (form A) only as fallback. Excel login handling is hardened
(fast detection of a stuck Login window, restart only the engine's own Excel). Recorded in the design spec §3.2.
