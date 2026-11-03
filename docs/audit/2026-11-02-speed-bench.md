# Speed bench (2026-11-02)

Times only, no figures. Question: consultants (lines_ranking), un.SY, C60080 + C60090, Yr 26 and Yr 25.

## Step 1 (drills): drill-only bench, `_drill_cell` x 4 cells at once (as lines_ranking runs them)
Script: session scratchpad `bench_drill.py` (any Excel call fails loudly - none happened).

| | Before (handoff, cold) | After step 1 |
|---|---:|---:|
| Excel during the drill | drill links ~0.8 s each, sequential, + cold hidden Excel | none |
| Drill pages | one per piece, ~26 s each in parallel tabs | one warm page per session |
| Empty child / too-wide cell | 60 s timeout / ~20 s refusal | immediate |
| Drills, both years (wall) | ~65 s (2026) with 2025 in parallel | **52 s** |
| Tie to the cube | gap 0 | gap 0 on all 4 cells |

Stage sums after step 1: browser start 1.1 s, sign-in 4.5 s, `http drill link` x30 = 652 s summed (~22 s each),
`direct drill` x26 = 66 s summed (~2.5 s each). Pieces: 22 per cell (12-19 of them empty).

## Why it is not faster yet
Everything runs on the drill browser's ONE worker thread (Playwright sync API): the HTTP route (`svhttp`
BrowserTransport) and every direct drill queue behind each other. A drill link costs ~0.05 s of server time but
waits ~22 s in the queue; the 26 direct pieces run strictly one after another (~66 s of work).

## Next (step 1b)
- Run direct pieces concurrently inside the browser (async fetches in the warm tab, results collected), so the
  worker thread is not blocked for each request.
- Fewer requests per split: child drill links are the parent's URL with one member changed - or at least ask for
  the links before the drills start, when the queue is empty.

## Step 1b (2026-11-03): everything in flight at once, one EPM drill query per split
Same four cells. Changes: requests start in the tab and are collected later (the worker never waits on EPM);
a pool of Smart View sessions (parallel requests on ONE sID came back corrupted: 2-3 of 8 grids wrong per round;
on their own sIDs 8 of 8 right); async direct drill; split siblings' drill queries built from one of EPM's
(drill_derive), each accepted only when its lines tie to the cube; the warm tab loads while the queries run.

| | Step 1 | Step 1b |
|---|---:|---:|
| Drills, both years, browser cold (start + sign-in ~6 s) | 52 s | **15 s** |
| Drills incl. normalise, browser already signed in | - | **9.2 s** |
| Same, second time in the process (warm tab) | - | **6.1 s** |
| Drill links / EPM drill queries / built pieces | 30 / 26 / 0 | 8 / 4 / 22 (all 22 tied) |
| Tie to the cube | gap 0 | exact on all 4 cells, 0 warnings |
| 12 grids through the HTTP route at once | one at a time | 0.2-0.3 s, 12/12 equal to one at a time (3 rounds) |

What is left: EPM's drill query (one at a time per user, ~1.5-3 s on these year cells) and the warm page (~4 s,
once per browser session). Scripts: session scratchpad `bench_drill.py`, `bench_timeline.py`, `stability.py`,
`spike_sid.py`.

## Answer speed (2026-11-03): pulls and pnl reads in parallel, prefetched while Excel opens
Live runs on the owner's PC through the MCP server (relay-reloaded on this code), one new workbook. Times and
COM-call counts only (`runs.jsonl`).

| Answer | Before (mean of the last 40 runs, by stage) | After |
|---|---|---|
| P&L Wholesale YTD (`pnl`) | `live read` 14 s mean, 76 s worst | `live read` **7.0 s** (5 pulls over HTTP) |
| Consultants Technology YTD (`lines_ranking`) | ~25 s Excel + drills | first frame **0.8 s**, drills **8.0 s**, Excel job **10.8 s**: ~20 s end to end |
| Revenue trend Telenor Norge (`monthly_trend`) | `fetch (HTTP)` up to 8 s | `fetch (HTTP)` **0.08 s** (3 pulls prefetched), job **12.1 s** |

Login spike: a fresh Excel shows the EPM login by itself after 19-22 s, so the 12 s grace stays; the P&L run that
had to start Excel spent **25.7 s** in `open target`.

### Where the Excel time goes now (input for the Excel-write plan)
| Stage | Time | COM calls |
|---|---:|---:|
| `write source sheet` | 2.9-4.6 s | ~78 per sheet (234 for 3, 401 for 5) |
| `answer` (mostly `style (tn-rapportstil)`; `G page setup` 1.3-1.4 s in 16 calls) | 2.9-4.2 s | 158-286 |
| `save + activate` (`arrange sheets`) | 0.9-1.3 s | 84-104 |
| `write lines d1` / `d2` | 1.4-1.7 s each | 92 each |
| `log sources` | 0.2-1.4 s | 15-78 |
| `open target` when Excel must start | 25.7 s | 3 |

Top three to cut next: source-sheet styling (~78 calls per sheet), sheet arranging on save (~100 calls), the
answer styling + page setup.
