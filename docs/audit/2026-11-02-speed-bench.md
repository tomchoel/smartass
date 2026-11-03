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
