# Drill lines over the token vs the browser route (Task 7, the gate for deleting the browser)

Plan: `docs/plans/2026-11-19-light-fast-p0-p1.md`, Task 7.

**The run**
- Date and place: 2026-11-21, on the owner's PC.
- Target: plan1 (FinLite), read-only.
- Code: smartass 0.5.0, branch `feat/drill-token`. The run used the script as committed with this file, and
  epmhttp with the two fixes below.
- Script: `scripts/compare_drill_routes.py`, which is deleted again in Task 9.
  - It drills each cell with `drillstep.drill` twice, in a process with `SV_DRILL_ROUTE=browser` and in one with
    `SV_DRILL_ROUTE=token`, and compares the results.
  - No Excel was used: the script makes drillstep's Excel fallbacks raise, and none was asked for.
- This file holds counts, booleans, seconds, key paths, value types and query item names only. It has no figures
  and no business text.

## Verdict

- **Steps 1-3 passed.** On all four cells the token route returns the browser route's lines, row for row. The
  columns, the tie to the cube and the cube value are equal too.
- **Step 4 has not been run yet.** It draws into Excel and is run with the owner's go. The gate for deleting the
  browser route (Task 9) is passed only once Step 4 is recorded below.

## Step 1 - the view definition (cell 1)

The 19.11 spike only noted "the view definition JSON differed from the browser's (not inspected)". This is the
explanation:

```
view of ac.C60 / un.SY.2TEC / Sep 26: 218 paths on the browser route, 218 on the token route; same documents: False
  columns the drill takes from it: 45/45, same labels and fields in the same order: True
  value differs (1): $.links[*].href (a link: same path True, query items that differ: _)
```

- The two documents have the same 218 key paths, and every path has the same value type on both routes.
- The only value that differs is `links[*].href`, the self-link EPM puts in its answer. Its host and path are the
  same on both routes; only its `_` query item differs, the millisecond cache buster each request adds.
- So the difference is a volatile field.
- The diff also looks inside JSON held in strings, including `viewDefinition`, where the drill's columns are.
  Nothing differs there.
- The drill takes 45 columns from the view on each route (`drill_lines.columns_from_view`), with the same labels
  and fields in the same order.
- The token route's view is equivalent, so the run went on to Step 3.

## Step 3 - the lines

Command (repo root): `engine\.venv\Scripts\python.exe scripts\compare_drill_routes.py`. Exit code 0. The
output is shown below verbatim:

```
compare_drill_routes (smartass 0.5.0): cells [1, 2, 3, 4]; the browser route, then the token route, one process each
browser route: its process ended (exit code 0) after 62 s
token route: its process ended (exit code 0) after 46 s
Step 1 - the view definition:
  view of ac.C60 / un.SY.2TEC / Sep 26: 218 paths on the browser route, 218 on the token route; same documents: False
    columns the drill takes from it: 45/45, same labels and fields in the same order: True
    value differs (1): $.links[*].href (a link: same path True, query items that differ: _)
Step 3 - the lines:
  ac.C60 / un.SY.2TEC / Sep 26: lines 2839/2839, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 25.1/7.8
    browser: pieces 1, too wide 0, too many lines 0, derived used/tried 0/0, result warnings 1, log warnings 0, wall 25.2 s, page route 0
    token: pieces 1, too wide 0, too many lines 0, derived used/tried 0/0, result warnings 1, log warnings 0, wall 7.9 s, key requests 1, renewals 0, requests: FdmeeMain.jsp 200 x1, getData 200 x1, servicetype 200 x1, view 200 x1, refused 0
  ac.C60 / un.SY / Sep 26: lines 10413/10413, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 10.7/10.7
    browser: pieces 22, too wide 1, too many lines 0, derived used/tried 12/12, result warnings 1, log warnings 0, wall 10.9 s, page route 0
    token: pieces 22, too wide 1, too many lines 0, derived used/tried 12/12, result warnings 1, log warnings 0, wall 10.8 s, key requests 0, renewals 0, requests: FdmeeMain.jsp 200 x2, getData 200 x13, refused 0
  ac.C60080 / un.SY / Yr 26: lines 57/57, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 3.1/3.9
    browser: pieces 22, too wide 1, too many lines 0, derived used/tried 2/2, result warnings 0, log warnings 0, wall 3.1 s, page route 0
    token: pieces 22, too wide 1, too many lines 0, derived used/tried 2/2, result warnings 0, log warnings 0, wall 3.9 s, key requests 0, renewals 0, requests: FdmeeMain.jsp 200 x2, getData 200 x3, refused 0
  ac.C60 / un.SY.2TEC / Q3 26: lines 7852/7852, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 20.2/21.8
    browser: pieces 3, too wide 0, too many lines 1, derived used/tried 0/0, result warnings 1, log warnings 0, wall 20.4 s, page route 0
    token: pieces 3, too wide 0, too many lines 1, derived used/tried 0/0, result warnings 1, log warnings 0, wall 21.9 s, key requests 0, renewals 0, requests: FdmeeMain.jsp 200 x4, getData 200 x4, refused 0
  the token route started a browser: False
PASS: the token route gives the browser route's lines
```

How the script judges a cell:
- **"same lines"** means the `DrillResult.rows` are equal as sorted lists of `json.dumps(row, sort_keys=True)`
  (each hashed in its own process) and the `columns` are equal.
- **The tie** (`reconciliation["status"]`) and **the cube values** (within 0.005) are printed and judged as
  separate fields. The brief's Step 2 folds them into "same lines". The exit code is the same either way: 1 if any
  of the three differs.

Neither process read or wrote the PC's `drill_too_wide.json`. Each route therefore met EPM's too-wide refusal
itself.

### Did the cells behave as described?

- **Cell 1:** yes. It came back in one piece with 2 839 lines, as on 19.11.
- **Cell 2:** only partly.
  - It still has more than 5 000 lines (10 413) and is still split into child cells: 22 pieces, 12 of them built
    with `drill_derive` and tied to their cube values.
  - But EPM now refuses the whole cell as too wide (EPMFDM-140373), so the split comes from that refusal ("too wide
    1") and not from the line count ("too many lines 0"). It takes the same path as cell 3.
  - It was kept, and a fourth cell was added for the split by line count.
- **Cell 3:** yes. EPM refused it as too wide, it was split by entity into 22 pieces, and 2 siblings were built
  with `drill_derive`.
- **Cell 4 (added, the nearest cell to cells 1 and 2 that does what cell 2 was chosen for):**
  `ac.C60 / un.SY.2TEC / Q3 26`, the same account and entity as cell 1 over a quarter.
  - It holds over 5 000 lines in one query (7 852), so the drill query's first getData page answers with more lines
    than one piece may hold (`TooManyLines`, "too many lines 1", "too wide 0").
  - It is then split by period into 3 pieces, which are drilled one by one. Period siblings are not derived.
  - Tried for this: `Q3 26` (this behaviour) and `Yr 26` (29 947 lines; 5 splits by line count down to 12
    pieces), both on the token route only, counts only. `Q3 26` is the smaller one.

### The token route: key requests, renewals and time

- **Key requests:** the process took one session key (servicetype) and one view, on the first cell. Both carried
  over to cells 2-4.
- **The large drills:** cell 2 (10 413 lines, 13 getData calls, at most `config.DRILL_PARALLEL` = 4 in flight)
  and cell 4 (7 852 lines) made no key request and needed no renewal.
- **Renewals: 0.** The session was never lost, so the stale-key case was **not exercised live**: a request with the
  old key after a renewal, and whether that signs the session out. Tasks 5 and 6 test it offline. A long drill that
  outlasts EPM's session would show it live.
- **Paging: not exercised live.** No piece went past one getData page (`config.DRILL_PAGE` = 5 000 =
  `DRILL_SPLIT_AT`; a cell with more lines is split instead), so paging by offset over the token was not seen.
- **No browser:** the token process never created a browser (`the token route started a browser: False`). Its
  drill links came over the bearer (`svhttp`), as they do on the browser route.
- **Seconds per cell (`DrillResult.secs`, browser/token):** 25.1/7.8, 10.7/10.7, 3.1/3.9 and 20.2/21.8.
  - The browser's first cell includes starting Edge and signing in by SSO. The first, cold run of the day took
    101.5 s for that cell.
  - The token's first cell includes the session key and the view.
  - After the first cell the two routes are about equally fast: EPM's own time dominates.
  - The routes ran one after the other, browser first, so EPM had seen the same queries just before the token
    run.

## Found and fixed on the way (commit c6bfc5b, each with a test)

The first live runs of cell 1 failed on the token route. The browser route gave 2 839 lines. The token route failed
twice, for two reasons:

1. **Every request was lost (OSError, WinError 10038).**
   - urllib lets go of the connection's socket once the answer's head is in, and the answer closes the socket for
     good when its body is read to the end.
   - `epmhttp._read` then set the timeout of the closed socket for one more receive.
   - Offline tests never showed it: their fake answers have no socket, and the socket-pair test never got to the
     end of its body.
   - Fix: a closed socket's timeout is not set. Test:
     `test_epmhttp.py::test_a_whole_answer_on_a_real_socket_is_read_after_its_socket_closes` (failed before the
     fix with the same WinError).
2. **Every drill query was refused by the allowlist.**
   - The refusal was "a query item this request does not take". The unlisted item is `rcp_version` on the
     `FdmeeMain.jsp` GET.
   - Smart View's drill link carries it, and its value is digits and dots.
   - Allowing it is safe: it is a read (a GET of the drill query's JSON form) to the same origin, under a fixed name.
     It is allowed on that path only, at most once, matching `\d{1,6}(\.\d{1,6}){0,3}` (`epmhttp.VERSION`).
   - Test: `test_epmhttp.py::test_smart_views_rcp_version_is_taken_on_the_drill_query`.

Nothing else was refused in the final run (`refused 0` on every cell).

## Step 4 - the test set's drill questions on the token route (NOT RUN YET)

To be run by the controller with the owner's go, because it draws into the owner's Excel. b16 is the only test-set
case whose spec has a `lines` block. From the repo root, in PowerShell:

```
$env:SV_DRILL_ROUTE = "token"; engine\.venv\Scripts\python.exe scripts\bench_answers.py --rounds 1 --case b16
```

**Expected:**
- The run finishes with `checks.sheet_equals_model` OK. The bench prints it on each answer's progress line
  (`<!-- ... b16: <s>s ok, sheet = model OK -->`), and the warm-up answer is b16 too.
- **Evidence that the token route ran:** a misspelt variable would run the browser route and still print OK.
  Check that b16's line in the run log (`config.RUN_LOG`) has a `drills.token drill (n)` timing, and has no
  `drills.direct drill` or `drills.browser page` timing.

**Result:** _(to be added: ok runs, seconds, `checks.sheet_equals_model`, the drills.* timing keys)_
