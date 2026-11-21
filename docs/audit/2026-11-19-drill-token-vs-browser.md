# Drill lines over the token vs the browser route (Task 7, the gate for deleting the browser)

Plan: `docs/plans/2026-11-19-light-fast-p0-p1.md`, Task 7. Run on the owner's PC on 2026-11-21 against plan1
(FinLite), read-only, smartass 0.5.0, branch `feat/drill-token` at c6bfc5b (the two fixes below in place).
Script: `scripts/compare_drill_routes.py` (deleted again in Task 9). It drills each cell with `drillstep.drill`,
once in a process with `SV_DRILL_ROUTE=browser` and once in a process with `SV_DRILL_ROUTE=token`, and compares
the results. No Excel was used: the script makes drillstep's Excel fallbacks raise, and none was asked for.

This file holds counts, booleans, seconds, key paths and value types only. It has no figures and no business text.

## Verdict

- **Steps 1-3 passed.** On all three cells the token route returns the browser route's lines, row for row. The
  columns, the tie to the cube and the cube value are equal too.
- **Step 4 has not been run yet.** It draws into Excel and is run with the owner's go. The gate for deleting the
  browser route (Task 9) is passed only once Step 4 is recorded below.

## Step 1 - the view definition (cell 1)

```
view of ac.C60 / un.SY.2TEC / Sep 26: 218 paths on the browser route, 218 on the token route; same documents: False
  columns the drill takes from it: 45/45, same labels and fields in the same order: True
  value differs (1): $.links[*].href
```

- The two documents have the same 218 key paths, and every path has the same value type on both routes.
- The only difference is the value of `links[*].href`, the self-link EPM puts in its answer. It is a volatile
  field.
- The columns the drill takes from the view (`drill_lines.columns_from_view`, which reads the JSON inside
  `viewDefinition`) are the same: 45 on each route, with the same labels and fields in the same order.
- The diff walks into JSON held in strings. It also looks inside `viewDefinition`, and found nothing different
  there.
- The view difference seen in the spike does not appear here. The token route is equivalent, so the run went on
  to Step 3.

## Step 3 - the lines

Command (repo root): `engine\.venv\Scripts\python.exe scripts\compare_drill_routes.py`. Exit code 0.

```
browser route: its process ended (exit code 0) after 41 s
token route: its process ended (exit code 0) after 23 s
ac.C60 / un.SY.2TEC / Sep 26: lines 2839/2839, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 21.4/8.2
  browser: pieces 1, too wide 0, too many lines 0, derived used/tried 0/0, result warnings 1, log warnings 0, wall 21.5 s, page route 0
  token: pieces 1, too wide 0, too many lines 0, derived used/tried 0/0, result warnings 1, log warnings 0, wall 8.3 s, key requests 1, renewals 0, requests: FdmeeMain.jsp 200 x1, getData 200 x1, servicetype 200 x1, view 200 x1, refused 0
ac.C60 / un.SY / Sep 26: lines 10413/10413, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 12.6/9.8
  browser: pieces 22, too wide 1, too many lines 0, derived used/tried 12/12, result warnings 1, log warnings 0, wall 12.8 s, page route 0
  token: pieces 22, too wide 1, too many lines 0, derived used/tried 12/12, result warnings 1, log warnings 0, wall 10.0 s, key requests 0, renewals 0, requests: FdmeeMain.jsp 200 x2, getData 200 x13, refused 0
ac.C60080 / un.SY / Yr 26: lines 57/57, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 3.8/3.2
  browser: pieces 22, too wide 1, too many lines 0, derived used/tried 2/2, result warnings 0, log warnings 0, wall 3.8 s, page route 0
  token: pieces 22, too wide 1, too many lines 0, derived used/tried 2/2, result warnings 0, log warnings 0, wall 3.2 s, key requests 0, renewals 0, requests: FdmeeMain.jsp 200 x2, getData 200 x3, refused 0
the token route started a browser: False
PASS: the token route gives the browser route's lines
```

How the script judges a cell:
- **same lines:** the `DrillResult.rows` are compared as sorted lists of `json.dumps(row, sort_keys=True)`,
  each hashed in its own process, and the `columns` must be equal.
- **tie:** `reconciliation["status"]` on each route.
- **cube value equal:** the two `cell_value`s are within 0.005 of each other.

Both processes started with nothing remembered: neither read or wrote the PC's `drill_too_wide.json`. Each route
therefore met EPM's too-wide refusal itself.

### Did the cells behave as described?

- **Cell 1:** yes. It came back in one piece with 2 839 lines, as on 19.11.
- **Cell 2:** yes. It has more than 5 000 lines (10 413) and was split into child cells (22 pieces, 12 of them built
  with `drill_derive` and tied to their cube values).
  - The split came from a different place than in the brief: EPM now refuses the cell as too wide (EPMFDM-140373),
    so the line count never triggered it ("too many lines 0").
  - The cell still covers what it was chosen for: over 5 000 lines and a split. It was kept.
  - No cell here takes the split that the line count triggers (`TooManyLines`). Both routes share that path, and
    only the source of the lines differs between them.
- **Cell 3:** yes. EPM refused it as too wide, it was split by entity into 22 pieces, and 2 siblings were built
  with `drill_derive`.

### The token route: key requests, renewals and time

- **Key requests:** the process took one session key (servicetype) and one view, on the first cell.
  - Both carried over to cells 2 and 3.
  - The large cell 2 (10 413 lines, 13 getData calls, four at a time) made no key request and needed no renewal.
- **Renewals: 0.** The session was never lost, so the stale-key case was **not exercised live**: a request with the
  old key after a renewal, and whether that signs the session out. Tasks 5 and 6 test it offline. A long drill that
  outlasts EPM's session would show it live.
- **Paging: not exercised live.** No piece went past one getData page (`config.DRILL_PAGE` = 5 000 =
  `DRILL_SPLIT_AT`), so paging by offset over the token was not seen.
- **No browser:** the token process never created a browser (`the token route started a browser: False`). Its
  drill links came over the bearer (`svhttp`), as they do on the browser route.
- **Seconds per cell (`DrillResult.secs`, browser/token):** 21.4/8.2, 12.6/9.8 and 3.8/3.2.
  - The browser's first cell includes starting Edge and signing in by SSO. The first, cold run of the day took
    101.5 s for that cell.
  - The token's first cell includes the key and the view.
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

Expected: the run finishes with `checks.sheet_equals_model` OK. The bench prints it on each answer's progress line
(`<!-- ... b16: <s>s ok, sheet = model OK -->`), and the warm-up answer is b16 too.

Result: _(to be added: ok runs, seconds, `checks.sheet_equals_model`)_
