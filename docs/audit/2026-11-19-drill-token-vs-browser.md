# Drill lines over the token vs the browser route (Task 7, the gate for deleting the browser)

Plan: `docs/plans/2026-11-19-light-fast-p0-p1.md`, Task 7.

**The run**
- Date and place: 2026-11-21, on the owner's PC.
- Target: plan1 (FinLite), read-only.
- Code: smartass 0.5.0, branch `feat/drill-token`. The runs used the script as committed with this file, and
  epmhttp with the two fixes below.
- Script: `scripts/compare_drill_routes.py`, deleted again in Task 9.
  - It drills cells with `drillstep.drill` in two processes, one per route. By default these are
    `SV_DRILL_ROUTE=browser` and `SV_DRILL_ROUTE=token`; `--routes token,token` compares the token route with
    itself. The script compares the results.
  - `--split off` drills without splitting, so a large piece is paged.
  - `--probe renewal` runs the forced-renewal probe on the token route.
  - No Excel was used: the script makes drillstep's Excel fallbacks raise, and none was asked for.
- This file holds counts, booleans, seconds, key paths, value types and query item names only. It has no figures
  and no business text.

## Verdict: the gate for Task 9

Task 9 may delete the browser stack only when **all three** of these have passed and are recorded in this file:

| | What | Status |
|---|---|---|
| (a) | Step 4: b16 on the token route, finishing with `checks.sheet_equals_model` OK. Its run-log line must have a `drills.token drill` timing, and no `drills.direct drill` or `drills.browser page` timing. | **NOT RUN**: the controller runs it with the owner's go (see below) |
| (b) | Paging probe: cell 4 drilled without a split, so the token route pages (getData at offsets 0 and 5 000), first against itself, then against the browser route. | **PASSED** 2026-11-21 |
| (c) | Forced-renewal probe: the session lost while cell 2's getData requests are queued and in flight, then a later drill in the same process. | **PASSED** 2026-11-21 |

Steps 1-3 also passed: on all four cells the token route returns the browser route's lines, row for row. The
columns, the value types, the tie to the cube and the cube value are equal too.

## Step 1 - the view definition (cell 1)

The 19.11 spike only noted "the view definition JSON differed from the browser's (not inspected)". This is the
explanation, from the run in Step 3:

```
Step 1 - the view definition (browser/token):
  view of ac.C60 / un.SY.2TEC / Sep 26: 218 paths on one route, 218 on the other; same documents: False
    columns the drill takes from it: 45/45, same labels and fields in the same order: True
    value differs (1): $.links[*].href (a link: same path True, query items that differ: _)
```

- The two documents have the same 218 key paths, and every path has the same value type on both routes.
- The only value that differs is `links[*].href`, the self-link EPM puts in its answer. Its host and path are the
  same on both routes; only its `_` query item differs, the millisecond cache buster each request adds.
- So the difference is a volatile field.
- The diff also looks inside JSON held in strings, including `viewDefinition`, where the drill's columns are.
  Nothing differs there.
- The drill takes 45 columns from the view on each route, with the same labels and fields in the same order.
- The view crosses between the processes as a digest only: key paths, types, and hashes of the leaf values and of
  the columns. A view the drill cannot read fails the run.

## Step 3 - the lines

Command (repo root): `engine\.venv\Scripts\python.exe scripts\compare_drill_routes.py`. Exit code 0. The
output is shown below verbatim:

```
compare_drill_routes (smartass 0.5.0): cells [1, 2, 3, 4], split 'auto'; routes browser then token, one process each
browser route: its process ended (exit code 0) after 66 s
token route: its process ended (exit code 0) after 50 s
Step 1 - the view definition (browser/token):
  view of ac.C60 / un.SY.2TEC / Sep 26: 218 paths on one route, 218 on the other; same documents: False
    columns the drill takes from it: 45/45, same labels and fields in the same order: True
    value differs (1): $.links[*].href (a link: same path True, query items that differ: _)
Step 3 - the lines:
  ac.C60 / un.SY.2TEC / Sep 26: lines 2839/2839, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 26.1/8.4
    browser: pieces 1, too wide 0, too many lines 0, derived used/tried 0/0, result warnings 1, log warnings 0, wall 26.2 s, page route 0
    token: pieces 1, too wide 0, too many lines 0, derived used/tried 0/0, result warnings 1, log warnings 0, wall 8.5 s, key requests 1, renewals 0, requests: FdmeeMain.jsp 200 x1, getData 200 x1, servicetype 200 x1, view 200 x1, getData offsets: 0, refused 0
  ac.C60 / un.SY / Sep 26: lines 10413/10413, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 10.6/10.8
    browser: pieces 22, too wide 1, too many lines 0, derived used/tried 12/12, result warnings 1, log warnings 0, wall 10.8 s, page route 0
    token: pieces 22, too wide 1, too many lines 0, derived used/tried 12/12, result warnings 1, log warnings 0, wall 11.1 s, key requests 0, renewals 0, requests: FdmeeMain.jsp 200 x2, getData 200 x13, getData offsets: 0 x13, refused 0
  ac.C60080 / un.SY / Yr 26: lines 57/57, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 3.8/4.4
    browser: pieces 22, too wide 1, too many lines 0, derived used/tried 2/2, result warnings 0, log warnings 0, wall 3.8 s, page route 0
    token: pieces 22, too wide 1, too many lines 0, derived used/tried 2/2, result warnings 0, log warnings 0, wall 4.4 s, key requests 0, renewals 0, requests: FdmeeMain.jsp 200 x2, getData 200 x3, getData offsets: 0 x3, refused 0
  ac.C60 / un.SY.2TEC / Q3 26: lines 7852/7852, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 22.2/23.2
    browser: pieces 3, too wide 0, too many lines 1, derived used/tried 0/0, result warnings 1, log warnings 0, wall 22.5 s, page route 0
    token: pieces 3, too wide 0, too many lines 1, derived used/tried 0/0, result warnings 1, log warnings 0, wall 23.4 s, key requests 0, renewals 0, requests: FdmeeMain.jsp 200 x4, getData 200 x4, getData offsets: 0 x4, refused 0
  the token route started a browser: False
PASS: the routes give the same lines
```

How the script judges a cell:
- **"same lines"** means two things:
  - the `DrillResult.rows` are equal as sorted lists of `json.dumps(row, sort_keys=True)`, each hashed in its
    own process. A value JSON has no type for keeps its type name, so a Decimal never equals a string. The
    per-column type names are compared too, and a difference is named by its column.
  - the `columns` are equal.
- **The tie** (`reconciliation["status"]`) and **the cube values** (within 0.005) are printed and judged as
  separate fields. The brief's Step 2 folds them into "same lines". The exit code is the same either way: 1 if any
  of them differs.

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
  - It is then split by period into 3 pieces, which are drilled one by one. Period siblings are not derived:
    `drill_derive.COLUMNS` covers entities and accounts only.
  - Tried for this: `Q3 26` (this behaviour) and `Yr 26` (29 947 lines; 5 splits by line count down to 12
    pieces), both on the token route only, counts only. `Q3 26` is the smaller one.

### Time and key requests in Step 3

- **Key requests:** the token process took one session key (servicetype) and one view, on the first cell. Both
  carried over to cells 2-4.
- **The large drills:** cell 2 (10 413 lines, 13 getData calls, at most `config.DRILL_PARALLEL` = 4 in flight) and
  cell 4 (7 852 lines) made no key request and needed no renewal.
- **No browser:** the token process never created a browser. Its drill links came over the bearer (`svhttp`), as
  they do on the browser route.
- **Seconds per cell:**
  - On the first cell the token route is faster. The browser's first cell includes starting Edge and signing in by
    SSO; the first, cold run of the day took 101.5 s for it.
  - After the first cell the two routes are about equally fast: EPM's own time dominates.
  - The routes ran one after the other, browser first.

## (b) The paging probe - PASSED

A piece is paged by offset in production when `TooManyLines` meets a piece that cannot be split (a base-level
leaf, or depth 4; `drillstep._drill_piece`, "leaf: page it all"). With any split other than "auto", every piece is
paged.

The probe drills cell 4 with `--split off`, so its 7 852 lines come as one piece over two getData pages
(`config.DRILL_PAGE` = 5 000).

EPM's offset pages have no stable order (`config.py`, `DRILL_PAGE`), so the token route was first compared with
itself, and only then with the browser route. The rows are compared as multisets: a row paged twice or missing
shows as "only on" or "paged twice" counts.

Token route against itself: `compare_drill_routes.py --cell 4 --split off --routes token,token`, exit code 0.

```
compare_drill_routes (smartass 0.5.0): cells [4], split 'off'; routes token then token, one process each
token route: its process ended (exit code 0) after 27 s
token route: its process ended (exit code 0) after 26 s
Step 1 - the view definition (token/token):
  view of ac.C60 / un.SY.2TEC / Q3 26: 218 paths on one route, 218 on the other; same documents: False
    columns the drill takes from it: 45/45, same labels and fields in the same order: True
    value differs (1): $.links[*].href (a link: same path True, query items that differ: _)
Step 3 - the lines:
  ac.C60 / un.SY.2TEC / Q3 26: lines 7852/7852, same lines: True, tie token/token: exact/exact, cube value equal: True, seconds 22.8/23.1
    token: pieces 1, too wide 0, too many lines 0, derived used/tried 0/0, result warnings 1, log warnings 0, wall 23.1 s, key requests 1, renewals 0, requests: FdmeeMain.jsp 200 x1, getData 200 x2, servicetype 200 x1, view 200 x1, getData offsets: 0, 5000, refused 0
    token: pieces 1, too wide 0, too many lines 0, derived used/tried 0/0, result warnings 1, log warnings 0, wall 23.3 s, key requests 1, renewals 0, requests: FdmeeMain.jsp 200 x1, getData 200 x2, servicetype 200 x1, view 200 x1, getData offsets: 0, 5000, refused 0
  the token route started a browser: False
  the token route started a browser: False
PASS: the routes give the same lines
```

Browser route against the token route: `compare_drill_routes.py --cell 4 --split off`, exit code 0.

```
compare_drill_routes (smartass 0.5.0): cells [4], split 'off'; routes browser then token, one process each
browser route: its process ended (exit code 0) after 58 s
token route: its process ended (exit code 0) after 27 s
Step 1 - the view definition (browser/token):
  view of ac.C60 / un.SY.2TEC / Q3 26: 218 paths on one route, 218 on the other; same documents: False
    columns the drill takes from it: 45/45, same labels and fields in the same order: True
    value differs (1): $.links[*].href (a link: same path True, query items that differ: _)
Step 3 - the lines:
  ac.C60 / un.SY.2TEC / Q3 26: lines 7852/7852, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds 48.6/22.8
    browser: pieces 1, too wide 0, too many lines 0, derived used/tried 0/0, result warnings 1, log warnings 0, wall 48.9 s, page route 0
    token: pieces 1, too wide 0, too many lines 0, derived used/tried 0/0, result warnings 1, log warnings 0, wall 23.0 s, key requests 1, renewals 0, requests: FdmeeMain.jsp 200 x1, getData 200 x2, servicetype 200 x1, view 200 x1, getData offsets: 0, 5000, refused 0
  the token route started a browser: False
PASS: the routes give the same lines
```

- The token route paged at offsets 0 and 5 000 and returned 7 852 lines with an exact tie and an equal cube value.
- What I saw: the rows were the same in both comparisons, so no row was missing or paged twice. Two token runs and
  the browser run all returned the same set of rows.
- On this cell, on this day, the order of EPM's pages did not change the lines. The same comparison held on an
  earlier run of both commands that same afternoon.
- `split` is not reachable from a tool call (fin_drill and the recipes use "auto"). The probe uses it only to
  reach the paging that a large leaf piece gets in production.

## (c) The forced-renewal probe - PASSED

Command: `compare_drill_routes.py --probe renewal`, exit code 0. It is one token-route process.

**How the session is lost.** Nothing is written to EPM.
- The probe drills cell 1 and then cell 2 (the baseline).
- It then drills cell 2 again. Just before that drill's second getData is sent, the session's cookie jar is
  emptied, so the server session is gone from that request on. That getData is the first of the 12 derived siblings
  that run at once, at most 4 getData in flight.
- EPM gets the old key without its session cookie and answers it, and the requests that follow it with the old key,
  with the redirect that `drillcore.expired` treats as "session lost". This is what happens when EPM's server
  session ends.
- So the drill goes through `TokenDrill._renew` with the old key in flight, not through a plain key request.

**Then, to answer Task 5's open question:**
- One getData is sent with the old key on the new session. It is a read: EPM's own query, `limit=1`.
- Cell 1 is drilled again in the same process.

```
compare_drill_routes (smartass 0.5.0): probe renewal
token route: its process ended (exit code 0) after 40 s
Renewal probe (token route, one process; cell 2's session cookies dropped before its getData #2):
  cell 1 (ac.C60 / un.SY.2TEC / Sep 26): lines 2839, tie exact, key requests 1, renewals 0, wall 8.4 s
  cell 2 (ac.C60 / un.SY / Sep 26), baseline: lines 10413, tie exact, key requests 0, renewals 0, wall 11.2 s
  cell 2 with the session lost (cookies dropped: True): lines 10413, tie exact, key requests 1, renewals 4, wall 10.0 s; same rows as the baseline: True, cube value equal: True
    requests: FdmeeMain.jsp 200 x1, getData 200 x13, getData 302 x4, servicetype 200 x1
    old-key requests: sent before the renewal began 6, sent between its start and the new key 0, sent after the new key 0; answered after the new key 0
  an old-key getData sent on the new session: HTTP 302
  cell 1 again: lines 2839, tie exact, key requests 1, renewals 1, wall 7.9 s; same rows as before: True -> the old-key request signed the new session out
  the token route started a browser: False
PASS: the drill renewed its key under load and kept working
```

- **Renewal under load:** 4 getData requests were answered 302 (session lost). The pieces called `_renew` 4
  times, but EPM was asked for a key once ("key requests 1"): the renewal was shared.
- **The lines:** the 10 413 lines are the baseline's, row for row, with an exact tie and an equal cube value.
- **The old key after the renewal:**
  - 6 requests carried the old key, and all 6 left before the renewal began: 2 answered 200 before the cookies
    were dropped, and 4 were in flight and lost.
  - None left between the renewal's start and the new key, or after the new key: TokenDrill reads the key inside
    the getData gate.
  - Every old-key answer had come back before the new key was taken ("answered after the new key 0").
- **A later drill in the same process works:** cell 1 again, with the same rows as the first time.
- Four runs of the probe gave the same counts.

**Task 5's open question: does a request with an old key sign the new session out? Yes.**
- The old-key getData on the new session was answered 302. The next drill (cell 1) then found its key gone and had
  to renew once ("renewals 1, key requests 1"), and then it worked.
- A control run made sure the old-key request caused this. It ran the same steps without the old-key request, as a
  scratchpad diagnostic using this script's own classes, and the later drill needed no renewal ("renewals 0, key
  requests 0").
- Our cookies did not change: the cookie jar is the same before and after the old-key request. So it is EPM that
  ends the server session (or its key) when an old key arrives on it.
- Consequence for production: a request with the old key that reaches EPM after the new key was taken would sign
  the new session out.
  - The drill would then renew again. A piece that has already used its one recover would fail with "getData
    returned ... at offset N".
  - In the probe no old-key request was still at EPM when the new key was taken, so this did not happen.
  - TokenDrill takes the new key as soon as a request finds the session lost. It does not wait for the requests
    still in flight with the old key. That race stays possible on a slow network.

## Found and fixed on the way (commit c6bfc5b, each with a test)

The first live runs of cell 1 failed on the token route. The browser route gave 2 839 lines. The token route failed
twice, for two reasons:

1. **Every request was lost (OSError, WinError 10038).**
   - urllib lets go of the connection's socket once the answer's head is in, and the answer closes the socket for
     good when its body is read to the end.
   - `epmhttp._read` then set the timeout of the closed socket for one more receive.
   - Fix: a closed socket's timeout is not set. Test:
     `test_epmhttp.py::test_a_whole_answer_on_a_real_socket_is_read_after_its_socket_closes`.
2. **Every drill query was refused by the allowlist.**
   - The refusal was "a query item this request does not take". The unlisted item is `rcp_version` on the
     `FdmeeMain.jsp` GET.
   - Smart View's drill link carries it, and its value is digits and dots.
   - Allowing it is safe: it is a read to the same origin, under a fixed name. It is allowed on that path only, at
     most once, matching `\d{1,6}(\.\d{1,6}){0,3}` (`epmhttp.VERSION`).
   - Test: `test_epmhttp.py::test_smart_views_rcp_version_is_taken_on_the_drill_query`.

Nothing else was refused in any later run (`refused 0` everywhere).

## (a) Step 4 - the test set's drill questions on the token route (NOT RUN YET)

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
