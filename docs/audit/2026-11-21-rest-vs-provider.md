# The second engine against the provider route (phase 2, Task 3 and 3b), 2026-11-21/22

Live, read-only against EPM (plan1 and stratmodel), without Excel, from the owner's PC with the stored refresh tokens
(no sign-in was needed). Script: `scripts/compare_grid_routes.py all` (deleted after the phase's final gate run,
22.11.2026 at 8049f2a: see "Final run"), run from `engine/` on the tree of the commit that adds this version of the file
(a copy of it made with `git archive`: commit d2cf8a4 plus this commit's script). Every line below is the script's own
output: counts, booleans, seconds and member CODES only - no figure, no label, no business text, no token, cookie or URL
query. Errors are shown as their class (and HTTP
status). A FinPlan member's name is its label: such a member is shown by its dimension and place (`Account #3`).

## Verdict (Task 3b re-run, 2026-11-22 01:17-01:22 UTC): passed

| Part | Result |
| --- | --- |
| 1. Order: members the atlas keeps under their label, asked against outline order (rows and columns, both applications) | **passed** (4 of 4 cases) |
| 2. Twenty real grids from the stored answers, on both routes (plus a FinLite extra and a split) | **passed** (21 of 21 identical `Fetched`; the split over a lowered budget identical) |
| 2b. The requests the provider route refuses for a member (`svhttp.WrongMember`), 6 | **2 left out** with one warning and one caveat each, what remains equal on both routes; **4 cannot be left out** (the request's only Entity): refused with the clear error (ruling (a)); **0 other** |
| 2c. The drop on real figures: an invented code in the middle of the accounts of a real FinMax and a real FinPlan request | **passed** (left out, one warning, every fact the provider route's for the request as stored: 315 and 116 facts) |
| Probes: unknown member, cell budget, the largest grid split | recorded below; split: 0 differences, **11 stale aliases** (accepted class) |

The two accepted classes (controller rulings of Task 3b): a **stale alias** (a member whose alias changed in EPM since
the atlas was built: same code, every number and fact equal; a REST `Fetched` only feeds facts) is counted by code,
not as a difference; a **wrong member** follows ruling (a). Every other difference is zero, so the Excel fallback is
deleted in the next commit of this branch.

The Excel fallback on these 4 requests (verified offline, no Excel, at commit d2cf8a4: `xl/pull_adhoc._drop_rejected`
on a grid whose rejected member is its axis's only member, the Smart View retrieve faked): **layout unchanged**, 0
members left out, 12 test retrieves - it failed these 4 requests too. The member behind them is 1 FinPlan Entity,
**not in today's atlas** (the stored answers name an entity EPM and the atlas no longer have: a question to re-ask
with today's entity, not a stale atlas).

## What a wrong member does now (Task 3b)

- The provider route's wrong-member errors ("Grid is Invalid", "Member name and Alias placed on row N do not match",
  "Member ... is out of place") are `svhttp.WrongMember`: final on that route (no reconnect), matched as the whole
  sentence from the start of the provider's message, and sent on to REST. Live: 6 of 6 such requests were
  `WrongMember`.
- Oracle's HTTP 400 for such a member (probed live 2026-11-22 with invented codes on rows, columns and the POV of both
  applications; the answer's shape only): keys `status` (int), `detail`, `message`, `localizedMessage` (the same
  text); the sentence `The member <M> does not exist in the <D> dimension for the specified cube or you do not have
  access to it.` - `<M>` as sent, in double quotes when it has a space, `<D>` the dimension's REST name; of two wrong
  members only the first is named. `rest.py` matches the whole sentence and acts only on a member asked exactly so in
  that dimension.
- A member on an axis of several is left out of every grid of the request (one per round, at most 3) and the reply
  carries one warning and one warn caveat per member: its figures are missing and every total with it is INCOMPLETE.
  A POV member or an axis's only member is an error that names it, the grid and `smartass atlas build`.

## Part 1: order

The atlas keeps periods, scenarios and all FinPlan accounts and entities under their label; REST answers their real
names (Y26M01 ...), which the atlas does not know, so for them the place alone ties a label to a number. Three such
members of one parent were asked as the third, the first, then the second child (`against`): on rows and on
columns, in FinMax (months of Q1 26: periods on rows against one account, then on columns) and in FinPlan (three
children of an account parent with three different rows of numbers on rows, three months of Q1 on columns, one grid).
Each member was also read alone over REST first, for the name REST gives it; the multi-member answer must carry those
names in the order asked. The provider's rows (or columns) must differ, so a swap could be seen. The atlas's `ord`
lists Q1's months as Mar, Feb, Jan; the asked Jan, Mar, Feb is against both that order and the calendar's.

Result: **REST answers rows and columns in the order asked.** No re-ordering is needed; `rest.py`'s docstring now
says so (it said it as a fact before it was confirmed).

## Part 2: twenty grids

Source: the run log (`runs.jsonl`) holds no spec - it names each answer's (workbook, Q) - so the specs are the stored
answers' (`answers.db`, the latest spec of each (workbook, Q) the run log names): 460 answers planned with today's
closed month into 913 grid requests, 69 different, 63 fetched by the provider route (the 6 others: `WrongMember`, a
member the cube no longer has - part 2b). The 20 were chosen to cover: a FinPlan request over
several years (01, 04, 06, 08, 09, 11: 2 grids; 13: 3 grids), the largest (02, 520 cells), rows with no data at all
(01, 02, 03, 05 ...), empty cells among numbers (04, 20), products as a third row dimension (02, 03), FinMax and
FinPlan; every grid has members the atlas keeps under their label (periods, FinPlan's accounts and entities). The
run log holds no FinLite request: line 21 is request 02 asked on FinLite (an extra, not one of the 20). No logged
grid comes near the cell budget, so 02 was also fetched over a lowered budget (247 counted cells: 4 requests,
stitched) and the probe below splits a 400 000-cell grid over the default budget.

Task 3 ran these parts first (21-22 November): the same result, except that the split probe's 11 label cells were
counted as differences (the stale-alias class, accepted since) and the wrong-member requests were not run.

## Probes

- **Empty cell**: REST answers `""` (1 521 times over the grids, no other non-number form); the provider route's
  `Fetched` holds `None` there. `rest.py` maps `""`, `" "` and `null` to `None`.
- **#NoAccess (open item)**: not seen on either route (0 cells in 63 grids). Oracle documents that exportdataslice
  exports "only cells for which the user has read-write access" and names no marker for the others, so over REST a
  cell the user may not read cannot be told apart from an empty one: it reads as missing, not as #NoAccess. The
  reply's one warning for REST grids says so. To be checked when a no-access cell is available.
- **The split loop after the gate**: the gate ran on d2cf8a4; 93c6215 changed the split loop afterwards
  (each part's answer is checked before the next part is sent), covered offline by item 1's test of Task 3c
  (`test_rest.py::test_a_split_grid_whose_first_answer_is_not_the_grid_asked_sends_no_other_part`, 8ab749e). The
  grids part was run again at a695d72 (Task 4b): all equal; the controller ran both gates once more at 8049f2a
  ("Final run" below).
- **Unknown member** (an invented code, `ac.ZZ_NO_SUCH_MEMBER`): REST's HTTP 400 names it (`MemberRefused`); on
  rows it can be left out, in the POV it cannot. The provider route fails the grid without naming it.
- **Cell budget**: EPM refuses a grid over 500 000 cells as it counts them: the grid's data rows and its two header
  rows, times a row's member headers and data cells (`rest.epm_cells`). Two live samples agree exactly: 588 028 for
  42 000 rows of 14 (Task 3) and 933 184 for 66 654 rows of 14 (below). `MAX_CELLS` is 250 000 counted that way.
- **A member asked twice on one axis**: REST answers it once; `rest.py` refuses such a grid before it is sent.
- **One run interrupted**: a run at 01:08 UTC met plan1 down on both routes (the provider "Unable to obtain a
  connection", REST HTTP 500); EPM answered again at 01:16 and the run below was made after it.

## The run (verbatim)

```
stored answers 460 | planned 460 | requests 913
order FinMax: entity un.SY account ac.C30 periods Jan 26, Mar 26, Feb 26
order A FinMax rows: members 3 | against outline order True | provider rows distinct True | REST names in the order asked True | layout equal True | cells 65 equal 65 close 0 differ 0 | provider 0.2s rest 0.5s | OK
order B FinMax cols: members 3 | against outline order True | provider cols distinct True | REST names in the order asked True | layout equal True | cells 39 equal 39 close 0 differ 0 | provider 0.1s rest 0.1s | OK
order FinPlan: parents tried 2
order C FinPlan rows: members 3 | against outline order True | provider rows distinct True | REST names in the order asked True | layout equal True | cells 35 equal 35 close 0 differ 0 | provider 0.1s rest 0.2s | OK
order D FinPlan cols: members 3 | against outline order True | provider cols distinct True | REST names in the order asked True | layout equal True | cells 35 equal 35 close 0 differ 0 | provider 0.1s rest 0.1s | OK
candidate FinPlan: provider failed (WrongMember) - not chosen
candidate FinPlan: provider failed (WrongMember) - not chosen
candidate FinPlan: provider failed (WrongMember) - not chosen
candidate FinPlan: provider failed (WrongMember) - not chosen
candidate FinPlan: provider failed (WrongMember) - not chosen
candidate FinPlan: provider failed (WrongMember) - not chosen
grids: 69 different requests from 460 answers, 63 fetched by the provider
extra: the run log holds no FinLite request - the last line is a FinMax request of the run log asked on FinLite (the same members)
01 FinPlan: grids 2 | rows 14 | cells 154 | products False | balance Periodic | missing cells 154 | all-missing rows 14 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 234 close 0 stale aliases 0 differ 0 | provider 0.3s rest 0.2s | OK
02 FinMax: grids 1 | rows 52 | cells 520 | products True | balance Periodic | missing cells 205 | all-missing rows 17 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 702 close 0 stale aliases 0 differ 0 | provider 0.2s rest 0.1s | OK
02 FinMax (budget 247): grids 1 | rows 52 | cells 520 | products True | balance Periodic | missing cells 205 | all-missing rows 17 | #NoAccess 0 | REST requests 4 | warnings 0 | equal 702 close 0 stale aliases 0 differ 0 | provider 0.2s rest 0.4s | OK
03 FinMax: grids 1 | rows 52 | cells 520 | products True | balance Periodic | missing cells 208 | all-missing rows 19 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 702 close 0 stale aliases 0 differ 0 | provider 0.2s rest 0.1s | OK
04 FinPlan: grids 2 | rows 22 | cells 121 | products False | balance Periodic | missing cells 5 | all-missing rows 0 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 247 close 0 stale aliases 0 differ 0 | provider 0.2s rest 0.2s | OK
05 FinMax: grids 1 | rows 15 | cells 150 | products False | balance Periodic | missing cells 31 | all-missing rows 2 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 221 close 0 stale aliases 0 differ 0 | provider 0.1s rest 0.1s | OK
06 FinPlan: grids 2 | rows 14 | cells 154 | products False | balance Periodic | missing cells 14 | all-missing rows 1 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 234 close 0 stale aliases 0 differ 0 | provider 0.3s rest 0.3s | OK
07 FinMax: grids 1 | rows 15 | cells 150 | products False | balance Periodic | missing cells 37 | all-missing rows 3 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 221 close 0 stale aliases 0 differ 0 | provider 0.1s rest 0.1s | OK
08 FinPlan: grids 2 | rows 24 | cells 132 | products False | balance Periodic | missing cells 132 | all-missing rows 24 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 266 close 0 stale aliases 0 differ 0 | provider 0.2s rest 0.2s | OK
09 FinPlan: grids 2 | rows 12 | cells 132 | products False | balance Periodic | missing cells 132 | all-missing rows 12 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 208 close 0 stale aliases 0 differ 0 | provider 0.2s rest 0.2s | OK
10 FinPlan: grids 1 | rows 10 | cells 100 | products False | balance Periodic | missing cells 100 | all-missing rows 10 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 144 close 0 stale aliases 0 differ 0 | provider 0.1s rest 0.1s | OK
11 FinPlan: grids 2 | rows 8 | cells 84 | products False | balance Periodic | missing cells 0 | all-missing rows 0 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 150 close 0 stale aliases 0 differ 0 | provider 0.2s rest 0.2s | OK
12 FinMax: grids 1 | rows 7 | cells 70 | products False | balance Periodic | missing cells 10 | all-missing rows 1 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 117 close 0 stale aliases 0 differ 0 | provider 0.1s rest 0.1s | OK
13 FinPlan: grids 3 | rows 9 | cells 66 | products False | balance Periodic | missing cells 66 | all-missing rows 9 | #NoAccess 0 | REST requests 3 | warnings 0 | equal 160 close 0 stale aliases 0 differ 0 | provider 0.5s rest 0.4s | OK
14 FinPlan: grids 1 | rows 6 | cells 60 | products False | balance Periodic | missing cells 0 | all-missing rows 0 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 96 close 0 stale aliases 0 differ 0 | provider 0.1s rest 0.1s | OK
15 FinMax: grids 1 | rows 1 | cells 9 | products False | balance Periodic | missing cells 0 | all-missing rows 0 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 39 close 0 stale aliases 0 differ 0 | provider 0.1s rest 0.1s | OK
16 FinMax: grids 1 | rows 1 | cells 10 | products False | balance Periodic | missing cells 0 | all-missing rows 0 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 39 close 0 stale aliases 0 differ 0 | provider 0.1s rest 0.1s | OK
17 FinMax: grids 1 | rows 1 | cells 10 | products False | balance Periodic | missing cells 0 | all-missing rows 0 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 39 close 0 stale aliases 0 differ 0 | provider 0.1s rest 0.1s | OK
18 FinMax: grids 1 | rows 7 | cells 42 | products False | balance Periodic | missing cells 6 | all-missing rows 1 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 117 close 0 stale aliases 0 differ 0 | provider 0.1s rest 0.1s | OK
19 FinMax: grids 1 | rows 7 | cells 42 | products False | balance Periodic | missing cells 6 | all-missing rows 1 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 117 close 0 stale aliases 0 differ 0 | provider 0.2s rest 0.1s | OK
20 FinMax: grids 1 | rows 11 | cells 110 | products False | balance Periodic | missing cells 5 | all-missing rows 0 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 169 close 0 stale aliases 0 differ 0 | provider 0.2s rest 0.1s | OK
21 FinLite: grids 1 | rows 52 | cells 520 | products True | balance Periodic | missing cells 205 | all-missing rows 17 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 702 close 0 stale aliases 0 differ 0 | provider 0.0s rest 0.2s | OK
probe left out FinMax: invented member at place 3 of 5 accounts | left out True | reply warnings 1 | facts 315, of them not 0 311, equal to the provider's for the request as stored True | OK
probe left out FinPlan: invented member at place 6 of 12 accounts | left out True | reply warnings 1 | facts 116, of them not 0 113, equal to the provider's for the request as stored True | OK
wrong members: 6 requests the provider route refused, 6 of them svhttp.WrongMember
wrong 01 FinPlan: provider WrongMember | REST refused a member that cannot be left out: Entity, axis of 1 | clear error True | REFUSED
wrong 02 FinPlan: provider WrongMember | REST refused a member that cannot be left out: Entity, axis of 1 | clear error True | REFUSED
wrong 03 FinPlan: provider WrongMember | REST refused a member that cannot be left out: Entity, axis of 1 | clear error True | REFUSED
wrong 04 FinPlan: provider WrongMember | REST refused a member that cannot be left out: Entity, axis of 1 | clear error True | REFUSED
wrong 05 FinPlan: provider WrongMember | REST left out 1 (Account #3) | reply warnings 1 caveats 1 | without them: grids 3 equal 128 close 0 stale aliases 0 differ 0 | facts 0, equal True | rest 0.4s | OK
wrong 06 FinPlan: provider WrongMember | REST left out 1 (Account #3) | reply warnings 1 caveats 1 | without them: grids 3 equal 128 close 0 stale aliases 0 differ 0 | facts 0, equal True | rest 0.4s | OK
wrong members that cannot be left out: 1 distinct member(s)
   FinPlan Entity #1: in today's atlas False
raw REST cells that are not numbers, over the grids: {"blank str of 0 -> None (blank)": 1521}
probe unknown member ac.ZZ_NO_SUCH_MEMBER in rows: provider failed: SvHttpError (different layout) | REST failed: MemberRefused; names the member True; can be left out True
probe unknown member ac.ZZ_NO_SUCH_MEMBER in POV: provider failed: SvHttpError (appended POV) | REST failed: MemberRefused; names the member True; can be left out False
probe budget: level-0 accounts 3326, entities 4443
probe budget 24996 cells (2083 x 1 x 12): answered in 0.7s (2085 rows, answer 0.2 MB)
probe budget 49992 cells (2083 x 2 x 12): answered in 1.2s (4168 rows, answer 0.3 MB)
probe budget 99972 cells (2777 x 3 x 12): answered in 2.0s (8333 rows, answer 0.7 MB)
probe budget 199944 cells (2777 x 6 x 12): answered in 3.4s (16664 rows, answer 1.4 MB)
probe budget 399960 cells (3030 x 11 x 12): answered in 7.2s (33332 rows, answer 2.8 MB)
probe budget 799848 cells: failed in 2.5s: RestError HTTP 400; EPM counted 933184, rest.epm_cells 933184, equal True
probe split: the largest grid over MAX_CELLS 250000: REST requests 2 | layout equal True | equal 466637 close 0 differ 0 stale aliases 11 | facts 1833, equal True | provider 7.7s rest 6.1s | OK
result: all equal
```

## Re-run with the atlas built over REST (Task 4b), 2026-11-22 06:12-06:16 UTC: passed

`scripts/compare_grid_routes.py grids` run by `scripts/compare_atlas.py <work folder> routes` at commit a695d72, with
config pointed at a copy of the atlas rebuilt over REST (its side table of real names trusted, so the REST engine's
exact real-name check of a45bda9 / b11bbd3 / 749574b is active) - the owner's atlas untouched. It also covers this
file's open re-run after 93c6215.

- Stored answers 460, planned 460, 913 requests; 69 different requests, 63 fetched by the provider (6 refused by it as
  WrongMember: their own class below).
- 21 grids (20 + the FinLite extra) and the split: **all equal** on both routes (22 OK), 0 stale aliases.
- Drop probes (an invented code in a real request): FinMax and FinPlan left out with one warning, every fact equal to
  the provider's for the request as stored (315 and 116 facts).
- Wrong members, 6 requests: 2 left out (Account #3) with a warning and a caveat, what remains equal; 4 refused with
  the clear error (the request's only Entity; 1 distinct FinPlan entity, not in the rebuilt atlas either).
- Result: all equal.

## Final run, 2026-11-22 12:27-12:44 local time (11:27-11:44 UTC), at 8049f2a: passed

Run by the controller from `engine/` at 8049f2a (the phase's head before its last round), read-only, without Excel:
`scripts/compare_atlas.py <work folder> all`, whose `routes` part runs `scripts/compare_grid_routes.py grids` with
config pointed at the atlas it rebuilt over REST (the owner's atlas untouched), and the `routes` part once more on its
own. Both printed the same lines and exited 0 (the part took 223.5 s in the first, 227.2 s in the second). What ran:
the grids part only - the 20 stored grids, the FinLite extra and request 02 over a lowered cell budget, the two drop
probes and the wrong-member requests. Not run here: Part 1 (order) and the probes (the 466 637-cell split, the cell
budget, the unknown member); their last live run is Task 3b's above.

| Part | Result |
| --- | --- |
| Requests | 460 stored answers, all planned, 913 grid requests; 69 different, 63 fetched by the provider route (the other 6 refused by it as `svhttp.WrongMember`) |
| Grids on both routes | 22 OK (the 21 grids and 02 over the lowered budget): all equal |
| Drop probes | FinMax (an invented member at place 3 of 5 accounts) and FinPlan (place 6 of 12): left out, one warning, 315 and 116 facts equal to the provider route's for the request as stored |
| Wrong members | 6 requests, all `WrongMember` on the provider route: 4 refused over REST with the clear error (the request's only Entity, one FinPlan entity, not in today's atlas); 2 left out (Account #3) with one warning and one caveat, the 3 grids without it equal (128 cells), facts equal |
| Raw REST cells that are not numbers | 1 521 blank strings, read as empty |
| Result | all equal |

The routes lines, verbatim (the second run's; the first printed the same):

```
stored answers 460 | planned 460 | requests 913
candidate FinPlan: provider failed (WrongMember) - not chosen
candidate FinPlan: provider failed (WrongMember) - not chosen
candidate FinPlan: provider failed (WrongMember) - not chosen
candidate FinPlan: provider failed (WrongMember) - not chosen
candidate FinPlan: provider failed (WrongMember) - not chosen
candidate FinPlan: provider failed (WrongMember) - not chosen
grids: 69 different requests from 460 answers, 63 fetched by the provider
extra: the run log holds no FinLite request - the last line is a FinMax request of the run log asked on FinLite (the same members)
probe left out FinMax: invented member at place 3 of 5 accounts | left out True | reply warnings 1 | facts 315, of them not 0 311, equal to the provider's for the request as stored True | OK
probe left out FinPlan: invented member at place 6 of 12 accounts | left out True | reply warnings 1 | facts 116, of them not 0 113, equal to the provider's for the request as stored True | OK
wrong members: 6 requests the provider route refused, 6 of them svhttp.WrongMember
wrong 01 FinPlan: provider WrongMember | REST refused a member that cannot be left out: Entity, axis of 1 | clear error True | REFUSED
wrong 02 FinPlan: provider WrongMember | REST refused a member that cannot be left out: Entity, axis of 1 | clear error True | REFUSED
wrong 03 FinPlan: provider WrongMember | REST refused a member that cannot be left out: Entity, axis of 1 | clear error True | REFUSED
wrong 04 FinPlan: provider WrongMember | REST refused a member that cannot be left out: Entity, axis of 1 | clear error True | REFUSED
wrong 05 FinPlan: provider WrongMember | REST left out 1 (Account #3) | reply warnings 1 caveats 1 | without them: grids 3 equal 128 close 0 stale aliases 0 differ 0 | facts 0, equal True | rest 0.4s | OK
wrong 06 FinPlan: provider WrongMember | REST left out 1 (Account #3) | reply warnings 1 caveats 1 | without them: grids 3 equal 128 close 0 stale aliases 0 differ 0 | facts 0, equal True | rest 0.4s | OK
wrong members that cannot be left out: 1 distinct member(s)
   FinPlan Entity #1: in today's atlas False
raw REST cells that are not numbers, over the grids: {"blank str of 0 -> None (blank)": 1521}
result: all equal
grids OK 22, exit 0
== routes: passed (227.2 s)
```

The commits after 8049f2a (Task 8 of the phase) change the scrub of texts (`core/logsetup.py`, which the
provider route's error text also passes), how `xl/target._open_path` reads the workbooks of an Excel it just
launched, one bounded retry for a file replace or delete (`core/retry.py`), the test guards, the CLI's help text and
the docs. None touches the fetch routes (`epm/svhttp.py`, `epm/rest.py`, `data/provider.py`, `epm/epmhttp.py`), the
name check or the atlas build (`epm/members.py`, `ops/atlas_build.py`, `knowledge/atlas.py`): `git diff --stat
8049f2a..` over those files is empty. The scrub change rewrites less ordinary text and adds only secret-named shapes
(a cookie list, quotes escaped by number, `device_code`); the provider's wrong-member sentence it may pass is matched
as before. After this run both gate scripts were deleted (`scripts/compare_grid_routes.py`,
`scripts/compare_atlas.py`); this file keeps their results and the "Final run" of docs/audit/2026-11-22-atlas-rest.md
those of the atlas gate.

The routes part was run once more by the controller at 2b0c693 (Task 8's head; 22.11.2026 13:34-13:41, read-only,
without Excel, the scripts taken from the commit before their deletion): passed, the same lines (all equal, both
left-out probes OK, the 6 wrong-member requests: 4 refused, 2 left out; exit 0), 426.2 s.

Task 9, the final fix round after both final reviews, DOES touch the fetch routes: `epm/svhttp.py` (the provider's
error is classified on its raw `<desc>`, which `epm/svxml.py` keeps beside the scrubbed text; and comments only),
`epm/rest.py` (a cell is read only as a plain number: any other text fails the grid), `epm/svxml.py` (the read-only
check refuses CDATA, a comment, a DOCTYPE or a processing instruction), the scrub (`core/logsetup.py`: narrower on
ordinary text after `Basic` and `Authorization:`, and the `%` values after a cookie's name, `SID` or `SESSION_ID`
that Task 8's change stopped catching are caught again) and `pipeline.py` (the REST warning's quoted error is
scrubbed); not `data/provider.py`, `epm/epmhttp.py`, the name check or the atlas build. The routes comparison was
run again after the final fix round; the audit records it.
<!-- Task 9: the controller records here the run of the routes part at Task 9's head (date, head, result). -->
