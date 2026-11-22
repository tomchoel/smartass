# The second engine against the provider route (phase 2, Task 3), 2026-11-21/22

Live, read-only against EPM (plan1 and stratmodel), without Excel, from the owner's PC with the stored refresh tokens
(no sign-in was needed). Script: `scripts/compare_grid_routes.py all` (deleted in Task 7), run from `engine/` on
the tree of the commit that adds this file. Every line below is the script's own output: counts, booleans, seconds
and member CODES only - no figure, no label, no business text, no token, cookie or URL query. Errors are shown as
their class (and HTTP status).

## Verdict

| Part | Result |
| --- | --- |
| 1. Order: members the atlas keeps under their label, asked against outline order (rows and columns, both applications) | **passed** (4 of 4 cases) |
| 2. Twenty real grids from the stored answers, on both routes | **passed** (20 of 20, plus a FinLite extra: identical `Fetched`) |
| Probes: empty cell, unknown member, cell budget | recorded below |
| Extra probe: the largest grid one request takes, split over the default `MAX_CELLS` | **11 label cells differ** (one account's alias changed in EPM after the atlas was built; same member code, every number and every fact equal) |

The script exits 1 because of the extra probe. **The Excel fallback is NOT deleted**: a class of grids differs
(below, "Stale alias") and this task cannot make it equal - the controller decides.

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
closed month into 913 grid requests, 69 different, 63 fetched by the provider route (the 6 others: `GridRefused`, a
member the cube no longer has - see "A wrong member" below). The 20 were chosen to cover: a FinPlan request over
several years (01, 04, 06, 08, 09, 11: 2 grids; 13: 3 grids), the largest (02, 520 cells), rows with no data at all
(01, 02, 03, 05 ...), empty cells among numbers (04, 20), products as a third row dimension (02, 03), FinMax and
FinPlan; every grid has members the atlas keeps under their label (periods, FinPlan's accounts and entities). The
run log holds no FinLite request: line 21 is request 02 asked on FinLite (an extra, not one of the 20). No logged
grid comes near the cell budget, so 02 was also fetched over a lowered budget (225: 4 requests, stitched) and the
probe below splits a 400 000-cell grid over the default budget.

## Probes (what Task 2 left unconfirmed)

- **Empty cell**: REST answers `""` (1 521 times over the grids, no other non-number form); the provider route's
  `Fetched` holds `None` there (svxml's blank). `rest.py` now maps `""`, `" "` and `null` to `None` (it
  mapped them to "#Missing" before: every missing cell differed on the first run).
- **#NoAccess**: not seen on either route (0 cells in 63 grids): still unconfirmed live.
- **Unknown member** (an invented code, `ac.ZZ_NO_SUCH_MEMBER`): REST answers HTTP 400 naming the member (Oracle:
  "The member ... does not exist in the ... dimension"); the provider route fails the grid too (rows: "different
  layout", POV: "appended POV").
- **Cell budget**: EPM refuses a grid over 500 000 cells *as it counts them* - a row's member headers and its data
  cells (HTTP 400 "Unable to load the form as the number of cells exceeded the threshold ... Error Threshold: 500000,
  Current Value: 588028" for 42 000 rows of 2 headers and 12 cells). 466 000 counted cells answered in 6.0 s (2.8 MB).
  `MAX_CELLS` is now 250 000 counted cells (half the threshold), and the split counts a row as EPM does.
- **A member asked twice on one axis**: REST answers it once (984 rows for 1 000 asked, 16 asked twice), the provider
  twice. `rest.py` refuses such a grid before it is sent (no logged request has one: GridRequest's accounts and
  products are de-duplicated).

## The classes found, and what was done

- **Stale alias (not fixed, the reason the Excel fallback stays)**: a member the atlas keeps under its real name
  gets its label from the atlas; the provider route shows its current Default alias. Where the alias changed in EPM
  after the atlas was built, the label TEXT differs (same code). Atlas 11.7 days old: 1 of 4 741 FinMax accounts
  (`ac.24074`) and 5 of 18 028 FinMax entities differ (one provider grid per dimension, labels compared as
  booleans). In the split probe this is all 11 differing cells; the facts an answer is computed from
  (`facts_from_parsed`: codes and numbers) are equal (1 833 facts). A REST `Fetched` only feeds
  `data/provider.py`'s facts (it is never written to a sheet), so no answer differs - but `Fetched` is not
  identical. REST has no alias option; Task 4's atlas over REST (rebuilt in about half a minute) is what shrinks it.
- **A member the atlas holds under a name REST gives another member (fixed)**: FinPlan's `ac.EBITDA` is in the atlas
  with a label of another form, while REST answers `ac.EBITDA` for the account the provider shows under the asked
  label (the layout check caught it). The label of a member asked by the label the atlas keeps is now that label.
- **A wrong member (provider route made deterministic)**: the 6 logged requests the provider cannot fetch fail on both
  routes (provider "Member name and Alias placed on row 3 do not match" / "Member ... is out of place"; REST HTTP 400
  "The member ... does not exist ..."). These provider errors are now `GridRefused` (no reconnect, no REST try), as
  "Grid is Invalid" and the alias clash were. **No REST equivalent of an Excel capability:** the Excel fallback
  isolated rejected members (rc=1000 bisect) and fetched the rest with a note; over REST a wrong member fails the
  whole grid.

## The run (verbatim)

```
stored answers 460 | planned 460 | requests 913
order FinMax: entity un.SY account ac.C30 periods Jan 26, Mar 26, Feb 26
order A FinMax rows: members 3 | against outline order True | provider rows distinct True | REST names in the order asked True | layout equal True | cells 65 equal 65 close 0 differ 0 | provider 0.2s rest 0.3s | OK
order B FinMax cols: members 3 | against outline order True | provider cols distinct True | REST names in the order asked True | layout equal True | cells 39 equal 39 close 0 differ 0 | provider 0.2s rest 0.1s | OK
order FinPlan: parents tried 2
order C FinPlan rows: members 3 | against outline order True | provider rows distinct True | REST names in the order asked True | layout equal True | cells 35 equal 35 close 0 differ 0 | provider 0.4s rest 0.6s | OK
order D FinPlan cols: members 3 | against outline order True | provider cols distinct True | REST names in the order asked True | layout equal True | cells 35 equal 35 close 0 differ 0 | provider 0.3s rest 0.1s | OK
candidate FinPlan: provider failed (GridRefused) - not chosen
candidate FinPlan: provider failed (GridRefused) - not chosen
candidate FinPlan: provider failed (GridRefused) - not chosen
candidate FinPlan: provider failed (GridRefused) - not chosen
candidate FinPlan: provider failed (GridRefused) - not chosen
candidate FinPlan: provider failed (GridRefused) - not chosen
grids: 69 different requests from 460 answers, 63 fetched by the provider
extra: the run log holds no FinLite request - the last line is a FinMax request of the run log asked on FinLite (the same members)
01 FinPlan: grids 2 | rows 14 | cells 154 | products False | balance Periodic | missing cells 154 | all-missing rows 14 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 234 close 0 differ 0 | provider 0.2s rest 0.2s | OK
02 FinMax: grids 1 | rows 52 | cells 520 | products True | balance Periodic | missing cells 205 | all-missing rows 17 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 702 close 0 differ 0 | provider 1.2s rest 0.2s | OK
02 FinMax (budget 225): grids 1 | rows 52 | cells 520 | products True | balance Periodic | missing cells 205 | all-missing rows 17 | #NoAccess 0 | REST requests 4 | warnings 0 | equal 702 close 0 differ 0 | provider 1.2s rest 0.5s | OK
03 FinMax: grids 1 | rows 52 | cells 520 | products True | balance Periodic | missing cells 208 | all-missing rows 19 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 702 close 0 differ 0 | provider 0.2s rest 0.2s | OK
04 FinPlan: grids 2 | rows 22 | cells 121 | products False | balance Periodic | missing cells 5 | all-missing rows 0 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 247 close 0 differ 0 | provider 0.2s rest 0.2s | OK
05 FinMax: grids 1 | rows 15 | cells 150 | products False | balance Periodic | missing cells 31 | all-missing rows 2 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 221 close 0 differ 0 | provider 0.1s rest 0.1s | OK
06 FinPlan: grids 2 | rows 14 | cells 154 | products False | balance Periodic | missing cells 14 | all-missing rows 1 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 234 close 0 differ 0 | provider 0.3s rest 0.3s | OK
07 FinMax: grids 1 | rows 15 | cells 150 | products False | balance Periodic | missing cells 37 | all-missing rows 3 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 221 close 0 differ 0 | provider 0.2s rest 0.1s | OK
08 FinPlan: grids 2 | rows 24 | cells 132 | products False | balance Periodic | missing cells 132 | all-missing rows 24 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 266 close 0 differ 0 | provider 0.2s rest 0.2s | OK
09 FinPlan: grids 2 | rows 12 | cells 132 | products False | balance Periodic | missing cells 132 | all-missing rows 12 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 208 close 0 differ 0 | provider 0.2s rest 0.2s | OK
10 FinPlan: grids 1 | rows 10 | cells 100 | products False | balance Periodic | missing cells 100 | all-missing rows 10 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 144 close 0 differ 0 | provider 0.1s rest 0.1s | OK
11 FinPlan: grids 2 | rows 8 | cells 84 | products False | balance Periodic | missing cells 0 | all-missing rows 0 | #NoAccess 0 | REST requests 2 | warnings 0 | equal 150 close 0 differ 0 | provider 0.2s rest 0.2s | OK
12 FinMax: grids 1 | rows 7 | cells 70 | products False | balance Periodic | missing cells 10 | all-missing rows 1 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 117 close 0 differ 0 | provider 0.1s rest 0.2s | OK
13 FinPlan: grids 3 | rows 9 | cells 66 | products False | balance Periodic | missing cells 66 | all-missing rows 9 | #NoAccess 0 | REST requests 3 | warnings 0 | equal 160 close 0 differ 0 | provider 0.5s rest 0.5s | OK
14 FinPlan: grids 1 | rows 6 | cells 60 | products False | balance Periodic | missing cells 0 | all-missing rows 0 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 96 close 0 differ 0 | provider 0.1s rest 0.1s | OK
15 FinMax: grids 1 | rows 1 | cells 9 | products False | balance Periodic | missing cells 0 | all-missing rows 0 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 39 close 0 differ 0 | provider 0.2s rest 0.1s | OK
16 FinMax: grids 1 | rows 1 | cells 10 | products False | balance Periodic | missing cells 0 | all-missing rows 0 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 39 close 0 differ 0 | provider 0.2s rest 0.1s | OK
17 FinMax: grids 1 | rows 1 | cells 10 | products False | balance Periodic | missing cells 0 | all-missing rows 0 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 39 close 0 differ 0 | provider 0.4s rest 0.1s | OK
18 FinMax: grids 1 | rows 7 | cells 42 | products False | balance Periodic | missing cells 6 | all-missing rows 1 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 117 close 0 differ 0 | provider 0.3s rest 0.1s | OK
19 FinMax: grids 1 | rows 7 | cells 42 | products False | balance Periodic | missing cells 6 | all-missing rows 1 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 117 close 0 differ 0 | provider 0.3s rest 0.1s | OK
20 FinMax: grids 1 | rows 11 | cells 110 | products False | balance Periodic | missing cells 5 | all-missing rows 0 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 169 close 0 differ 0 | provider 0.2s rest 0.1s | OK
21 FinLite: grids 1 | rows 52 | cells 520 | products True | balance Periodic | missing cells 205 | all-missing rows 17 | #NoAccess 0 | REST requests 1 | warnings 0 | equal 702 close 0 differ 0 | provider 0.0s rest 0.3s | OK
raw REST cells that are not numbers, over the grids: {"'' -> None (blank)": 1521}
probe unknown member ac.ZZ_NO_SUCH_MEMBER in rows: provider failed: SvHttpError (different layout) | REST failed: RestError HTTP 400; names the member True
probe unknown member ac.ZZ_NO_SUCH_MEMBER in POV: provider failed: SvHttpError (appended POV) | REST failed: RestError HTTP 400; names the member True
probe budget: level-0 accounts 3326, entities 4443
probe budget 24996 cells (2083 x 1 x 12): answered in 0.8s (2085 rows, answer 0.2 MB)
probe budget 49992 cells (2083 x 2 x 12): answered in 1.2s (4168 rows, answer 0.3 MB)
probe budget 99972 cells (2777 x 3 x 12): answered in 2.1s (8333 rows, answer 0.7 MB)
probe budget 199944 cells (2777 x 6 x 12): answered in 3.4s (16664 rows, answer 1.4 MB)
probe budget 399960 cells (3030 x 11 x 12): answered in 6.0s (33332 rows, answer 2.8 MB)
probe budget 799848 cells: failed in 2.4s: RestError HTTP 400
probe split: the largest grid over MAX_CELLS 250000: REST requests 2 | layout equal True | equal 466637 close 0 differ 11 (of them member labels 11) | facts 1833, equal True | provider 9.6s rest 6.2s | DIFF
result: DIFFERENCES
```
