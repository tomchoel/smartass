# The atlas over the token session (phase 2, Task 4), 2026-11-22

Live, read-only against EPM (plan1 and stratmodel), without Excel, from the owner's PC with the stored refresh tokens
(no sign-in was needed). Every line below is a script's own output or a count taken from it: HTTP statuses, seconds,
sizes, key names, Oracle's property values (data storage, account type, time balance), counts, booleans and member
CODES only - no label, no alias text, no figure, no business text, no token, cookie or URL query. The owner's
`atlas.db` and `knowledge.db` were never written: the comparisons read a copy of today's atlas (built 11.8 days
earlier by the hidden Excel's zoom) made in a scratch folder.

## Probe (Step 0)

Scripts (not in the repo): `.superpowers/sdd/2026-11-21-light-fast-p2-rest/spike/atlas_probe.py` .. `atlas_probe4.py`.
Each request went through `epmhttp.EpmSession` and a rule added in its own commit before the probe that used it
(1e8abd9: two dimensions; 14047a8: every dimension; 70e1c77: the `fields` projection).

### Candidate 1: Oracle's "Get Member" - answered; the route

`GET /HyperionPlanning/rest/v3/applications/{app}/dimensions/{dim}/members/{dim}` - the dimension's top member and,
by default (`maxDepth` -1), its whole visible subtree. Oracle documents it for Service Administrators; it answered
this (non-administrator) user, as the plan-type dimension list did in Task 2. It creates nothing (Add Member is a POST
to `.../members`, never allowlisted).

- Status: **201** on every one of the 22 dimensions (14 FinMax, 8 FinPlan), both pods - Oracle documents 201 as this
  GET's historical success status. Content type `application/json; charset=UTF-8`.
- Shape: one JSON object per member, nested: `children` is a list of member objects (the WHOLE hierarchy in one
  answer, not one level). Keys of a member without a projection: `alias` (str, absent when the member has none),
  `aliasPath`, `children` (list, absent on a leaf), `dataStorage`, `dataType`, `dimName`, `displayPath`,
  `generation` (int), `hasMbrFormula` (bool), `id`, `invalidDueToValidIntersection` (bool), `level` (int),
  `links` (list), `name`, `objectType` (int), `oldName`, `parentId`, `parentName` (str, not on the top member),
  `path`, `twoPass` (bool), `uniqueName` (str, on shared members), `usedIn` (list of plan-type names), `valid`
  (bool); plus `description`, `udas`, `formula` on some dimensions, and `accountType`, `timeBalance`,
  `skipValue` on account-type dimensions (FinPlan Account; FinMax BalanceTypes). FinMax's Accounts dimension is not
  an account-type dimension: it carries no account type or time balance.
- `parentName` equals the parent the nesting gives on every member (175 of 175, 39 of 39). Shared members appear
  under each of their parents with `dataStorage` "Shared" (FinMax TimePeriods 38, FinPlan Period 12).
- Values seen: `dataStorage` Store Data, Shared, Never Share, Label Only, Dynamic Calc, Dynamic Calc and Store;
  `accountType` (FinPlan Account) Revenue 1319, Liability 35, Asset 32, Equity 4, Expense 3, Non expense 2;
  `timeBalance` None 1019, Non Last 274, Average 102.
- Size and time without a projection: 690 to 1300 bytes a member; 20 dimensions read (all but Projects and Reserves)
  in 10.7 s (plan1) and 2.2 s (stratmodel); FinMax Products 25.6 MB in 2.5 s. Projects (140 000 members) would be
  about 110 to 180 MB: over `MAX_ANSWER_BYTES` (64 MB), so it was not sent without a projection.
- With `?fields=name,alias,dataStorage,usedIn,accountType,timeBalance,children` (the rule's one query item): the
  same tree with only those keys (`accountType` / `timeBalance` only where the dimension has them), 87 to 172 bytes a
  member. **Every dimension of both pods: 7.3 s (plan1, 14 dimensions, 25.1 MB in all) and 1.1 s (stratmodel, 8
  dimensions).** The largest: FinMax Projects 139 782 members, 16.8 MB, 2.1 s (+0.3 s to parse); Reserves 104 777,
  11.6 MB, 1.4 s.

### What today's rows need, and whether the answer gives it

Rows rebuilt from the answer by the old code itself (`ops.tree_from_labels`, which derives the name from the label,
depth and parent from indentation, and reverses an ancestors-at-the-bottom display), fed with what Smart View's zoom
would display: each member's **label = its Default alias, else its name**; **ancestors at the bottom** (today's atlas
has the siblings of every dimension in reverse outline order - in all 13 FinMax and 2 FinPlan period parents the
second month comes before the first - which is the zoom's display read bottom-up); left out: the 8 Dynamic Time Series
members (`H-T-D` .. `D-T-D`, the first 8 children of FinMax TimePeriods and of FinPlan Period: the zoom does not show
them) and the members whose `usedIn` does not name the plan type (2 in FinPlan Account). FinLite's 10 dimensions are
FinMax's application dimensions: the same answers serve both plan types (`usedIn` names FinMax and FinLite).

Compared with the copy of today's atlas, per connection and dimension (32 in all), on the keys `(name, parent)`
after `INSERT OR IGNORE` (as `atlas.store` keeps them):

| Result | Dimensions |
| --- | --- |
| keys, labels, depths and orders all equal, and the row count and `build.n` equal | 20: FinMax and FinLite BalanceTypes, Currencies, Functions (FinMax), GLSources, Products, Scenarios, TimePeriods; FinPlan Account (1391 / 1383 rows, the 8 `INSERT OR IGNORE` duplicates and the 2 plan-type exclusions included), BalanceTypes, Products, Period, Scenario, Version, Years |
| keys, depths and orders equal; 2 labels differ | FinMax and FinLite Accounts (an alias changed since: the accepted class of Task 3) |
| members added / removed since; every common key's depth and label equal, the relative order of the common keys equal | FinMax and FinLite Counterparts (+31 / -1), Entities (+41 / -2), LegalEntities (+31 / -1); FinMax Systems (+4), Reserves (+1, 104 776 / 104 772 rows); FinPlan Entity (+6) |
| members added / removed since, 15 labels differ, relative order of the common keys NOT equal | FinMax Projects (+61 / -2): looked at in the gate |

So the answer carries everything the rows need: the hierarchy (nesting; `parentName` agrees), the label (alias, else
name), the order (outline order, reversed as today), shared members under each parent, a name used twice (the 8
`INSERT OR IGNORE` duplicates of FinPlan Account and the 2 of Projects / 4 of Reserves reproduced), the plan type
(`usedIn`), and the extra columns (real name, alias, data storage, account type and time balance where the
dimension has them).

### Candidates 2 and 3: not needed

Candidate 1 answered, so the plan-type "Get Dimension Details" (Oracle: immediate children only), "Evaluate member
selection members" (a POST with a member-function body) and the provider XML's `EnumMembers` / zoom were not probed
and nothing was added to `svxml.check_read_only`.

### Requests sent

Get Member: 2 + 2 + 20 without a projection, 22 with it (46 GETs, 25 of them on plan1's 14 dimensions); the plan-type
dimension list (already allowlisted in Task 2): 6 GETs. Nothing else.

(The commit that first recorded this section, 147117c, said 22 of 32 in its message: 20 is the count.)

### Verdict: go on with Get Member

A route returns the hierarchy with what the atlas rows need, for every dimension, in about 8.4 s for both pods
together (today's build table records 133 s of hidden-Excel zooms, Excel's start not counted). Step 1 builds on it.

### Not used yet: EPM's account type and time balance (controller ruling, 2026-11-22)

Get Member carries `accountType` and `timeBalance` on account-type dimensions only (FinPlan Account; FinMax
BalanceTypes - FinMax's Accounts dimension has neither). They are stored in the atlas (since the fix round: its side
table `member_real`, columns `account_type`, `time_balance`) and **nothing reads them in this task**: the behaviour
cards keep the account type and time balance derived from data, exactly as before. FinPlan Account, EPM's value
against today's card (counts, one GET, read-only):

| EPM `accountType` / card `account_type` | members |
| --- | --- |
| Revenue / statistic | 894 |
| Revenue / None | 147 |
| Revenue / revenue | 76 |
| Revenue / expense | 47 |
| Revenue / balance | 1 |
| Liability / balance | 35 |
| Asset / balance | 32 |
| Equity / balance | 2 |
| Non expense / statistic | 2 |
| (not in Get Member's tree) / statistic, balance, None | 16, 2, 5 |

| EPM `timeBalance` / card `time_balance` | members |
| --- | --- |
| None / flow | 607 |
| None / None | 294 |
| None / balance_last | 1 |
| Non Last / balance_last | 131 |
| Non Last / None | 109 |
| Average / average | 62 |
| Average / None | 32 |
| (not in Get Member's tree) / None, flow | 22, 1 |

EPM's account type is not the cards' concept (FinPlan types nearly every account Revenue, statistics and costs
included). Its time balance agrees wherever the data decided (1 disagreement); using it would newly decide about 435
accounts the data left undecided (294 + 109 + 32) - which changes how the resolver treats them (levels or flows). A
decision for the owner later, not for this phase.

### The success status

Get Member answers **201** (Oracle documents it as this GET's "historical" success status); the client
(`members.read_tree`) accepts exactly 201 - a 200 or any other status is an error naming the status.

## Gate (Step 2), 2026-11-22 03:25-03:35 UTC: passed

Script: `scripts/compare_atlas.py <work folder> all` (deleted after the phase's last gate run on the release head), run
from `engine/` at commit a45bda9. It copies today's atlas (read-only) into a work folder in the scratch area, builds
over REST into a copy there (`atlas_build.build_into` + `build_rest`, as `smartass atlas build` does) and points every
consumer check at those copies (`config.ATLAS_DB` / `KNOWLEDGE_DB` are set before anything opens a database; a work
folder in the smartass home is refused). The owner's `atlas.db` and `knowledge.db` were not written (their modification
time is still 2026-11-10). Every line is the script's own output, shortened.

| Part | Result |
| --- | --- |
| Build | 32 dimensions, 0 errors, swapped in: 22.0 s in all (copying, SQLite backup and storing included); Get Member **9.7 s for plan1** (FinMax and FinLite: 14 trees, each read once) and **0.9 s for stratmodel** (8) |
| Compare | 20 of 32 dimensions identical (keys, depths, labels, order, row count, `build.n`); the rest only in the classes below; **real differences: 0** |
| Labels | the alias-changed class proven live on the provider route (12 of 12 samples: the rebuilt label is EPM's; today's label is refused) |
| Knowledge | knowledge.db from each atlas: built, 0 errors, 0 warnings, every count equal (37 metrics, 45 facts, 88 metric members, 68 794 cards ...) except `member_fts` 54 275 -> 54 359 (+84: the entities added since, FinMax 39 + FinLite 39 + FinPlan 6) |
| Offline test set | golden numbers: 9 of 9 passed with today's atlas, **9 of 9 with the rebuilt one** |
| Stored specs | 460 stored specs of answers.db, all planned; 18 782 members asked, 18 771 resolved with either atlas; the same 11 missed by both (FinPlan Account 6, Entity 5: label-keyed members the stored answers name and EPM no longer has - Task 3b's class); missed only with the rebuilt atlas: none |
| grid_smoke | b01 prepare OK 4.1 s, b07 OK 1.1 s, b16 OK 14.1 s (1 line set) |

### The classes (per connection and dimension)

- **Added since** (counts; codes when fewer than 20; a sample proven after the gate run: 11 of 11 added code-keyed
  members across Counterparts, Entities, Legal Entities, Projects, Systems and Reserves are shown by the provider route
  under their rebuilt label): Counterparts +30 (FinMax and FinLite each), Entities +39, Legal
  Entities +30, Projects +59, Reserves +1 (`re.SN.0017960393`), Systems +4 (`sy.S05666`, `sy.S05795`, `sy.S05796`,
  `sy.S05797`), FinPlan Entity +6 (label-keyed, not named). **Removed since**: 0.
- **Re-parented since** (one code removed under one parent, added under another; its subtree keeps its keys, its
  place in the order and its depth follow it): `un.SY.G332` (Entities, FinMax and FinLite), `pj.L1.L1.L1PBA6` and
  `pj.L8.L8.L8PBA6` (Projects). The 24 Projects keys whose relative order differs are all below these two (0
  elsewhere); every parent's children are in the same order in both atlases.
- **Alias changed since**: code-keyed (the same key, another label): `ac.24074` (FinMax and FinLite Accounts, under 2
  parents) and 15 projects: `pj.BR.BR.BRZ142`, `pj.BR.BR.BRZ186`, `pj.SY.SY.SY4Y15`, `pj.SY.SY.SY5Y13`,
  `pj.SY.SY.SY5Y15`, `pj.SY.SY.SY5Z22`, `pj.SY.SY.SY5Z24`, `pj.SY.SY.SY6P74`, `pj.SY.SY.SY7Z39`,
  `pj.SY.SY.SY9Z11`, `pj.SY.SY.SYHA29`, `pj.SY.SY.SYLA18`, `pj.SY.SY.SYNA13`, `pj.SY.SY.SYRZ06`,
  `pj.SY.SY.SYRZ15` - 16 members. Task 3 counted 6 stale aliases (its report: `ac.24074` and 5 of 18 028 FinMax
  entities, measured over the real-named members it fetched); `ac.24074` is in this class; the 5 entities could not be
  matched by code (Task 3 lists none, and this gate finds no code-keyed entity whose label changed - Entities differ
  only by 39 added, 1 re-parented and 1 label-keyed rename).
  Label-keyed (today's key was a label without the "code alias (prefix)" form; its new alias has it, so its key is
  now its code): 1 each in Counterparts, Entities and Legal Entities (FinMax and FinLite), paired with today's key by
  place (parents, depth and children). Live, by asking the provider route for each sample by its label: the rebuilt
  label is EPM's (12 of 12) and today's label is refused (12 of 12).
- Depth differences: 0. Label differences outside the alias class: 0.

### Found on the way (not changed here)

- The atlas's keys derived from a "code alias (prefix)" label are not always EPM's member names: in the rebuilt atlas
  139 634 of 139 777 Projects keys (`pj.SY.SY.X` for EPM's `pj.SY.X`), 1 437 of 5 609 Systems keys, 14 Reserves, 5
  Entities, 4 GL Sources, 1 Counterparts, 1 Legal Entities and all 25 code-keyed FinPlan Account keys differ from the
  real name. The provider route takes such a key only where EPM resolves it (a project control answered; the 15
  alias-changed projects were refused by key as "Grid is Invalid" and answered by label). The keys stay as they are
  (the contract of this task); the REST engine now matches these members by their real name (next point).
- The exact-name check (Task 3's review, Minor 2; commit a45bda9): with the rebuilt atlas, Task 3's live comparison
  (`scripts/compare_grid_routes.py` with config pointed at the copies) passed - order cases 4 of 4, the 21 real grids
  and the split equal on both routes, the wrong-member class as in Task 3b (2 left out, 4 refused with the clear
  error).

### Verdict: passed

The gate ran at a45bda9. Commits after it: 276247a (a message string), b11bbd3 (the REST engine accepts either real
name of a key used twice - live: 5 FinPlan account keys, 1 entity key), 0f62eb8 (the swap's wait for a lock bounded
to 30 s), 9e21637 (test guards); none changes what the gate compares.

`smartass atlas build` builds over REST by default (commit 9de44d6); `--excel` keeps the hidden Excel's zoom until
Task 6. The owner's `atlas.db` is replaced by a rebuilt one only by the controller, after review. (The concern
recorded here at first - an older release's positional INSERT failing on a member table with new columns - is
resolved by the fix round below: the member table keeps its seven columns.)

## Fix round (Task 4b) and the gates re-run, 2026-11-22 06:03-06:16 UTC: passed

What changed after the review (each its own commit): the real name, alias, data storage, account type and time
balance left the `member` table for a side table `member_real` (keyed connection, dimension, key, parent, real name;
every real member behind a key; stamped with its build) - `member` has exactly the seven columns every release knows
again (749574b); a dimension EPM lists without an allowlist rule is skipped, its rows kept (ef4dc5d); a build that
replaced nothing says so and fails, `atlas status` shows the last attempt (ab447b0); one build at a time (4099ab7);
the X-T-D filter applies to the period dimension only - the plan type's list says dimType 'Period' for TimePeriods and
Period (aaab1d3); the test guards (71f1eb3); line lengths (e306f02).

Script: `scripts/compare_atlas.py <work folder> all` at commit a695d72, from `engine/`, the same work-folder rules as
above (the owner's files untouched). Every line is the script's own output, shortened.

| Part | Result |
| --- | --- |
| Build | 32 dimensions, 0 errors, 0 skipped, swapped in: 23.4 s in all; Get Member 10.9 s plan1, 1.1 s stratmodel |
| Schema | the member table has exactly the seven columns; 376 753 member rows, 376 767 side rows, all trusted; 11 (key, parent) pairs have more than one real member behind them (the twins the member table keeps once) |
| An older release on this atlas | release 0.6.0's purge and positional 7-value INSERT of FinPlan on a copy of the rebuilt atlas: worked, rows as before; FinPlan's side rows no longer trusted (0 keys), FinMax's still trusted |
| Compare | the same as at a45bda9: 20 of 32 identical, 0 real differences; added 268, removed 0, re-parented 4, alias changed 16 members (code-keyed) + 3 label-keyed renames per FinMax and FinLite, 24 out-of-order keys all below a re-parented member. The period-only X-T-D filter changed no row |
| Labels | 12 of 12 samples: the rebuilt label is EPM's, today's is refused |
| Knowledge | equal but `member_fts` 54 275 -> 54 359 (the added entities) |
| Offline test set | 9 of 9 golden with either atlas |
| Stored specs | 460 specs, 18 771 of 18 782 members resolved with either atlas, the same 11 missed by both |
| grid_smoke | b01 OK 2.7 s, b07 OK 0.8 s, b16 OK 18.6 s |
| Both engines (`compare_grid_routes.py grids` on the rebuilt atlas) | 21 real grids + the split: all equal (22 OK); the drop probes OK on FinMax and FinPlan; wrong members: 2 left out with a warning, 4 refused with the clear error (one FinPlan entity, not in the atlas) - as in Task 3b; recorded in docs/audit/2026-11-21-rest-vs-provider.md too |

### How an older release and this one share one atlas file

The `member`, `build` and `card` tables are exactly what release 0.6.0 reads and writes. This release adds only
`member_real`, which 0.6.0 never opens. 0.6.0's `atlas build` / `knowledge refresh` on an atlas this release built:
its purge and positional INSERT work (shown above on a copy); the side rows it leaves carry the old build's stamp,
and this release's REST engine then ignores them and uses exactly the check before a45bda9 (key or place) until the
next REST build writes them again. This release's reads never add a table; only a store of real names creates
`member_real`, and a store without them (the hidden-Excel route, until Task 6) clears that dimension's side rows.

### Deviations and open items

- The alias class is proven by asking the provider route by LABEL, not by code (the review's Minor 4): a code-keyed
  label carries its code, and the proof shows the rebuilt label is a live alias while today's is refused.
- Order of commits (the review's Minor 5): the default switch (9de44d6) was committed before the first gate run
  (a45bda9); this re-run at a695d72 covers every later change.
- Task 3's 5 entity "stale aliases" (the review's question (c)): likely the 5 Entities keys that are not EPM's real
  names (Task 3 probably counted a key / real-name mismatch); a question for the controller, nothing depends on it.
- Open, not for this phase (the review's Minor 10): `rest.py` still SENDS the atlas key in the exportdataslice body;
  sending a trusted real name instead would stop depending on EPM resolving keys like `pj.SY.SY.X`.
