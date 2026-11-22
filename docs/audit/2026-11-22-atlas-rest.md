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
| keys, labels, depths and orders all equal, and the row count and `build.n` equal | 22: FinMax and FinLite BalanceTypes, Currencies, Functions (FinMax), GLSources, Products, Scenarios, TimePeriods; FinPlan Account (1391 / 1383 rows, the 8 `INSERT OR IGNORE` duplicates and the 2 plan-type exclusions included), BalanceTypes, Products, Period, Scenario, Version, Years |
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

### Verdict: go on with Get Member

A route returns the hierarchy with what the atlas rows need, for every dimension, in about 8.4 s for both pods
together (today's build table records 133 s of hidden-Excel zooms, Excel's start not counted). Step 1 builds on it.
