# Light and fast, phase 2: the second engine over REST, the atlas over REST, the hidden Excel gone

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan
> task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** smartass fetches everything from Oracle EPM over one token session per pod. A grid the Smart View
provider route cannot fetch is fetched once over Oracle's documented Planning REST API instead of through Smart
View in a hidden Excel; the member catalogue (the atlas) is built over REST in about half a minute; and the
engine's own hidden Excel, the VBA bridge and the Smart View sign-in handling are deleted. The COM job runner for
the **user's** Excel (turns, edit mode, the answer job, the `xl_*` tools) stays.

**Spec:** `docs/specs/2026-11-19-light-fast-design.md`, section "1. EPM access: one token session per pod" and the
"Order of work" row for phase 2. **Proof before the old part is deleted:** 20 grids from the run log fetched on
both routes with identical `Fetched`; a rebuilt atlas equal to today's apart from members added since.

**What exists (0.6.0):** `epm/oauth.py` (device-code sign-in, the stored refresh token), `epm/epmhttp.py`
(`EpmSession`: the allowlist - method, path, query names, header names, body shapes; the bearer only over https to
the pod's own origin; `RequestLost`), `epm/svhttp.py` (`BearerTransport`: the provider XML with the bearer;
`Fetched`), `epm/drill_token.py`, `data/adhoc.py` + `xl/pull_adhoc.py` (the Excel fallback for a grid),
`epm/ops.py` (its bridge half), `xl/host.py` / `xl/hostproc.py` (the hidden Excel and the user-job runner in one
class), `xl/dialogs.py`, `xl/login.py`, `epm/sso.py`, `engine/sso_watcher.ps1`, `engine/vba/`.

## Global constraints

- Read-only towards EPM: every request goes through `EpmSession`'s allowlist; the only new requests are
  `POST .../plantypes/{cube}/exportdataslice` (a read: it exports a slice) and `GET`s under
  `/HyperionPlanning/rest/v3/applications/`; each new rule is as narrow as the others (query names, header names,
  body shape) and has attack tests. Nothing that imports, clears, runs a rule or a job is ever allowlisted.
- A token, a session key, a cookie or a URL query string never appears in a log line, an exception text, a tool
  reply, an audit file or a test fixture. The bearer goes only over https to the origin of a pod with a scope.
- No figure and no business text from EPM is printed or written by a script or into an audit file: counts,
  booleans, seconds, member CODES of the test cells, key paths and value types only.
- Same numbers: the offline test set's golden numbers unchanged; `checks.sheet_equals_model` on every answer.
- No change to tool names, tool parameters or reply fields, except that `routes` may name the new route and a
  warning says when the second engine was used.
- The old route is deleted only after its task's live comparison is recorded as passed in `docs/audit/`.
- Live runs: read-only against EPM from the owner's PC, without Excel, are part of the tasks that say so; anything
  in the owner's Excel needs his explicit answer for that run (controller only).
- Tests never reach EPM, the identity provider or Excel. Line length 120; ruff and the suite green before every
  commit; commits end with the session's `Co-Authored-By` trailer.

---

### Task 1: The provider route rides on `EpmSession`

One security boundary instead of two. `BearerTransport` builds its own request today; the allowlist already holds
the Smart View POST rule (pinned as uncalled by a test since 0.6.0).

**Files:** `engine/smartass/epm/svhttp.py`, `epm/epmhttp.py` (only if the rule needs a header or query name the
provider really sends), tests `tests/test_svhttp.py`, `tests/test_epmhttp.py`.

- `BearerTransport` sends through `EpmSession.request` (the body admitted by `svxml.check_read_only` as today);
  its own origin / https checks (0.6.0, Task F) become the session's; a 401 is the session's one retry;
  `RequestLost` maps to today's error for a lost request. Same `Fetched`, same timings keys.
- The test that pinned "nothing sends the Smart View rule" flips to "the provider route sends it and nothing else
  does".
- [ ] Failing tests first (the transport's requests pass through the session's allowlist: a provider body the
  read-only check refuses never leaves; the bearer never goes to another origin), then the change, the suite.
- [ ] Live, read-only, no Excel: `scratch` script as in 0.6.0's smoke - three test-set questions prepared; ok,
  seconds. Commit `refactor(epm): the provider route sends through EpmSession`.

### Task 2: `epm/rest.py` - a grid over `exportdataslice`

**Files:** create `engine/smartass/epm/rest.py`, `engine/tests/test_rest.py`; modify `epm/epmhttp.py` (two rules).

**Interface:** `fetch(session, request: GridRequest, atlas) -> Fetched` with `route = "rest"`: the same `Fetched`
shape the provider route returns, so everything after the fetch is unchanged.
- The request: `gridDefinition` with `pov`, `columns`, `rows` built from the `GridRequest`; dimensions named as
  REST names them (without "All " and without spaces: derive the mapping from the atlas / the application's
  dimension list, never by string surgery alone - test the FinMax, FinLite and FinPlan dimension names);
  `suppressMissingBlocks` and the other options chosen so the answer has the same cells as the provider's.
- The answer carries member NAMES, not labels: labels come from the atlas (the same labels the provider route
  returns; a member the atlas does not know keeps its name and adds a warning).
- A grid over the cell budget is split along rows into requests of at most N cells and stitched (N a constant,
  default 100 000, to be confirmed live; no cell cap is documented).
- The grid is accepted only when the same layout check the provider route applies passes.
- Errors: 400 with Oracle's message -> the provider route's error class for a refused grid (the message scrubbed
  of anything secret); 401 -> the session's retry; 404 on the plan type -> a clear configuration error.
- Allowlist: `POST /HyperionPlanning/rest/v3/applications/{app}/plantypes/{cube}/exportdataslice` (app and cube
  from config only, exact segments, no extra path; JSON body of the known keys; a size cap) and
  `GET /HyperionPlanning/rest/v3/applications/{app}/...` for exactly the reads Tasks 2 and 4 need.
- [ ] Failing tests with recorded-shape fakes (build them from Oracle's documented response shape; no live data in
  fixtures): the mapping of names, labels from the atlas, the split and stitch, each error, the attack tests for
  the two new rules. Then `rest.py`. Suite. Commit `feat(epm): grids over Planning REST exportdataslice`.

### Task 3: The second engine replaces the Excel fallback

**Files:** `engine/smartass/server.py`, `pipeline.py`, `data/adhoc.py` (deleted at the end of this task),
`epm/drillstep.py` (`_link` / `_child_values`: their Excel fallback), `scripts/compare_grid_routes.py` (new,
deleted in Task 7), `docs/audit/<date>-rest-vs-provider.md`.

- Where a grid the provider route could not fetch falls back to Smart View in Excel today (`adhoc_fetch`), it is
  fetched once over REST; `routes` says `"rest"` and one warning names the fallback. `drillstep._link` and
  `_child_values` lose their Excel fallback the same way (what they need over REST, or a clear error).
- **The gate (live, read-only, no Excel):** `scripts/compare_grid_routes.py` takes 20 different grids from the run
  log (`runs.jsonl`: different cubes, shapes, sizes, at least one large enough to be split, one with a missing
  member, one FinPlan) and fetches each on both routes: identical `Fetched` (cells, labels, order) within the
  provider's own number precision; per grid one line of counts, booleans and seconds; exit 1 on any difference.
  The audit file records it. If a class of grids differs and cannot be made equal, STOP: the Excel fallback is
  not deleted and the controller decides.
- [ ] Tests (fakes): the fallback order, the reply's `routes` and warning, no Excel call on that path. [ ] The
  live comparison and the audit. [ ] Only then delete `data/adhoc.py` and what only it used. Commit.

### Task 4: The atlas over REST

**Files:** `engine/smartass/atlas/` (the build), `epm/rest.py` (members), `ops/cli.py` (`smartass atlas build`),
`knowledge/cards_fetch.py` (account type and time balance straight from the answer), tests,
`scripts/compare_atlas.py` (new, deleted in Task 7), `docs/audit/<date>-atlas-rest.md`.

- `GET .../applications/{app}/dimensions/{dim}/members/{dim}` returns a whole hierarchy with alias, parent, data
  storage, account type and time balance. `smartass atlas build` uses it for every dimension of every
  application; the consolidation operator is not in the answer, so `cards_fetch` keeps deriving it from data.
- Oracle documents this endpoint for Service Administrators only: when it answers 401 / 403 the build walks the
  hierarchy with the provider's `EnumMembers` instead (over `EpmSession`) and says so.
- The build writes a NEW database file and swaps it in only when complete (never a half-built atlas in place);
  a tool call without an atlas says which command to run.
- **The gate (live, read-only, no Excel):** a rebuilt atlas in a temp file compared with today's `atlas.db`:
  member counts per dimension, the set of member codes, parents and aliases equal apart from members added or
  removed since (listed by count, and by code when fewer than 20); build seconds (goal: under a minute per pod).
- [ ] Tests (fakes) [ ] the live comparison and the audit [ ] commit.

### Task 5: The user-Excel job runner cut loose from the hidden Excel

**Files:** `engine/smartass/xl/host.py`, `xl/hostproc.py`, `xl/jobs.py`, `core/serverkit.py`, callers, tests.

`ExcelHost` is two things in one class: the engine's own hidden Excel (started on demand, with the Smart View
bridge and the sign-in handling) and the runner of jobs in the USER's Excel (a thread per job, the deadline, the
abandon path, the release of a job's objects - hardened in 0.5.0). Separate them without changing the second:
- a `UserJobs` runner (name free) with exactly today's behaviour for user jobs - every test of it keeps passing
  unchanged (the turn, the given-up mark, the release on the job's own thread, the first-contact guard);
- nothing on the path of `fin_analyze`, the follow-up, `fin_refresh`, `fin_drill`, `fin_status`, the `xl_*` tools
  or the workbench needs the hidden Excel any more (after Tasks 1-4 it is only reached by the Smart View
  fallback and the old atlas build): prove it with a test that fails the suite if anything starts it.
- [ ] Characterisation tests first where the split cuts through untested behaviour; then the split; the suite;
  `scripts/check_inproc.py` and `scripts/check_insert.py` (hidden Excel of the builder's own) exit 0.
- [ ] Controller, with the owner's go: `scripts/check_followups.py` through both executors in the owner's Excel.

### Task 6: Delete the hidden Excel and what only it used

Only after Tasks 3 and 4's audits say passed and Task 5 is reviewed.
**Delete:** the hidden-Excel half of `xl/host.py` / `xl/hostproc.py`, `xl/dialogs.py`, `xl/login.py`,
`xl/pull_adhoc.py`, the bridge half of `epm/ops.py`, `epm/sso.py`, `engine/sso_watcher.ps1`, `engine/vba/` and the
bridge add-in's build (`xl/instances.py` `ensure_bridge_file` and the bridge attach in `xl/target.py` -
"ready + bridge" becomes "ready"), the Smart View global-option handling, the doctor checks for Excel-as-engine,
Smart View, VBA access, Smart View start-up; their tests; their docs.
- Grep every deleted name through engine, scripts, skills, docs, addin and the MCP instructions; state the test
  count before and after and list what went. `test_doctor` gets a test that nothing in the engine imports the
  deleted modules.
- [ ] Controller, with the owner's go: a live smoke in the owner's Excel (follow-ups; a lines answer); `smartass
  doctor` clean.

### Task 7: Docs, version, the scripts of this phase

`docs/DESIGN.md`, `docs/troubleshooting.md`, the skill's `reference/` files, `README.md`, `CHANGELOG.md` (what
changed for the user; the audits' figures; what a PC needs), the MCP instructions within their budgets; delete
`scripts/compare_grid_routes.py` and `scripts/compare_atlas.py` (their results stay in the audits); version 0.7.0.

## Also to test in this phase (from the research on Oracle's MCP landscape)

The newer agent-facing REST endpoints Oracle announced for 26.04 / 26.11: check which the two pods answer today
(plan1 is two updates behind stratmodel), read-only, and record in the audit what each returns by shape. No code
depends on them in this phase.

## As built (2026-11-22)

What differs from the tasks above. The ledger of every ruling is
`.superpowers/sdd/2026-11-21-light-fast-p2-rest/progress.md`; the gates are in `docs/audit/`.

- **Task 3 became 3, 3b and 3c.** Task 3's gate passed (the order proof, 20 real grids identical), but a member EPM no
  longer has failed the whole grid over REST, where the Excel fallback had dropped it silently. So the fallback was
  not deleted then. Task 3b leaves such a member out over REST when its axis has others, with one warning and one
  caveat that every total with it is INCOMPLETE; a POV member or an axis's only member is a clear error. The gate
  was re-run and passed; only then were `data/adhoc.py` and `xl/pull_adhoc.py` deleted. Task 3c was the review's
  fix round (the split loop, a dated Known-limits line, the drill link's next step by cause, no raw provider text in
  an error, the recorder refusing a run with a member left out).
- **Task 4 became 4, 4b and 4c, on a revised brief.** It started with a probe and a hard stop: Oracle's Get Member
  answered this non-administrator user for every dimension, so the provider's `EnumMembers` fallback was never
  built. The member keys and rows stay exactly as the Excel zoom made them. 4b moved the real names and properties
  from new columns into a stamped side table, `member_real`, so that an older installed release can still write the
  atlas; a dimension EPM adds is skipped instead of blocking the build; a build that replaces nothing says so. 4c
  made the REST engine's name check strict with a trusted side table. EPM's account type and time balance are stored,
  not used (an owner decision for later).
- **Tasks 5 and 6 ran on revised briefs.** Task 5 also took the review carry-ins of 3c, moved the engine's errors to a
  neutral module and pinned the user-job runner with characterisation tests before the split; it kept the two cold
  paths of a user job (the bridge add-in, a first workbook with no Excel) as a hand-off for Task 6. Task 6 replaced
  them (no bridge; a first workbook built without Excel or opened from the TN template in the user's Excel), and
  deleted the SSO watcher by ruling: the automatic "Telenor SSO" click is gone for every Excel.
- **Kept, though the plan's Task 6 deleted it:** `xl/dialogs.py`, slimmed to a dialog finder for busy errors and a
  guard that refuses at once to work in an Excel whose Smart View Login is open (COM into it blocks); `xl/hostproc.py`
  (a job's COM thread, and a one-time sweep of the hidden Excels an older release left, by its pid file); the doctor's
  `smartview_startup` check, reworded (Smart View's start-up Login blocks smartass in that Excel). New:
  `smartass doctor --fix` and the `leftovers` check for what an older release left.
- **Task 7 ran on a revised brief.** `scripts/compare_grid_routes.py` and `scripts/compare_atlas.py` were NOT
  deleted in it: code the gates cover changed after their run of 22.11 08:13 (local time), so the controller ran
  both gates again at 8049f2a (22.11 12:27-12:44, passed); Task 8 recorded that run in both audits ("Final run") and
  then deleted the scripts. Task 7 also removed the resolve's unused runner parameter that Task 6 left.
- **Task 8, not in the plan:** the last review findings (the scrub's narrower named keys, `_open_path`'s busy guard,
  the launch guard's aliases, one retry helper), the docs corrections and the final gate figures.
- **Task 9, not in the plan: the final fix round** of the two whole-branch reviews: the relay's stop of a server
  spares an Excel and what runs under it; a cold first workbook opened read-only is a clear error; the provider's
  error is classified on its raw text; Smart View's read-only check refuses CDATA and comments; the REST warning's
  quoted error is scrubbed; the scrub narrower on ordinary text and catching the shapes still found; REST reads only
  plain numbers; setup keeps an atlas with real names over a release copy without them, and the doctor warns on an
  atlas without them; the test guards raise a `BaseException`; `scripts/probe_bearer.py` deleted; the docs (an
  "Upgrading from 0.6.0" section, the open items). The controller runs the routes comparison once more after it.
- **Live steps still owed (before the merge, each with the owner's go):** Task 5's `scripts/check_followups.py`
  through both executors in the owner's Excel, `scripts/bench_answers.py`, and Task 6's live smoke with the first
  workbook when no Excel is open (the cold path). `scripts/check_insert.py` ran in an Excel of its own (188 of 188
  checks, 22.11.2026 at 8049f2a).
- **Not done in this phase:** the plan's "Also to test" (Oracle's newer agent-facing REST endpoints of 26.04 /
  26.11): no probe was run, and nothing depends on them.
