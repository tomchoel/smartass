# smartass v2 · Plan 2A: semantic layer, caveats, monthly refresh

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the interim knowledge (`knowledge/interim.py` + `AtlasResolver`) with the semantic layer of
sub-spec 2 (`docs/specs/2026-11-04-v2-2-knowledge-design.md` §1–§4, §8):

- **Typed files.** Metrics and facts as small TOML files in git, compiled with the atlas into `knowledge.db`.
- **Lookup.** One `KnowledgeResolver` serves the model, and `fin_lookup` searches aliases.
- **Caveats.** Every reply carries the caveats for the members, metrics and periods it touched.
- **Monthly refresh.** `smartass knowledge refresh` produces an impact report and leaves drafts for Tom's one-time
  approval with `smartass knowledge approve`.

**Out of scope (Plan 2B):**
- the critic and rubric
- playbooks as MCP prompts
- answer memory
- product concepts (`splitt`, `fam_*`), which need the Product dimension in queries

**Known limit:** the atlas stores no consolidation operators, so metric identities (C70 = C50S + C60) are checked
from `catalog.toml [checks]` against data in the test set, not against the tree.

**Conventions:**
- Branch `v2/2a` from `main`. Git Bash from `/c/dev/smartass/engine`.
- TDD: write the test first, see it fail, implement, see it pass.
- Never run python through a stdin heredoc; bare python only with `< /dev/null`.
- Stage files explicitly.
- Commit messages end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- The gates after every task: `uv run pytest -q`, `uv run ruff check .`, and
  `uv run python -m smartass.testset run < /dev/null` (8/8, goldens unchanged).

## File layout

| Path | What |
|---|---|
| `knowledge/semantic/metrics/<id>.toml` | One metric: name, aliases, description, unit, sign, balance, aggregation, materiality, status, approved, `[finmax]`/`[finlite]`/`[finplan]` members, `differs`, `[[valid]]` |
| `knowledge/semantic/facts/<id>.toml` | One fact: text, scope (members and/or metrics, per cube), valid_from/valid_to (`YYYY-MM` or empty), severity (info / warn / block), source, status |
| `engine/smartass/knowledge/semantic.py` | Load and validate the TOML files (pydantic): `Metric`, `Fact`, `load_semantic(dir)` |
| `engine/smartass/knowledge/seed.py` | One-off converter from catalog.toml, capabilities.toml and the maps into semantic files (kept for reruns) |
| `engine/smartass/knowledge/build.py` | Compile atlas + semantic into `knowledge.db` (tables metric, metric_member, alias + FTS5, fact, fact_scope, meta); build checks |
| `engine/smartass/knowledge/resolver.py` | `KnowledgeResolver` (implements `model.resolver.Resolver` + `caveats()` + `search()`) |
| `engine/smartass/knowledge/refresh.py` | Monthly refresh: atlas diff, impact report, drafts; `approve` |
| `engine/smartass/ops/cli.py` | `smartass knowledge build | refresh | approve <id|all> | status` |

`knowledge.db` lives in `%LOCALAPPDATA%\smartass\knowledge.db`. It is build output, never in git.

---

### Task 0: Branch

- [ ] `cd /c/dev/smartass && git switch main && git pull --ff-only && git switch -c v2/2a`

### Task 1: Semantic schema and loader

`semantic.py`, with pydantic models (`extra="forbid"`):

```python
class Members(BaseModel):            # per cube
    members: list[str]

class Metric(BaseModel):
    id: str                          # file stem, [a-z0-9_]+
    name: str
    aliases: list[str] = []
    description: str
    unit: Literal["NOK", "FTE", "count", "ratio"] = "NOK"
    sign: Literal["income", "cost", "neutral"] = "neutral"
    balance: Literal["Periodic", "Accumulated"] = "Periodic"
    aggregation: Literal["sum", "last"] = "sum"
    materiality: tuple[float, float] = (5_000_000.0, 0.10)
    status: Literal["approved", "draft"] = "draft"
    approved: dict | None = None     # {"by", "on"}
    finmax: Members | None = None
    finlite: Members | None = None
    finplan: Members | None = None
    differs: str = ""
    valid_from: str = ""             # "YYYY-MM"; empty = always
    valid_to: str = ""

class Fact(BaseModel):
    id: str
    text: str
    scope_members: dict[str, list[str]] = {}   # cube -> members
    scope_metrics: list[str] = []              # metric ids
    valid_from: str = ""
    valid_to: str = ""
    severity: Literal["info", "warn", "block"] = "info"
    source: str
    status: Literal["approved", "draft"] = "draft"
```

- **Validators:**
  - Accumulated with sum is refused, as in `MetricDef`.
  - `valid_*` must match `^\d{4}-\d{2}$` or be empty.
  - A metric needs at least one cube.
  - A fact needs at least one scope.
- **`load_semantic(dir) -> (list[Metric], list[Fact])`** reads `metrics/*.toml` and `facts/*.toml`. The `id` is the
  file stem. A duplicate id or an invalid file raises an error naming the file.
- **Tests** use `tmp_path` TOML files: a valid metric and fact load; Accumulated with sum is refused; a bad
  `valid_from` is refused; a duplicate id is refused.
- **Commit:** `feat(v2): semantic metric and fact schema`.

### Task 2: Seed converter and the seeded files

`seed.py` with `seed(knowledge_dir, out_dir)`. It writes TOML files. Run it once and commit the generated files.

- **Metrics.**
  - Each `catalog.toml [pnl]` line with a concept, joined across cubes on `concept`, as `interim.py` does now.
    That covers the members, FinLite = FinMax except stock lines, `SIGN` (move the dict into `seed.py`), stock as
    Accumulated with last, and materiality from `[materiality]`.
  - **Aliases** come from the `capabilities.toml` `[[measure]]` whose `concept` matches (its `synonyms`), plus
    the FinMax account code (e.g. "C70") and the FinPlan member name.
  - **Description** is the measure's `note`, or the catalogue label.
  - **Unit** is FTE for the stock line, otherwise NOK.
  - **Status** is approved with `approved = {by="tom", on="2026-10-15", via="catalog.toml"}`. The catalogue and
    the maps were approved on 15.10.
  - **`differs`** on metrics with a FinPlan member: "FinPlan groups accounts differently from FinMax (e.g.
    EBITDA 2025: 14.16 bn FinPlan vs 12.40 bn FinMax); compare across cubes only with a reconciliation line." This
    comes from the finplan_map meta notes.
- **Facts** (status approved, `source` naming the map and its approval):
  1. **Abolished entities.** One fact per catalogue `[[entity]]` with `abolished` / `successors`, scoped to that
     entity's FinMax and FinPlan members, `valid_from = "2026-06"`, severity warn. Text: "<label> was abolished on
     1 Jun 2026; successors <...>; history is restated to today's organisation."
  2. **Reorg moves, collapsed.** Group `org_map.json` rows with `moved` (e.g. "CWK → KYS") by the move pair, giving
     about 20 facts. Each is scoped to that group's FinMax members, `valid_from = "2026-06"`, severity info. Text:
     "Moved in the 1 Jun 2026 reorg (<from> → <to>); history is restated to the new organisation."
  3. **Account traps.** The `accounts_map.json` pitfalls other than the boilerplate "Ikke bokført i un.SY i 2025…",
     about 14 facts, scoped to the account member, severity warn. Keep the pitfall text, translated to English
     where it is Norwegian. The fact text is sheet and reply text, so it is English.
  4. **Revenue-bearing cost centres and adjustment units** (`org_map` pitfall patterns): one fact per pattern,
     scoped to the members, severity info.
  5. **FinPlan facts** (finplan_map meta notes), scoped to FinPlan metrics or members:
     - BANORWAY = Telenor Norge;
     - KPIs read Periodic (Accumulated sums the months);
     - actual volumes are not in FinPlan for BANORWAY;
     - Target carries no depreciation, so budget comparisons stop at EBITDA.
- **Tests.**
  - Running the seed into `tmp_path` from the real `knowledge/` gives at least 11 metrics (the 11
    `catalog_metrics()` names) and at least 30 facts, and every file loads with `load_semantic`.
  - The 2CON abolished fact exists with scope `un.SY.2CON` and `valid_from 2026-06`.
  - Metric "Opex" has the aliases "opex" and "C70", and Norwegian synonyms from capabilities.
- **Run it** into `knowledge/semantic/` and commit the files. They are reviewable in git and contain no amounts.
  Pitfall texts that quote NOK amounts are fine, since they come from approved maps already in git. Run
  `scripts/check_staged.py` on them.
- **Commit:** `feat(v2): seed the semantic layer from the approved catalogue and maps`.

### Task 3: Build `knowledge.db` with build checks

`build.py` with `build(semantic_dir, atlas_db, out_db) -> BuildReport`. It writes a fresh database and replaces the
old one atomically: write to `out_db.tmp`, then `os.replace`.

- **Tables:**
  - `meta(key, value)`: built_at, semantic hash, atlas built_at.
  - `metric(id, name, json)`.
  - `metric_member(metric_id, cube, member)`.
  - `alias(metric_id, text)` with an FTS5 table `alias_fts(text, metric_id)`. Also index metric names and member
    labels (from the atlas) for search.
  - `fact(id, json)`.
  - `fact_scope(fact_id, cube, member, metric_id)`.
- **Build checks** (`BuildReport.errors` stop the build; `warnings` do not):
  - **Error:** a metric member missing from the atlas for its cube.
  - **Error:** an alias used by two metrics. Matching ignores case.
  - **Warning:** a fact scope member missing from the atlas. It may be historical.
  - **Warning:** a draft metric or fact (counted).
- **Tests** run with a tiny atlas built in `tmp_path` (as `test_v2_atlas_resolver.py` does) and tiny semantic
  files:
  - a good build;
  - a missing member is an error and the old db is kept;
  - an alias clash is an error;
  - FTS search finds "driftskostnader" → opex.
- **CLI:** `smartass knowledge build` prints the report.
- **Commit:** `feat(v2): knowledge.db build with checks`.

### Task 4: `KnowledgeResolver` with caveats and search

`resolver.py`:
- **`KnowledgeResolver(kdb, atlas_db, forecast_year)`** implements every `Resolver` method:
  - Metrics come from `knowledge.db`. The lookup ignores case and matches the id, the name or an alias.
  - Tree, labels and `exists` come from the atlas, exactly as `AtlasResolver` (reuse its code by moving the shared
    atlas part into a helper).
  - `status` is the metric's status, so a draft metric is refused by the existing guard.
- **`caveats(touched: dict[cube, set[member]], metric_ids: set[str], months: set[Month]) -> list[dict]`**
  returns `{id, text, severity}` for every approved fact whose scope matches any touched member or metric and
  whose validity overlaps any touched month. The list is deduplicated and sorted block, then warn, then info.
  - Draft facts in scope add one warn: "unapproved knowledge change touches this answer".
- **`search(text, kind, cube, limit)`** searches FTS over aliases and names (metrics) and the atlas search (members).
  Every member hit carries its caveats (without month filtering).
- **Tests:** the resolver passes the existing resolver-protocol tests (copy them from `test_v2_atlas_resolver.py`
  onto the new class). A caveat for un.SY.2CON appears for Sep 26 but not for Jan 26 (`valid_from 2026-06`).
  Drafts give the warn line. Search finds by a Norwegian alias.
- **Commit:** `feat(v2): KnowledgeResolver with caveats and search`.

### Task 5: Wire it in

- **`pipeline.run_prepare` and `testset/runner._resolver_for`** use `KnowledgeResolver` when `knowledge.db`
  exists. They fall back to building it on the fly (`build()` into the home folder) when missing, and say so in the
  warnings.
  - Delete `interim.py` and `atlas_resolver.py` once nothing imports them. Keep the shared atlas helper.
- **Caveats in `fin_analyze` and `fin_drill` replies.**
  - `Prepared` collects the touched members per cube from its `GridRequest`s (entities and accounts), the metric
    ids from the spec queries, and the months from the requests' periods.
  - `reply["caveats"]` is `resolver.caveats(...)`.
  - The Answer sheet gets a "Caveats" section (note rows, wrapped like notes) when there are any.
  - Lines blocks contribute their cell's members.
- **`fin_lookup`** uses `resolver.search` and returns `{metrics: [...with aliases, description, differs],
  members: [... with caveats]}`.
- **`fin_status`** reports the knowledge build time, counts and drafts awaiting approval.
- **`ops/doctor` `_catalog` check** becomes `_knowledge`. It runs the build checks against the current atlas
  (FAIL on errors, WARN when the db is older than the atlas or has drafts).
- **Gates:**
  - The test set must stay 8/8.
  - The goldens' key_values must not change, because the metrics resolve to the same members. If a golden changes,
    stop and report: that would be a real change to investigate, not to approve.
- **Commit:** `feat(v2): knowledge layer wired into analyze, drill, lookup, status and doctor`.

### Task 6: Monthly refresh and approval

`refresh.py` + CLI:
- **`smartass knowledge refresh [--no-atlas]`:**
  1. Unless `--no-atlas`, rebuild the atlas through the same function as `smartass atlas build`. This is live EPM
     work, about 3 minutes.
  2. Diff the atlas against the snapshot of the last refresh (`knowledge/atlas_snapshot.json` in the home folder:
     member, parent, label per cube and dimension). The diff lists added, removed, moved (parent changed) and
     relabelled members.
  3. Draft proposals into `knowledge/semantic/` with `status = "draft"`:
     - a fact per moved or removed member that a metric or fact references ("<member> moved from <a> to <b> in
       <month>");
     - an added child of a metric's member gets no draft, because metrics use parents;
     - a removed metric member, which makes the build fail, gets a fact plus an entry in the report.
  4. Rebuild `knowledge.db`, run the offline test set, and write the report
     `%LOCALAPPDATA%\smartass\knowledge\reports\<yyyy-mm>.md` (the diff counts, metrics affected, the drafts
     written, build warnings, the test set result).
  5. Save the new snapshot.
- **`smartass knowledge approve <id|all> [--by tom]`** sets `status = approved` and `approved = {by, on}` in the
  TOML files, then rebuilds.
- **`smartass knowledge status`** prints the last build, the drafts and the last report path.
- **Due flag.** `fin_status` says "knowledge refresh due" when the closed month is newer than the last refresh
  month (stored in meta). The refresh itself is not run automatically inside a tool call.
- **Tests.**
  - A diff on two tiny snapshots finds a moved member.
  - A draft fact is written for a referenced moved member.
  - `approve` flips the status and the db rebuilds.
  - Use fakes for the atlas rebuild and the test set.
- **Commit:** `feat(v2): monthly knowledge refresh with impact report and approval`.

### Task 7: Docs, skill, cleanup

- Delete the v1-only advisor files `gate.md` and `findings.md`, and update `test_vendor.py` hashes or drop those
  entries. Keep `variance-commentary.md` (seed for Plan 2B) and `NOTICE.md` for it. Delete the
  `products_map.json.leaves.json` and `.q4.json` side files if nothing reads them (grep first).
- Skill `skills/smartass/SKILL.md`: mention `caveats` (read them and pass them on), aliases in `fin_lookup`, and the
  monthly refresh (`smartass knowledge refresh`, then `approve`). Keep it under the size limit test.
- `docs/DESIGN.md`: the knowledge section as built.
- CHANGELOG line.
- **Commit:** `docs(v2): knowledge layer docs and cleanup`.

### Task 8: Live and finish (controller)

1. `smartass knowledge build` against the real atlas. The build must have zero errors; report the warnings.
2. `fin_lookup("driftskostnader")` returns Opex. `fin_lookup("2CON")` returns the abolished caveat.
3. `fin_analyze` with a query on un.SY.2BUS YTD returns caveats, if any are in scope. Sheet = model OK.
4. `smartass knowledge refresh --no-atlas`: the report is written and the test set is 8/8.
5. Merge into main, push.
