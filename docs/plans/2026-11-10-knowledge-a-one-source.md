# Knowledge layer, phase A: one source, inherited caveats, freshness Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Every piece of approved knowledge lives in `knowledge/semantic/*` and reaches the model at answer time. Caveats reach the members above and below the one they name. A stale atlas is said out loud.

**Architecture:**
- Four new strict pydantic kinds join `metrics/`, `facts/` and `concepts/` in `semantic.py`: `Rule`, `Entity`, `Kpi`, `Vocab`, plus member `notes/`.
- `build.py` compiles them into new tables of `knowledge.db`. It expands every fact's scope through the atlas tree (`inherit`) and fails on contradictions.
- `KnowledgeResolver` serves them to `fin_lookup` (new kinds `entity`, `kpi`, `vocab`, `rule`) and to the model:
  - `series.py` takes its scenarios from the `scenarios` rule;
  - raw members take their materiality from the `materiality` rule.
- A one-off fold script converts `catalog.toml`, `capabilities.toml` and the three business maps once. Those sources are then deleted.

**Tech Stack:** Python 3.12, pydantic v2, sqlite3 (FTS5), tomllib, pytest, ruff, uv. Repo: `C:\dev\smartass`, engine in `engine/`.

**Spec:** `docs/specs/2026-11-10-knowledge-layer-design.md` (phase A = §2, §4, §6 build checks + freshness, §7).

**Refinements of the spec, decided while planning:**
- **Account bundles are plain metrics.**
  - A bundle such as "consultants" (C60080 + C60090) is exactly a `Metric`: aliases, FinMax/FinLite members and `sign = "cost"`.
  - So it goes to `metrics/`, and there is no `bundles/` folder.
- **The three maps become `notes/`.**
  - They hold a meaning and a booking pattern per member: 234 FinMax accounts, 1 478 entities and 654 FinPlan accounts.
  - These become `notes/*.toml` (one file per map) and are returned on member hits.
  - Entities and KPIs come from `catalog.toml`.
- **`inherit` defaults to `"self"`, not `"both"`.**
  - Every existing fact gets an explicit `inherit` in Task 3.
  - `up` climbs only to the nearest ancestor that a metric or an entity names (C40510 stops at C40S). That stops one account trap from riding up into every EBITDA answer.
- **Unchanged:**
  - `products_map.json`, `build_products_map.py`, `build_family_concepts.py` and `review.py` stay (they generate `concepts/`). `review.py` loses its other three areas.
  - The `Forretningskart *.xlsx` files stay as documents.

**Standing rules for whoever executes this:**
- **Running Python and tests:**
  - Never run Python from stdin (`python -` or a heredoc). Write a `.py` file and run it with `< /dev/null`.
  - Run tests from `engine/`: `uv run python -B -m pytest -q tests < /dev/null`.
  - Lint: `uv run ruff check . < /dev/null`.
  - Goldens: `uv run python -B -m smartass.testset run < /dev/null` (needs EPM; 9/9 today).
- **Before every commit:** `git checkout engine/uv.lock` (uv rewrites it).
- **Before every push:** ruff and the full pytest must be green. After the push, `gh run list --limit 1` must show success: CI runs on UTC Linux, so a test must not depend on the local time zone.
- **Goldens are Tom's to approve.** If a golden changes, stop and report the diff. Do not re-pin it.
- **Branch:** `v2/knowledge-a` from `main`. Merge to main and push at the end (Task 12).
- **Commit messages** end with:
  ```
  Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01MW9ANq493fHqddMk5ipVTq
  ```

---

## File structure

| File | Responsibility |
|---|---|
| `engine/smartass/knowledge/semantic.py` (modify) | `Fact.inherit`; new models `Rule`, `Entity`, `Kpi`, `Vocab`, `MemberNote`; loader `load_all()` |
| `engine/smartass/knowledge/expand.py` (create) | Fact scope expansion over the atlas (down / up-to-stop / both), pure functions on a sqlite connection |
| `engine/smartass/knowledge/build.py` (modify) | New tables, the expansion, new checks (one rule per kind, entity placement, KPI metrics, aliases unique across kinds), `atlas_built_at` rebuild trigger |
| `engine/smartass/knowledge/resolver.py` (modify) | caveats with `via`; `entity()`, `kpi()`, `vocab()`, `rule()`, `scenarios()`, raw-member materiality from the rule; search kinds `entity`, `kpi`, `vocab`, `rule`; member notes; freshness notes |
| `engine/smartass/model/resolver.py` (modify) | `Resolver.scenarios()` in the protocol; `DictResolver` reads the same rule file |
| `engine/smartass/model/series.py` (modify) | scenario table from `r.scenarios()` |
| `engine/smartass/server.py` (modify) | `fin_lookup` docstring + kinds; INSTRUCTIONS line on caveats/rules |
| `scripts/fold_knowledge.py` (create, deleted in Task 11) | one-off: catalog/capabilities/maps → TOML |
| `knowledge/semantic/{rules,entities,kpis,vocab,notes}/*.toml` (generated, then hand-checked) | the knowledge |
| `engine/tests/v2/test_v2_knowledge_a.py` (create) | all new unit tests of this plan |

---

### Task 1: Branch and baseline

**Files:** none

- [ ] **Step 1: Branch**

```bash
cd /c/dev/smartass && git checkout main && git pull -q && git checkout -b v2/knowledge-a
```

- [ ] **Step 2: Baseline**

Run: `cd engine && uv run python -B -m pytest -q tests < /dev/null | tail -1 && uv run ruff check . < /dev/null | tail -1`
Expected: `1252 passed` and `All checks passed!`

---

### Task 2: Fact `inherit` and scope expansion

**Files:**
- Create: `engine/smartass/knowledge/expand.py`
- Modify: `engine/smartass/knowledge/semantic.py` (class `Fact`, around line 90)
- Test: `engine/tests/v2/test_v2_knowledge_a.py`

- [ ] **Step 1: Write the failing tests**

Create `engine/tests/v2/test_v2_knowledge_a.py`:

```python
"""Knowledge layer phase A (docs/plans/2026-11-10-knowledge-a-one-source.md)."""
import sqlite3

import pytest

from smartass.knowledge import atlas
from smartass.knowledge.expand import expand
from smartass.knowledge.semantic import Fact

# A small FinMax tree: accounts EBITDA > C40S > C405 > C40510 ; entities un.SY > 2CON > CC1, CC2
TREE = [("FinMax", "All Accounts", "ac.EBITDA", "EBITDA", "", 1, 1),
        ("FinMax", "All Accounts", "ac.C48", "C48 Gross margin", "ac.EBITDA", 2, 1),
        ("FinMax", "All Accounts", "ac.C40S", "C40S COGS", "ac.C48", 3, 1),
        ("FinMax", "All Accounts", "ac.C405", "C405 Goods", "ac.C40S", 4, 1),
        ("FinMax", "All Accounts", "ac.C40510", "C40510 Use of goods", "ac.C405", 5, 1),
        ("FinMax", "All Entities", "un.SY", "SY Telenor Norge", "", 1, 1),
        ("FinMax", "All Entities", "un.SY.2CON", "2CON Consumer Norway", "un.SY", 2, 1),
        ("FinMax", "All Entities", "CC1", "CC1 unit", "un.SY.2CON", 3, 1),
        ("FinMax", "All Entities", "CC2", "CC2 unit", "un.SY.2CON", 3, 2)]


@pytest.fixture
def adb(tmp_path):
    p = tmp_path / "atlas.db"
    con = atlas.connect(p)
    con.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)", TREE)
    con.execute("INSERT INTO build VALUES ('FinMax','All Accounts',1790000000.0,9,1.0,NULL)")
    con.commit()
    yield con
    con.close()


def _fact(inherit, members):
    return Fact(id="f", text="t", source="test", status="approved", inherit=inherit,
                scope_members={"FinMax": members})


def test_inherit_self_is_the_member_only(adb):
    assert expand(_fact("self", ["un.SY.2CON"]), adb, stops={}) == {("FinMax", "un.SY.2CON"): "un.SY.2CON"}


def test_inherit_down_reaches_every_descendant(adb):
    got = expand(_fact("down", ["un.SY.2CON"]), adb, stops={})
    assert set(got) == {("FinMax", "un.SY.2CON"), ("FinMax", "CC1"), ("FinMax", "CC2")}
    assert got[("FinMax", "CC1")] == "un.SY.2CON"


def test_inherit_up_stops_at_the_nearest_named_ancestor(adb):
    got = expand(_fact("up", ["ac.C40510"]), adb, stops={"FinMax": {"ac.C40S", "ac.EBITDA"}})
    assert set(got) == {("FinMax", "ac.C40510"), ("FinMax", "ac.C405"), ("FinMax", "ac.C40S")}


def test_inherit_up_without_a_named_ancestor_takes_only_the_parent(adb):
    got = expand(_fact("up", ["ac.C40510"]), adb, stops={})
    assert set(got) == {("FinMax", "ac.C40510"), ("FinMax", "ac.C405")}


def test_inherit_both(adb):
    got = expand(_fact("both", ["un.SY.2CON"]), adb, stops={"FinMax": {"un.SY"}})
    assert set(got) == {("FinMax", "un.SY.2CON"), ("FinMax", "CC1"), ("FinMax", "CC2"), ("FinMax", "un.SY")}


def test_a_member_missing_from_the_atlas_stays_as_itself(adb):
    assert expand(_fact("down", ["un.SY.GONE"]), adb, stops={}) == {("FinMax", "un.SY.GONE"): "un.SY.GONE"}


def test_fact_inherit_defaults_to_self_and_is_strict():
    assert Fact(id="f", text="t", source="s", scope_metrics=["opex"]).inherit == "self"
    with pytest.raises(ValueError):
        Fact(id="f", text="t", source="s", scope_metrics=["opex"], inherit="sideways")
```

- [ ] **Step 2: Run them to see them fail**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py < /dev/null`
Expected: ERROR, `ModuleNotFoundError: No module named 'smartass.knowledge.expand'`

- [ ] **Step 3: Add the field to `Fact`**

In `engine/smartass/knowledge/semantic.py`, class `Fact`, after `scope_metrics: list[str] = []` add:

```python
    # self: only the named members; down: and every descendant; up: and the ancestors up to (and including) the
    # nearest one a metric or an entity names (only the parent when none does); both: down + up.
    inherit: Literal["self", "down", "up", "both"] = "self"
```

- [ ] **Step 4: Write `expand.py`**

Create `engine/smartass/knowledge/expand.py`:

```python
"""A fact's member scope expanded over the atlas tree (spec 2026-11-10 §4). Pure reads of atlas.db; the build
stores the result in fact_scope with `via` = the member the fact names."""
from __future__ import annotations

import sqlite3

from .semantic import Fact


def _rows(db: sqlite3.Connection, cube: str, member: str) -> list[tuple[str, str, str]]:
    """(name, dimension, parent) of every placement of the member (a shared member sits under several parents)."""
    return db.execute("SELECT name, dimension, coalesce(parent, '') FROM member WHERE connection=? AND name=? "
                      "COLLATE NOCASE", (cube, member)).fetchall()


def descendants(db: sqlite3.Connection, cube: str, member: str) -> set[str]:
    out: set[str] = set()
    todo = [(r[0], r[1]) for r in _rows(db, cube, member)]
    while todo:
        name, dim = todo.pop()
        for (kid,) in db.execute("SELECT DISTINCT name FROM member WHERE connection=? AND dimension=? AND parent=?",
                                 (cube, dim, name)):
            if kid not in out:
                out.add(kid)
                todo.append((kid, dim))
    return out


def ancestors(db: sqlite3.Connection, cube: str, member: str, stops: set[str]) -> set[str]:
    """Up every parent chain to the nearest ancestor in `stops` (included); a chain without one gives its parent."""
    out: set[str] = set()
    folded = {s.casefold() for s in stops}
    for _name, _dim, parent in _rows(db, cube, member):
        chain, p = [], parent
        while p:
            chain.append(p)
            if p.casefold() in folded:
                break
            row = _rows(db, cube, p)
            p = row[0][2] if row else ""
        else:
            chain = chain[:1]
        out.update(chain)
    return out


def expand(fact: Fact, db: sqlite3.Connection, stops: dict[str, set[str]]) -> dict[tuple[str, str], str]:
    """{(cube, member): via} for the fact's member scope; `via` is the member the fact itself names."""
    out: dict[tuple[str, str], str] = {}
    for cube, members in fact.scope_members.items():
        for m in members:
            out.setdefault((cube, m), m)
            reach: set[str] = set()
            if fact.inherit in ("down", "both"):
                reach |= descendants(db, cube, m)
            if fact.inherit in ("up", "both"):
                reach |= ancestors(db, cube, m, stops.get(cube, set()))
            for x in reach:
                out.setdefault((cube, x), m)
    return out
```

- [ ] **Step 5: Run the tests**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py < /dev/null`
Expected: `7 passed`

- [ ] **Step 6: Commit**

```bash
git add engine/smartass/knowledge/expand.py engine/smartass/knowledge/semantic.py engine/tests/v2/test_v2_knowledge_a.py
git commit -m "feat(knowledge): facts inherit down/up/both over the atlas tree (expand.py)"
```

---

### Task 3: Build stores the expanded scope; caveats say `via`; rebuild on a new atlas

**Files:**
- Modify: `engine/smartass/knowledge/build.py` (SCHEMA, SCHEMA_VERSION, `_write`, `build`)
- Modify: `engine/smartass/knowledge/resolver.py` (`_in_scope`, `caveats`, `ensure_built`)
- Modify: every `knowledge/semantic/facts/*.toml` (add `inherit`)
- Test: `engine/tests/v2/test_v2_knowledge_a.py`

- [ ] **Step 1: Write the failing tests** (append to the test file)

```python
from smartass.knowledge.build import build
from smartass.knowledge.resolver import KnowledgeResolver

COGS = """
name = "COGS"
aliases = ["cogs"]
description = "COGS."
sign = "cost"
status = "approved"
[finmax]
members = ["ac.C40S"]
"""
TRAP = """
text = "Handset COGS is booked on products, not offers."
severity = "warn"
source = "test"
status = "approved"
inherit = "up"
[scope_members]
FinMax = ["ac.C40510"]
"""
REORG = """
text = "Consumer Norway was abolished on 1 Jun 2026."
severity = "warn"
source = "test"
status = "approved"
inherit = "down"
[scope_members]
FinMax = ["un.SY.2CON"]
"""


def _write(d, files):
    for rel, text in files.items():
        (d / rel).parent.mkdir(parents=True, exist_ok=True)
        (d / rel).write_text(text, encoding="utf-8")
    return d


@pytest.fixture
def built(tmp_path, adb):
    sem = _write(tmp_path / "semantic", {"metrics/cogs.toml": COGS, "facts/trap.toml": TRAP,
                                          "facts/reorg.toml": REORG})
    out = tmp_path / "knowledge.db"
    rep = build(sem, tmp_path / "atlas.db", out)
    assert rep.ok, rep.to_text()
    k = sqlite3.connect(out)
    r = KnowledgeResolver(k, adb, 26)
    yield r
    k.close()


def test_a_trap_on_a_leaf_reaches_the_metric_total_with_via(built):
    (c,) = built.caveats({"FinMax": {"ac.C40S"}})
    assert c["id"] == "trap" and c["via"] == ["ac.C40510"]


def test_a_trap_does_not_ride_above_the_metric_total(built):
    assert built.caveats({"FinMax": {"ac.EBITDA"}}) == []


def test_a_reorg_reaches_the_units_below(built):
    (c,) = built.caveats({"FinMax": {"CC2"}})
    assert c["id"] == "reorg" and c["via"] == ["un.SY.2CON"]


def test_a_direct_hit_has_no_via(built):
    (c,) = built.caveats({"FinMax": {"un.SY.2CON"}})
    assert "via" not in c
```

- [ ] **Step 2: Run them to see them fail**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py -k "trap or reorg or direct" < /dev/null`
Expected: FAIL (no caveat on `ac.C40S`; `KeyError: 'via'`)

- [ ] **Step 3: Build: schema, expansion, atlas stamp**

In `engine/smartass/knowledge/build.py`:

1. Make the schema and version changes:
   - Replace `CREATE TABLE fact_scope (fact_id TEXT, cube TEXT, member TEXT, metric_id TEXT);` with `CREATE TABLE fact_scope (fact_id TEXT, cube TEXT, member TEXT, metric_id TEXT, via TEXT);`.
   - Set `SCHEMA_VERSION = "3a"`.
2. Add the import `from .expand import expand`.
3. In `_write`, replace the two `fact_scope` inserts with:

```python
    stops: dict[str, set[str]] = {}
    for m in metrics:
        for cube, ms in m.members().items():
            stops.setdefault(cube, set()).update(ms)
    for e in entities:
        for cube, member in e.members().items():
            stops.setdefault(cube, set()).add(member)
    for f in facts:
        db.execute("INSERT INTO fact VALUES (?,?)", (f.id, f.model_dump_json()))
        db.executemany("INSERT INTO fact_scope VALUES (?,?,?,NULL,?)",
                       [(f.id, cube, member, via) for (cube, member), via in expand(f, atlas, stops).items()])
        db.executemany("INSERT INTO fact_scope VALUES (?,NULL,NULL,?,NULL)", [(f.id, mid) for mid in f.scope_metrics])
```

   The signature becomes `_write(db, metrics, facts, atlas, meta, concepts=(), entities=())`. `Entity.members()` arrives in Task 5; until then `build` passes `entities=()`.

4. In `meta` inside `build()`, `atlas_built_at` already exists. Leave it as it is: `ensure_built` compares it.

- [ ] **Step 4: Resolver: `via` and the atlas trigger**

In `engine/smartass/knowledge/resolver.py`, replace `_in_scope` and the first three lines of `caveats`:

```python
    def _in_scope(self, touched: dict[str, set[str]], metric_ids: set[str]) -> dict[str, set[str]]:
        """{fact id: the members it came from when it reached a touched member by inheritance}."""
        ids: dict[str, set[str]] = {}
        for cube, members in touched.items():
            for m in members:
                for fid, member, via in self.kdb.execute(
                        "SELECT fact_id, member, via FROM fact_scope WHERE cube=? AND member=? COLLATE NOCASE",
                        (cube, m)):
                    got = ids.setdefault(fid, set())
                    if via and via.casefold() != member.casefold():
                        got.add(via)
        if metric_ids:
            marks = ",".join("?" * len(metric_ids))
            for (fid,) in self.kdb.execute(f"SELECT fact_id FROM fact_scope WHERE metric_id IN ({marks})",
                                           sorted(metric_ids)):
                ids.setdefault(fid, set())
        return ids
```

```python
        scope = self._in_scope(touched, set(metric_ids))
        hits = [self._facts[i] for i in scope if i in self._facts]
        hits = [f for f in hits if self._valid(f, months)]
        out = [{"id": f.id, "text": f.text, "severity": f.severity, **({"via": sorted(scope[f.id])} if scope[f.id]
                                                                        else {})}
               for f in hits if f.status == "approved"]
```

Also in `ensure_built` add the atlas trigger. Facts now depend on the tree, so a new atlas must rebuild knowledge.db. Replace the early-return condition with:

```python
        meta = summary(path) if not missing else {}
        fresh = (str(meta.get("semantic_hash", "")) == semantic_hash(semantic)
                 and str(meta.get("atlas_built_at", "")) == _atlas_stamp(Path(config.ATLAS_DB)))
        if not missing and fresh:
            return path
```

and add to `resolver.py`:

```python
def _atlas_stamp(path: Path) -> str:
    """The atlas build time as build() writes it into meta ('' without an atlas)."""
    import time
    if not path.exists():
        return ""
    db = sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True)
    try:
        t = db.execute("SELECT max(built_at) FROM build").fetchone()[0]
    finally:
        db.close()
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(t)) if t else ""
```

Change `REBUILT_NOTE` to `"knowledge.db rebuilt: semantic files or the atlas changed"`.

- [ ] **Step 5: Run the new tests, then the knowledge tests**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py tests/v2/test_v2_knowledge_build.py tests/v2/test_v2_knowledge_resolver.py tests/v2/test_v2_knowledge_rebuild.py tests/v2/test_v2_knowledge_wiring.py < /dev/null`
Expected: all pass.

`test_v2_knowledge_rebuild.py` may pin `REBUILT_NOTE` or the rebuild condition. If it fails, update the pinned text to the new constant. A test that writes an atlas after knowledge.db and expects no rebuild must now expect a rebuild. Change it and say why in the test name.

- [ ] **Step 6: Set `inherit` on every existing fact**

Write `scripts/set_inherit.py` (scratch, not committed) that appends one line after the `status` line of each fact file:
- `reorg_*`, `abolished_*`, `org_*`, `finplan_banorway`: `inherit = "down"`;
- `account_*`: `inherit = "up"`;
- `finplan_kpis_periodic`, `finplan_no_actual_volumes`, `finplan_target_stops_at_ebitda`: leave `self`.

```python
import pathlib
import re
import sys

d = pathlib.Path(sys.argv[1])
for p in sorted(d.glob("*.toml")):
    t = p.read_text(encoding="utf-8")
    if "inherit =" in t:
        continue
    kind = ("down" if re.match(r"(reorg|abolished|org)_|finplan_banorway$", p.stem)
            else "up" if p.stem.startswith("account_") else None)
    if kind:
        p.write_text(re.sub(r'^(status = .*)$', rf'\1\ninherit = "{kind}"', t, count=1, flags=re.M), encoding="utf-8")
        print(p.stem, kind)
```

Run: `uv run python -B ../scripts/set_inherit.py ../knowledge/semantic/facts < /dev/null` (from `engine/`; write the script in the scratchpad, not the repo)
Expected: one line per fact. `finplan_kpis_periodic`, `finplan_no_actual_volumes` and `finplan_target_stops_at_ebitda` are not listed.

- [ ] **Step 7: Build the real knowledge and check the two cases from the spec**

Run: `uv run smartass knowledge build < /dev/null | head -3`
Expected: `knowledge.db built: ...` with 0 errors.

Then write `scratch_check_caveats.py` in the scratchpad:

```python
import contextlib
from smartass.knowledge.resolver import open_resolver
with contextlib.closing(open_resolver(26)) as r:
    print("C40S", [c["id"] for c in r.caveats({"FinMax": {"ac.C40S"}})])
    print("2TEC", [c["id"] for c in r.caveats({"FinMax": {"un.SY.2TEC"}})])
    print("EBITDA", [c["id"] for c in r.caveats({"FinMax": {"ac.EBITDA"}})])
```

Run: `uv run python -B <scratchpad>/scratch_check_caveats.py < /dev/null`
Expected:
- `C40S` lists `account_c40510`.
- `2TEC` lists at least one `reorg_*_tec` fact. Check which facts are scoped to units under 2TEC; `reorg_ig_tec` moved units into TEC.
- `EBITDA` lists no `account_*` fact, except `account_c73`/`account_c75` if they name EBITDA-level members.

If `2TEC` lists none: the reorg facts name the moved cost centres, and `down` from them never reaches the division. Then change those reorg facts to `inherit = "both"`. `up` stops at the entity member `un.SY.2TEC` once entities exist (Task 5). Rerun after Task 5.

- [ ] **Step 8: Full tests, commit**

Run: `uv run python -B -m pytest -q tests < /dev/null | tail -1`
Expected: all pass.

```bash
git checkout engine/uv.lock
git add -A engine/smartass/knowledge engine/tests knowledge/semantic/facts
git commit -m "feat(knowledge): fact scopes expanded at build (via), caveats say where they came from; a new atlas rebuilds knowledge.db"
```

---

### Task 4: New semantic kinds `Rule`, `Entity`, `Kpi`, `Vocab`, `MemberNote`

**Files:**
- Modify: `engine/smartass/knowledge/semantic.py` (models + `load_all`)
- Test: `engine/tests/v2/test_v2_knowledge_a.py`

- [ ] **Step 1: Write the failing tests** (append)

```python
from smartass.knowledge.semantic import Entity, Kpi, MemberNote, Rule, SemanticError, Vocab, load_all


def test_rule_kinds_validate_their_data():
    Rule(id="scenarios", title="Scenarios", text="t", kind="scenarios",
         data={"FinMax": {"value": ["Actual", ""], "ly": ["Actual", ""]},
               "FinPlan": {"value": ["Actual", "Working"], "ly": ["Actual", "Working"],
                           "budget": ["Target", "Final"], "forecast": ["latest_forecast", "Final"]}})
    with pytest.raises(ValueError, match="FinPlan"):
        Rule(id="scenarios", title="S", text="t", kind="scenarios",
             data={"FinMax": {"value": ["Actual", ""], "ly": ["Actual", ""]}})
    Rule(id="materiality", title="M", text="t", kind="materiality", data={"abs_nok": 5e6, "pct": 0.1})
    with pytest.raises(ValueError, match="pct"):
        Rule(id="materiality", title="M", text="t", kind="materiality", data={"abs_nok": 5e6})
    Rule(id="signs", title="Signs", text="Costs are negative in FinMax.")          # kind text needs no data


def test_entity_members_and_finplan_total():
    e = Entity(id="business", label="Business", aliases=["bedrift"], finmax="un.SY.2BUS",
               finplan="2BUS Business", finplan_total="1BUS Business", source="s", status="approved")
    assert e.members() == {"FinMax": "un.SY.2BUS", "FinLite": "un.SY.2BUS", "FinPlan": "2BUS Business"}
    with pytest.raises(ValueError):
        Entity(id="x", label="X", source="s")                                       # no member in any cube


def test_kpi_needs_two_metrics_and_a_known_unit():
    Kpi(id="ebitda_margin", name="EBITDA margin", cube="FinMax", numerator="ebitda", denominator="revenue",
        unit="pct", source="s")
    with pytest.raises(ValueError):
        Kpi(id="k", name="K", cube="FinMax", numerator="ebitda", denominator="ebitda", unit="pct", source="s")


def test_vocab_kinds():
    Vocab(id="costs", kind="vague", words=["costs", "kostnader"], options=["opex", "opex + COGS"], source="s")
    with pytest.raises(ValueError, match="options"):
        Vocab(id="costs", kind="vague", words=["costs"], source="s")
    Vocab(id="per_customer", kind="grain", words=["per customer"], feasible=[], proxy="per business area",
          source="s")
    Vocab(id="currency", kind="currency", words=["eur"], codes={"eur": "EUR"}, reporting=["NOK"], source="s")


def test_load_all_reads_every_kind_and_ids_are_unique(tmp_path):
    d = _write(tmp_path / "semantic", {
        "metrics/cogs.toml": COGS,
        "rules/signs.toml": 'title = "Signs"\ntext = "Costs negative."\n',
        "entities/cogs.toml": 'label = "X"\nfinmax = "un.SY"\nsource = "s"\n'})
    with pytest.raises(SemanticError, match="duplicate id 'cogs'"):
        load_all(d)


def test_member_notes_load_from_a_list_file(tmp_path):
    d = _write(tmp_path / "semantic", {"notes/finmax_accounts.toml":
               '[[note]]\ncube = "FinMax"\nmember = "ac.C85"\ntext = "Profit before tax."\n'})
    (n,) = load_all(d).notes
    assert n == MemberNote(cube="FinMax", member="ac.C85", text="Profit before tax.")
```

- [ ] **Step 2: Run them to see them fail**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py -k "rule or entity or kpi or vocab or load_all or notes" < /dev/null`
Expected: ImportError on `Entity`.

- [ ] **Step 3: Add the models**

Append to `engine/smartass/knowledge/semantic.py` (before `to_toml` helpers is fine; keep `CUBES` import as it is):

```python
# ---- knowledge layer phase A (spec 2026-11-10 §2)

class _Item(BaseModel):
    model_config = ConfigDict(extra="forbid")
    id: str
    source: str = ""
    status: Literal["approved", "draft"] = "approved"
    approved: dict | None = None               # {"by": "tom" | "auto" | <colleague>, "on": "YYYY-MM-DD"}
    verified: dict | None = None

    @field_validator("id")
    @classmethod
    def _check_id(cls, v: str) -> str:
        return _id(v)


SCENARIO_SERIES = {"FinMax": ("value", "ly"), "FinLite": ("value", "ly"),
                   "FinPlan": ("value", "ly", "budget", "forecast")}


class Rule(_Item):
    """An always-true statement about the cubes. kind text: just the text; the other kinds carry data the engine
    reads (one rule per such kind)."""
    title: str
    text: str
    kind: Literal["text", "scenarios", "materiality", "pnl", "group_companies"] = "text"
    data: dict = {}

    @model_validator(mode="after")
    def _check(self):
        d = self.data
        if self.kind == "scenarios":
            for cube in ("FinMax", "FinPlan"):
                need = SCENARIO_SERIES[cube]
                got = d.get(cube, {})
                if set(got) != set(need) or any(len(v) != 2 for v in got.values()):
                    raise ValueError(f"rule {self.id}: scenarios.{cube} needs exactly {need} as [scenario, version]")
        elif self.kind == "materiality":
            if not (isinstance(d.get("abs_nok"), (int, float)) and isinstance(d.get("pct"), (int, float))):
                raise ValueError(f"rule {self.id}: materiality needs numbers abs_nok and pct")
        elif self.kind == "pnl":
            for cube in ("FinMax", "FinPlan"):
                if not isinstance(d.get(cube), list) or not d[cube]:
                    raise ValueError(f"rule {self.id}: pnl.{cube} needs the metric ids in report order")
        elif self.kind == "group_companies":
            if not isinstance(d.get("patterns"), list):
                raise ValueError(f"rule {self.id}: group_companies needs patterns")
        elif d:
            raise ValueError(f"rule {self.id}: a text rule has no data")
        return self


class Entity(_Item):
    label: str
    aliases: list[str] = []
    generic_aliases: list[str] = []            # count only when no other entity is named ("total")
    finmax: str | None = None                  # FinLite shares the FinMax member
    finplan: str | None = None
    finplan_total: str | None = None           # the FinPlan level-1 parent when `finplan` is not under BANORWAY
    abolished: str = ""
    successors: list[str] = []
    note: str = ""

    @model_validator(mode="after")
    def _check(self):
        if not (self.finmax or self.finplan):
            raise ValueError(f"entity {self.id!r} needs a FinMax or FinPlan member")
        return self

    def members(self) -> dict[str, str]:
        out = {}
        if self.finmax:
            out["FinMax"] = out["FinLite"] = self.finmax
        if self.finplan:
            out["FinPlan"] = self.finplan
        return out


class Kpi(_Item):
    """A ratio of two metrics of one cube, computed in the workbook."""
    name: str
    aliases: list[str] = []
    description: str = ""
    cube: Literal["FinMax", "FinLite", "FinPlan"]
    numerator: str
    denominator: str
    unit: Literal["pct", "nok", "nok_per_fte"]
    basis: Literal["period", "average_stock"] = "period"
    annualise: bool = False
    sign: Literal[1, -1] = 1
    better: Literal["higher", "lower", "none"] = "higher"
    headline: bool = False

    @model_validator(mode="after")
    def _check(self):
        if self.numerator == self.denominator:
            raise ValueError(f"kpi {self.id!r}: numerator and denominator are the same metric")
        return self


class Vocab(_Item):
    """Words the engine must interpret: vague (ask with options), grain (what is answerable per X), currency."""
    kind: Literal["vague", "grain", "currency"]
    words: list[str]
    options: list[str] = []
    feasible: list[str] = []                   # grain: metric/kpi ids answerable at it; ["*"] = all
    proxy: str = ""
    note: str = ""
    codes: dict[str, str] = {}                 # currency: word -> ISO code
    reporting: list[str] = []                  # currency: the codes the cubes report

    @model_validator(mode="after")
    def _check(self):
        if self.kind == "vague" and len(self.options) < 2:
            raise ValueError(f"vocab {self.id!r}: a vague word needs at least two options")
        if self.kind == "currency" and not (self.codes and self.reporting):
            raise ValueError(f"vocab {self.id!r}: currency needs codes and reporting")
        return self


class MemberNote(BaseModel):
    """What one member means and how it is booked (folded from the 15.10 business maps)."""
    model_config = ConfigDict(extra="forbid")
    cube: Literal["FinMax", "FinLite", "FinPlan"]
    member: str
    text: str


@dataclass
class Knowledge:
    metrics: list
    facts: list
    concepts: list
    rules: list
    entities: list
    kpis: list
    vocab: list
    notes: list


def _load_notes(directory: Path) -> list[MemberNote]:
    out = []
    for p in sorted((directory / "notes").glob("*.toml")):
        try:
            data = tomllib.loads(p.read_text(encoding="utf-8"))
            out += [MemberNote(**n) for n in data.get("note", [])]
        except (tomllib.TOMLDecodeError, ValidationError, TypeError) as e:
            raise SemanticError(f"{p.name}: {e}") from e
    return out


def load_all(directory: Path) -> Knowledge:
    """Every kind under `directory`, validated; ids are unique across all kinds with an id."""
    directory = Path(directory)
    metrics, facts = load_semantic(directory)
    concepts = load_concepts(directory)
    kinds = {"rules": Rule, "entities": Entity, "kpis": Kpi, "vocab": Vocab}
    loaded = {k: [_load(p, cls) for p in sorted((directory / k).glob("*.toml"))] for k, cls in kinds.items()}
    seen: dict[str, str] = {}
    for kind, items in (("metrics", metrics), ("facts", facts), ("concepts", concepts), *loaded.items()):
        for it in items:
            if it.id in seen:
                raise SemanticError(f"duplicate id {it.id!r}: {seen[it.id]}/{it.id}.toml and {kind}/{it.id}.toml")
            seen[it.id] = kind
    return Knowledge(metrics, facts, concepts, loaded["rules"], loaded["entities"], loaded["kpis"], loaded["vocab"],
                     _load_notes(directory))
```

Add `from dataclasses import dataclass` to the imports if it is missing. `Path`, `tomllib`, `ValidationError`, `ConfigDict`, `field_validator` and `model_validator` are already imported; check the top of the file.

- [ ] **Step 4: Run the tests**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py < /dev/null`
Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/knowledge/semantic.py engine/tests/v2/test_v2_knowledge_a.py
git commit -m "feat(knowledge): semantic kinds Rule, Entity, Kpi, Vocab, MemberNote and load_all"
```

---

### Task 5: Build the new kinds, with the contradiction checks

**Files:**
- Modify: `engine/smartass/knowledge/build.py`
- Test: `engine/tests/v2/test_v2_knowledge_a.py`

- [ ] **Step 1: Write the failing tests** (append)

```python
PLAN_TREE = [("FinPlan", "Entity", "BANORWAY", "BANORWAY", "Entity", 1, 1),
             ("FinPlan", "Entity", "1BUS Business", "1BUS Business", "BANORWAY", 2, 1),
             ("FinPlan", "Entity", "2BUS Business", "2BUS Business", "1BUS Business", 3, 1),
             ("FinPlan", "Entity", "2INT_B Ambition Business", "2INT_B Ambition Business", "1BUS Business", 3, 2),
             ("FinMax", "All Entities", "un.SY.2BUS", "2BUS Business", "un.SY", 2, 3),
             ("FinMax", "All Accounts", "ac.C30", "C30 Revenue", "ac.C48", 3, 2)]
BUS = 'label = "Business"\naliases = ["bedrift"]\nfinmax = "un.SY.2BUS"\nfinplan = "2BUS Business"\nsource = "s"\n'
REV = ('name = "Revenue"\naliases = ["revenue"]\ndescription = "R."\nsign = "income"\nstatus = "approved"\n'
       '[finmax]\nmembers = ["ac.C30"]\n')


def _build(tmp_path, adb, files):
    adb.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)", PLAN_TREE)
    adb.commit()
    sem = _write(tmp_path / "semantic", {"metrics/cogs.toml": COGS, "metrics/revenue.toml": REV, **files})
    return build(sem, tmp_path / "atlas.db", tmp_path / "knowledge.db")


def test_a_finplan_entity_below_level_one_needs_its_total(tmp_path, adb):
    rep = _build(tmp_path, adb, {"entities/business.toml": BUS})
    assert any("2BUS Business" in e and "1BUS Business" in e and "finplan_total" in e for e in rep.errors)
    rep = _build(tmp_path, adb, {"entities/business.toml": BUS + 'finplan_total = "1BUS Business"\n'})
    assert rep.ok, rep.to_text()


def test_a_kind_with_data_is_defined_once(tmp_path, adb):
    mat = 'title = "M"\ntext = "t"\nkind = "materiality"\ndata = { abs_nok = 5e6, pct = 0.1 }\n'
    rep = _build(tmp_path, adb, {"rules/materiality.toml": mat, "rules/materiality2.toml": mat})
    assert any("materiality" in e and "once" in e for e in rep.errors)


def test_a_kpi_names_metrics_of_its_cube(tmp_path, adb):
    kpi = ('name = "GM %"\ncube = "FinMax"\nnumerator = "cogs"\ndenominator = "nope"\nunit = "pct"\n'
           'source = "s"\n')
    rep = _build(tmp_path, adb, {"kpis/gm_pct.toml": kpi})
    assert any("nope" in e for e in rep.errors)


def test_an_alias_belongs_to_one_item_across_kinds(tmp_path, adb):
    ent = 'label = "Rev"\naliases = ["revenue"]\nfinmax = "un.SY.2BUS"\nsource = "s"\n'
    rep = _build(tmp_path, adb, {"entities/rev_unit.toml": ent})
    assert any("'revenue'" in e for e in rep.errors)
```

- [ ] **Step 2: Run them to see them fail**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py -k "level_one or once or kpi_names or across_kinds" < /dev/null`
Expected: FAIL (the build is ok, and no errors are found).

- [ ] **Step 3: Implement**

In `build.py`:

1. Make the schema and version changes:
   - Add to `SCHEMA`:

```sql
CREATE TABLE rule (id TEXT PRIMARY KEY, kind TEXT, json TEXT);
CREATE TABLE entity (id TEXT PRIMARY KEY, json TEXT);
CREATE TABLE kpi (id TEXT PRIMARY KEY, json TEXT);
CREATE TABLE vocab (id TEXT PRIMARY KEY, kind TEXT, json TEXT);
CREATE TABLE term (text TEXT, kind TEXT, item_id TEXT);
CREATE INDEX ix_term_text ON term(text COLLATE NOCASE);
CREATE TABLE member_note (cube TEXT, member TEXT, text TEXT);
CREATE INDEX ix_member_note ON member_note(cube, member COLLATE NOCASE);
```

   - Set `SCHEMA_VERSION = "3a2"`.

2. In `build()`, replace the two `load_*` calls with `k = load_all(Path(semantic_dir))`. Then use `k.metrics`, `k.facts` and `k.concepts` where `metrics`, `facts` and `concepts` were used. Pass `k` into `check` and `_write`. Add `rules`, `entities`, `kpis`, `vocab` and `notes` counts to `meta`.

3. Add the new checks (call them at the end of `check`, with `k` as an extra argument):

```python
DATA_KINDS = ("scenarios", "materiality", "pnl", "group_companies")


def check_layer(k, atlas: sqlite3.Connection, report: BuildReport) -> None:
    """Spec 2026-11-10 §6: contradictions stop the build."""
    for kind in DATA_KINDS:
        ids = [r.id for r in k.rules if r.kind == kind]
        if len(ids) > 1:
            report.errors.append(f"rule kind {kind!r} must be defined once, found {', '.join(ids)}")
    metric_ids = {m.id: m for m in k.metrics}
    entity_ids = {e.id for e in k.entities}
    for e in k.entities:
        for cube, member in e.members().items():
            if cube != "FinLite" and _missing(atlas, cube, [member]):
                report.errors.append(f"entity {e.id}: {cube} member {member!r} is not in the atlas")
        for s in e.successors:
            if s not in entity_ids:
                report.errors.append(f"entity {e.id}: successor {s!r} is not an entity")
        if e.finplan and not _missing(atlas, "FinPlan", [e.finplan]):
            parent = atlas.execute("SELECT parent FROM member WHERE connection='FinPlan' AND dimension='Entity' "
                                   "AND name=? ORDER BY depth LIMIT 1", (e.finplan,)).fetchone()
            under = parent[0] if parent else ""
            if under not in ("BANORWAY", "Entity", "") and e.finplan_total != under:
                report.errors.append(f"entity {e.id}: FinPlan {e.finplan!r} sits under {under!r}, not BANORWAY; "
                                     f"set finplan_total = {under!r} (the division total) or fix finplan")
    for kp in k.kpis:
        for side in (kp.numerator, kp.denominator):
            m = metric_ids.get(side)
            if m is None or kp.cube not in m.members():
                report.errors.append(f"kpi {kp.id}: {side!r} is not a metric with {kp.cube} members")
    for v in k.vocab:
        for f in v.feasible:
            if f != "*" and f not in metric_ids and f not in {x.id for x in k.kpis}:
                report.errors.append(f"vocab {v.id}: feasible {f!r} is not a metric or kpi")
    owner: dict[str, str] = {}
    for m in k.metrics:
        for t in {m.name, *m.aliases}:
            owner.setdefault(t.casefold(), f"metric {m.id}")
    for kind, items in (("entity", k.entities), ("kpi", k.kpis), ("vocab", k.vocab)):
        for it in items:
            terms = {*getattr(it, "aliases", []), *getattr(it, "generic_aliases", []), *getattr(it, "words", [])}
            if kind == "kpi":
                terms.add(it.name)
            for t in terms:
                was = owner.setdefault(t.casefold(), f"{kind} {it.id}")
                if was != f"{kind} {it.id}":
                    report.errors.append(f"alias {t!r} belongs to {was} and {kind} {it.id}")
    gone = [f"{n.cube} {n.member}" for n in k.notes if _missing(atlas, n.cube, [n.member])]
    if gone:
        report.warnings.append(f"{len(gone)} member note(s) name members not in the atlas: {', '.join(gone[:10])}"
                               f"{' ...' if len(gone) > 10 else ''}")
```

   The note check queries the atlas once per note, about 2 400 point lookups. That is fine for a build.

4. Make `_write` store the new kinds:

```python
    for r in k.rules:
        db.execute("INSERT INTO rule VALUES (?,?,?)", (r.id, r.kind, r.model_dump_json()))
    for e in k.entities:
        db.execute("INSERT INTO entity VALUES (?,?)", (e.id, e.model_dump_json()))
        db.executemany("INSERT INTO term VALUES (?,?,?)",
                       [(t, "entity", e.id) for t in dict.fromkeys([e.id, e.label, *e.aliases, *e.generic_aliases])])
    for kp in k.kpis:
        db.execute("INSERT INTO kpi VALUES (?,?)", (kp.id, kp.model_dump_json()))
        db.executemany("INSERT INTO term VALUES (?,?,?)",
                       [(t, "kpi", kp.id) for t in dict.fromkeys([kp.id.replace("_", " "), kp.name, *kp.aliases])])
    for v in k.vocab:
        db.execute("INSERT INTO vocab VALUES (?,?,?)", (v.id, v.kind, v.model_dump_json()))
        db.executemany("INSERT INTO term VALUES (?,?,?)", [(t, "vocab", v.id) for t in dict.fromkeys(v.words)])
    db.executemany("INSERT INTO member_note VALUES (?,?,?)", [(n.cube, n.member, n.text) for n in k.notes])
```

   The `entities` argument of the fact expansion (Task 3) is now `k.entities`.

- [ ] **Step 4: Run all knowledge tests**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py tests/v2/test_v2_knowledge_build.py tests/v2/test_v2_knowledge_refresh.py < /dev/null`
Expected: all pass. `refresh.py` calls `build()` unchanged, so it must still work.

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/knowledge/build.py engine/tests/v2/test_v2_knowledge_a.py
git commit -m "feat(knowledge): build rules, entities, KPIs, vocab and member notes; contradictions stop the build"
```

---

### Task 6: Fold catalog, capabilities and the maps into TOML

**Files:**
- Create: `scripts/fold_knowledge.py`
- Create (generated): `knowledge/semantic/rules/*.toml`, `entities/*.toml`, `kpis/*.toml`, `vocab/*.toml`, `notes/{finmax_accounts,finmax_entities,finplan_accounts}.toml`, `metrics/<bundle>.toml`
- Modify: `knowledge/semantic/metrics/fte.toml` (materiality)

- [ ] **Step 1: Write the script**

Create `scripts/fold_knowledge.py`:

```python
"""One-off (spec 2026-11-10 §2): fold knowledge/catalog.toml, capabilities.toml and the 15.10 business maps into
knowledge/semantic/{rules,entities,kpis,vocab,notes} and bundle metrics. Run once from engine/:
    uv run python -B ../scripts/fold_knowledge.py < /dev/null
Refuses to overwrite an existing file. Deleted after the fold (Task 11); git history keeps it."""
import json
import pathlib
import sys
import tomllib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "engine"))
from smartass.knowledge.semantic import Entity, Kpi, Metric, Rule, Vocab, to_toml  # noqa: E402

K = pathlib.Path(__file__).resolve().parents[1] / "knowledge"
S = K / "semantic"
TOM = {"by": "tom", "on": "2026-10-15", "via": "catalog.toml"}
LINE_METRIC = {"rev": "revenue", "cogs": "cogs", "gm": "gross_margin", "gp": "gross_margin", "pers": "personnel",
               "oth": "other_opex", "opex": "opex", "boi": "ebitda_boi", "c73": "other_items",
               "ebitda": "ebitda", "da": "da", "fte": "fte"}
cat = tomllib.loads((K / "catalog.toml").read_text(encoding="utf-8"))
cap = tomllib.loads((K / "capabilities.toml").read_text(encoding="utf-8"))
written = []


def put(folder: str, item_id: str, text: str) -> None:
    p = S / folder / f"{item_id}.toml"
    if p.exists():
        raise SystemExit(f"refusing to overwrite {p}")
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding="utf-8")
    written.append(p.relative_to(S).as_posix())


def toml_data(item) -> str:
    """to_toml covers Metric/Fact/Concept; the new kinds are written by hand: scalars and lists first, then tables."""
    d = item.model_dump(exclude_defaults=True, exclude={"id"})
    out, tables = [], {}
    for key, v in d.items():
        if isinstance(v, dict):
            tables[key] = v
        else:
            out.append(f"{key} = {json.dumps(v, ensure_ascii=False)}")
    for key, v in tables.items():
        out.append(f"\n[{key}]")
        for k2, v2 in v.items():
            out.append(f"{json.dumps(k2, ensure_ascii=False) if ' ' in k2 or '&' in k2 else k2} = "
                       f"{json.dumps(v2, ensure_ascii=False)}")
    return "\n".join(out) + "\n"


# rules
fm = cat["pnl"]["FinMax"]
fp = cat["pnl"]["FinPlan"]
put("rules", "pnl", toml_data(Rule(
    id="pnl", title="P&L chain", kind="pnl", source="catalog.toml [pnl] + [checks]", approved=TOM,
    text="FinMax: C30 + C40S = C48; C50S + C60 = C70; C48 + C70 = EBITDA_BOI; + C73 = EBITDA; D&A C75 below. "
         "FinPlan: Revenue + COGS = Gross_Profit; + OPEX = EBITDA before Other Items; + Total Other Items = EBITDA "
         "(FinPlan stops at EBITDA). Join actual vs budget on the metric, never on the line key.",
    data={"FinMax": [LINE_METRIC[x["key"]] for x in fm], "FinPlan": [LINE_METRIC[x["key"]] for x in fp],
          "checks": {c: [{"label": x["label"], "expr": x["expr"]} for x in cat["checks"][c]] for c in cat["checks"]}})))
sc = cat["scenarios"]
put("rules", "scenarios", toml_data(Rule(
    id="scenarios", title="Scenarios per cube", kind="scenarios", source="catalog.toml [scenarios] + series.py",
    approved=TOM,
    text="Actuals: FinMax/FinLite scenario Actual. Plan lives only in FinPlan: actual = Actual/Working, budget = "
         "Target/Final, forecast = the latest Forecast Q<q><yy>/Final.",
    data={"FinMax": {"value": ["Actual", ""], "ly": ["Actual", ""]},
          "FinLite": {"value": ["Actual", ""], "ly": ["Actual", ""]},
          "FinPlan": {"value": [sc["actual_finplan"]["scenario"], sc["actual_finplan"]["version"]],
                      "ly": [sc["actual_finplan"]["scenario"], sc["actual_finplan"]["version"]],
                      "budget": [sc["budget"]["scenario"], sc["budget"]["version"]],
                      "forecast": ["latest_forecast", sc["forecast"]["version"]]}})))
mt = cat["materiality"]
put("rules", "materiality", toml_data(Rule(
    id="materiality", title="Materiality", kind="materiality", source="catalog.toml [materiality]", approved=TOM,
    text=f"A variance is material at {mt['pct']:.0%} of |comparison| (from NOK {mt['pct_floor_nok']:,}) or at NOK "
         f"{mt['abs_nok']:,}; revenue, personnel and EBITDA are always commented.".replace(",", " "),
    data={"abs_nok": float(mt["abs_nok"]), "pct": mt["pct"], "pct_floor_nok": float(mt["pct_floor_nok"]),
          "rule": mt["rule"], "always": [LINE_METRIC[x] for x in mt["always"]]})))
gc = cat["group_companies"]
put("rules", "group_companies", toml_data(Rule(
    id="group_companies", title="Group companies", kind="group_companies", source="catalog.toml [group_companies]",
    approved=TOM, text="Group companies get their own block in a supplier ranking, never in its top 10.",
    data={"column": gc["column"], "patterns": gc["patterns"]})))
put("rules", "signs", toml_data(Rule(
    id="signs", title="Signs and variance", source="catalog.toml header", approved=TOM,
    text="FinMax signs: revenue positive, costs negative, so a higher value is better on every P&L line. Variance = "
         "current - comparison (positive = better); variance % is empty on a tiny base.")))
put("rules", "cube_choice", toml_data(Rule(
    id="cube_choice", title="Which cube", source="skill smartass §6", approved=TOM,
    text="Actuals: FinMax (FTE ac.STA20Y). FinLite is lighter (no FTE) and the drill source for ledger lines. Plan "
         "(budget, forecast) only in FinPlan. History is restated to today's organisation.")))

# entities
for e in cat["entity"]:
    put("entities", e["key"], toml_data(Entity(
        id=e["key"], label=e["label"], aliases=e.get("aliases", []), generic_aliases=e.get("generic_aliases", []),
        finmax=e.get("finmax"), finplan=e.get("finplan"), abolished=e.get("abolished", ""),
        successors=e.get("successors", []), note=e.get("note", ""), source="catalog.toml [[entity]]",
        approved=TOM)))

# kpis (synonyms from capabilities [[measure]] kpi = ...)
syn = {m["kpi"]: m["synonyms"] for m in cap["measure"] if "kpi" in m}
note = {m["kpi"]: m.get("note", "") for m in cap["measure"] if "kpi" in m}
for kp in cat["kpi"]:
    put("kpis", kp["key"], toml_data(Kpi(
        id=kp["key"], name=kp["label"], aliases=[s for s in syn.get(kp["key"], []) if s.casefold() !=
                                                 kp["label"].casefold()],
        description=note.get(kp["key"], ""), cube=kp["connection"], numerator=LINE_METRIC[kp["numerator"]],
        denominator=LINE_METRIC[kp["denominator"]], unit=kp["unit"], basis=kp.get("basis", "period"),
        annualise=kp.get("annualise", False), sign=kp.get("sign", 1), better=kp.get("better", "higher"),
        headline=kp.get("headline", False), source="catalog.toml [[kpi]]", approved=TOM)))

# bundles -> metrics; vague words; grains; currency
for m in cap["measure"]:
    if "accounts" in m:
        put("metrics", m["key"], to_toml(Metric(
            id=m["key"], name=m["label"], aliases=m["synonyms"], description=m.get("note", "") or m["label"],
            sign="cost", status="approved", approved=TOM | {"via": "capabilities.toml"},
            finmax={"members": m["accounts"]}, finlite={"members": m["accounts"]})))
    elif m.get("vague"):
        put("vocab", m["key"], toml_data(Vocab(
            id=m["key"], kind="vague", words=m["synonyms"], options=m["options"],
            source="capabilities.toml [[measure]]", approved=TOM | {"via": "capabilities.toml"})))
for g in cap["grain"]:
    put("vocab", f"grain_{g['key']}", toml_data(Vocab(
        id=f"grain_{g['key']}", kind="grain", words=g["synonyms"],
        feasible=[LINE_METRIC.get(x, x) if x != "*" else x for x in g["feasible"]], proxy=g.get("proxy", ""),
        note=g.get("note", ""), source="capabilities.toml [[grain]]", approved=TOM | {"via": "capabilities.toml"})))
cur = cap["currency"]
put("vocab", "currency", toml_data(Vocab(
    id="currency", kind="currency", words=sorted(cur["words"]), codes=cur["words"], reporting=cur["reporting"],
    proxy=cur["proxy"], source="capabilities.toml [currency]", approved=TOM | {"via": "capabilities.toml"})))

# member notes from the maps (meaning + booking pattern / pitfalls; Norwegian as approved)
def notes(file, rows_key, cube, fmt):
    d = json.loads((K / file).read_text(encoding="utf-8"))
    lines = []
    for r in d[rows_key]:
        text = fmt(r).strip()
        if text:
            lines.append("[[note]]\n" + f"cube = {json.dumps(cube)}\nmember = {json.dumps(r['member'], ensure_ascii=False)}\n"
                         f"text = {json.dumps(text, ensure_ascii=False)}\n")
    return "\n".join(lines)


def join(*parts):
    return " ".join(p.rstrip(".") + "." for p in parts if p)


put("notes", "finmax_accounts", notes("accounts_map.json", "accounts", "FinMax",
                                      lambda r: join(r.get("meaning"), r.get("booking_pattern"), r.get("pitfalls"))))
put("notes", "finmax_entities", notes("org_map.json", "entities", "FinMax",
                                      lambda r: join(r.get("meaning"), r.get("kind") and f"Type: {r['kind']}",
                                                     r.get("moved") and f"Moved: {r['moved']}", r.get("pitfalls"))))
put("notes", "finplan_accounts", notes("finplan_map.json", "accounts", "FinPlan",
                                       lambda r: join(r.get("meaning"),
                                                      r.get("finmax_equivalent") and
                                                      f"FinMax equivalent: {r['finmax_equivalent']}",
                                                      r.get("pitfalls"))))
print("\n".join(written))
print(len(written), "files")
```

- [ ] **Step 2: Run it**

Run (from `engine/`): `uv run python -B ../scripts/fold_knowledge.py < /dev/null | tail -5`
Expected: one line per file, then `N files`. N is roughly 6 rules + 11 entities + 4 KPIs + 12 bundles + 2 vague + 7 grains + 1 currency + 3 notes = about 46.

- [ ] **Step 3: Fix the FTE materiality**

In `knowledge/semantic/metrics/fte.toml`, replace `materiality = [5000000.0, 0.1]` with `materiality = [10.0, 0.05]`. That means 10 FTE or 5 %; Tom confirms it in the report of Task 12.

- [ ] **Step 4: Build and resolve every error**

Run: `uv run smartass knowledge build < /dev/null`

Expected errors and how to fix each one:
- **`entity <x>: FinPlan '2BUS Business' sits under '1BUS Business'`** (also for SSM, EI and TEC): add `finplan_total = "1<CODE> <name>"` to that entity file, using the exact parent named in the error. Then append this to its `note`: `" FinPlan division total = <finplan_total> (incl. Ambition and IFRS adjustments); <finplan> is the operational unit."`
- **`entity ig`: 2IG sits under `2OTH Other entities`:** set `finplan_total = "2OTH Other entities"` and keep the note.
- **`entity consumer`:** it has no `finplan`, so there is nothing to fix.
- **`alias 'x' belongs to metric a and kpi b`:** remove the alias from the KPI (the metric keeps it).
- **`metric <bundle>: FinMax member 'ac.67410' is not in the atlas`:** look it up with `uv run smartass atlas status`, or search for it. If the member really is gone, drop it from `members` and add it to `description` as "(<member> no longer in the cube)".

Repeat the build until it reports `0 errors`.

- [ ] **Step 5: Live check of the FinPlan division totals (records the quirk as a fact)**

Use the smartass MCP tool `fin_analyze` (or ask the controller to do it if MCP is not available to you). FinPlan, metric `Revenue`, periods `closed-ytd`, entity `1BUS Business` and then `2BUS Business`; FinMax `revenue`, entity `un.SY.2BUS`. Note the three YTD values from `key_values`.

Write `knowledge/semantic/facts/finplan_division_levels.toml`:

```toml
text = "FinPlan divisions have two levels: 1<CODE> (division total, incl. Ambition 2INT_* and IFRS adjustments) and 2<CODE> (the operational unit). YTD <month> revenue Business: 1BUS <v1>, 2BUS <v2>, FinMax un.SY.2BUS <v3>; compare FinMax with <the closer one>."
severity = "warn"
source = "atlas 2026-11-10 + live check <date>"
status = "approved"
inherit = "down"
approved = { by = "auto", on = "<date>" }

[scope_members]
FinPlan = ["1BUS Business", "1SSM Security Solutions & Mobile", "1EI Entertainment & Internet", "1TEC Technology"]
```

Fill the `<…>` from the live values. If the live call is impossible, write the text without numbers ("which one matches FinMax is not verified") and set `severity = "warn"`.

- [ ] **Step 6: Tests + commit**

Run: `uv run python -B -m pytest -q tests < /dev/null | tail -1`
Expected: all pass.

```bash
git checkout engine/uv.lock
git add scripts/fold_knowledge.py knowledge/semantic
git commit -m "feat(knowledge): fold catalog, capabilities and the business maps into rules, entities, KPIs, vocab, notes and bundle metrics"
```

---

### Task 7: Resolver serves the new kinds; `fin_lookup` kinds

**Files:**
- Modify: `engine/smartass/knowledge/resolver.py`
- Modify: `engine/smartass/server.py` (`lookup`, `fin_lookup` docstring)
- Test: `engine/tests/v2/test_v2_knowledge_a.py`

- [ ] **Step 1: Write the failing tests** (append)

```python
def _resolver(tmp_path, adb, files):
    rep = _build(tmp_path, adb, files)
    assert rep.ok, rep.to_text()
    return KnowledgeResolver(sqlite3.connect(tmp_path / "knowledge.db"), adb, 26)


VAGUE = 'kind = "vague"\nwords = ["kostnader", "costs"]\noptions = ["opex", "opex + COGS"]\nsource = "s"\n'
NOTE = '[[note]]\ncube = "FinMax"\nmember = "ac.C30"\ntext = "Revenue incl. interconnect."\n'
SIGNS = 'title = "Signs"\ntext = "Costs are negative in FinMax."\n'


def test_search_finds_an_entity_by_alias_with_its_members(tmp_path, adb):
    r = _resolver(tmp_path, adb, {"entities/business.toml": BUS + 'finplan_total = "1BUS Business"\n'})
    (e,) = r.search("bedrift", kind="entity")["entities"]
    assert e["members"]["FinMax"] == "un.SY.2BUS" and e["finplan_total"] == "1BUS Business"
    assert e["finplan_total_children"] == ["2BUS Business", "2INT_B Ambition Business"]


def test_search_any_offers_the_options_of_a_vague_word(tmp_path, adb):
    r = _resolver(tmp_path, adb, {"vocab/costs.toml": VAGUE})
    (v,) = r.search("kostnader")["vocab"]
    assert v["options"] == ["opex", "opex + COGS"]


def test_a_member_hit_carries_its_note(tmp_path, adb):
    r = _resolver(tmp_path, adb, {"notes/finmax_accounts.toml": NOTE})
    assert r.note("ac.c30", "FinMax") == "Revenue incl. interconnect."         # any case, like Essbase
    assert r.note("ac.C70", "FinMax") == ""


def test_rules_are_listed_by_kind_rule(tmp_path, adb):
    r = _resolver(tmp_path, adb, {"rules/signs.toml": SIGNS})
    assert [x["id"] for x in r.search("", kind="rule")["rules"]] == ["signs"]
```

- [ ] **Step 2: Run them to see them fail**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py -k "search or note or rules_are" < /dev/null`
Expected: `KeyError: 'entities'`

- [ ] **Step 3: Implement in `KnowledgeResolver`**

In `__init__`, after the concepts:

```python
        from .semantic import Entity, Kpi, Rule, Vocab
        self._rules = {i: Rule.model_validate_json(js) for i, js in kdb.execute("SELECT id, json FROM rule ORDER BY id")}
        self._entities = {i: Entity.model_validate_json(js) for i, js in kdb.execute("SELECT id, json FROM entity")}
        self._kpis = {i: Kpi.model_validate_json(js) for i, js in kdb.execute("SELECT id, json FROM kpi")}
        self._vocab = {i: Vocab.model_validate_json(js) for i, js in kdb.execute("SELECT id, json FROM vocab")}
```

Add the methods:

```python
    def _terms_of(self, text: str, kind: str) -> list[str]:
        """Item ids of a kind whose term is the phrase (any case), else whose term appears in it as whole words."""
        t = text.strip()
        exact = [i for (i,) in self.kdb.execute("SELECT DISTINCT item_id FROM term WHERE kind=? AND text=? "
                                                 "COLLATE NOCASE", (kind, t))]
        if exact or not t:
            return exact
        words = f" {' '.join(re.findall(r'[\w&%]+', t.casefold()))} "
        return list(dict.fromkeys(i for term, i in self.kdb.execute("SELECT text, item_id FROM term WHERE kind=?",
                                                                    (kind,))
                                  if f" {' '.join(re.findall(r'[\w&%]+', term.casefold()))} " in words))

    def rule(self, rid: str):
        return self._rules.get(rid)

    def rule_of_kind(self, kind: str):
        return next((r for r in self._rules.values() if r.kind == kind), None)

    def _entity_hit(self, e) -> dict:
        out = {"id": e.id, "label": e.label, "aliases": e.aliases, "members": e.members(), "note": e.note,
               "abolished": e.abolished, "successors": e.successors,
               "caveats": self.caveats({c: {m} for c, m in e.members().items()})}
        if e.finplan_total:
            out["finplan_total"] = e.finplan_total
            out["finplan_total_children"] = list(self.children(e.finplan_total, "FinPlan"))
        return out

    def _kpi_hit(self, k) -> dict:
        return {**k.model_dump(exclude={"source", "approved", "verified"}),
                "caveats": self.caveats({}, {k.numerator, k.denominator})}

    def note(self, member: str, cube: str) -> str:
        row = self.kdb.execute("SELECT text FROM member_note WHERE cube=? AND member=? COLLATE NOCASE",
                               (cube, member)).fetchone()
        return row[0] if row else ""
```

Add `import re` at the top. Then extend `search`:
- In the member loop, add `"note": self.note(h["name"], h["connection"])` to each member dict.
- Add the new kinds:

```python
        entities = vocab = kpis = rules = []
        if kind in ("any", "entity"):
            entities = [self._entity_hit(self._entities[i]) for i in self._terms_of(text, "entity")][:limit]
        if kind in ("any", "kpi"):
            kpis = [self._kpi_hit(self._kpis[i]) for i in self._terms_of(text, "kpi")][:limit]
        if kind in ("any", "vocab"):
            vocab = [self._vocab[i].model_dump(exclude={"source", "approved", "verified"})
                     for i in self._terms_of(text, "vocab")][:limit]
        if kind == "rule":
            rules = [{"id": r.id, "title": r.title, "text": r.text} for r in self._rules.values()
                     if not text.strip() or text.strip().casefold() in f"{r.id} {r.title} {r.text}".casefold()]
        return {"metrics": metrics, "members": members, "concepts": concepts, "entities": entities, "kpis": kpis,
                "vocab": vocab, "rules": rules}
```

Update the `search` docstring: `kind: any | metric | member | concept | entity | kpi | vocab | rule (rule: only by kind; '' lists all)`.

- [ ] **Step 4: Server**

In `engine/smartass/server.py`, replace the `fin_lookup` docstring with:

```python
    """Look up before you build. kind 'metric' (name or alias, e.g. 'driftskostnader': members per cube, how the
    cubes differ), 'member' (cube members, each with its meaning note), 'entity' (a business area by name or alias:
    FinMax/FinPlan members, FinPlan division total, successors of abolished areas), 'kpi' (ratios: numerator,
    denominator, basis), 'vocab' (vague words -> the options to offer; 'per customer' etc. -> what is answerable),
    'concept' (product concepts such as 'splitt'), 'rule' (the always-true cube rules; query '' lists them),
    'answer' (past answers). 'any' = metric + member + concept + entity + kpi + vocab. Every hit carries its caveats
    (reorgs, abolished units, account traps; 'via' names the member a caveat was inherited from): pass them on.
    cube: FinMax | FinLite | FinPlan to narrow. Local, instant."""
```

The `lookup` docstring also changes: the result now has `entities`, `kpis`, `vocab` and `rules`.

- [ ] **Step 5: Run tests**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py tests/v2/test_v2_server.py tests/v2/test_v2_knowledge_resolver.py < /dev/null`
Expected: all pass. A test that pins the exact keys of `search()` must get the four new keys added. Update it and say so in the commit.

- [ ] **Step 6: Commit**

```bash
git add engine/smartass/knowledge/resolver.py engine/smartass/server.py engine/tests
git commit -m "feat(knowledge): fin_lookup finds entities, KPIs, vague words, rules and member notes"
```

---

### Task 8: Scenarios and default materiality from the rules (one definition)

**Files:**
- Modify: `engine/smartass/model/resolver.py` (Protocol, `DEFAULT_MATERIALITY`, `DictResolver`)
- Modify: `engine/smartass/model/series.py`
- Modify: `engine/smartass/knowledge/resolver.py` (`scenarios`, `materiality`)
- Test: `engine/tests/v2/test_v2_knowledge_a.py`

- [ ] **Step 1: Write the failing tests** (append)

```python
from smartass.model import series
from smartass.model.resolver import DictResolver, default_materiality, rule_data


def test_series_reads_the_scenarios_rule_from_the_repo():
    r = DictResolver(metrics=[], forecast="Forecast Q326")
    assert series.scenario("FinPlan", "budget", r) == ("Target", "Final")
    assert series.scenario("FinPlan", "forecast", r) == ("Forecast Q326", "Final")
    assert series.scenario("FinMax", "ly", r) == ("Actual", "")
    with pytest.raises(ValueError, match="only in FinPlan"):
        series.scenario("FinMax", "budget", r)


def test_raw_members_take_the_materiality_rule():
    assert default_materiality() == (rule_data("materiality")["abs_nok"], rule_data("materiality")["pct"])
```

Check the `DictResolver` constructor arguments in `engine/smartass/model/resolver.py` before you write the first test. Use the existing keyword names; `forecast=` is the attribute `latest_forecast()` returns.

- [ ] **Step 2: Run them to see them fail**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py -k "scenarios_rule or materiality_rule" < /dev/null`
Expected: ImportError on `default_materiality`.

- [ ] **Step 3: Implement**

In `engine/smartass/model/resolver.py`:

```python
import functools
import tomllib


@functools.cache
def rule_data(kind: str) -> dict:
    """The data of the one rule of this kind in the semantic folder (knowledge/semantic/rules/*.toml). The model
    reads the same file the build checks, so a rule is defined once."""
    from ..core import config
    for p in sorted((config.SEMANTIC_DIR / "rules").glob("*.toml")):
        d = tomllib.loads(p.read_text(encoding="utf-8"))
        if d.get("kind") == kind:
            return d["data"]
    raise LookupError(f"no rule of kind {kind!r} in {config.SEMANTIC_DIR / 'rules'}")


def default_materiality() -> tuple[float, float]:
    d = rule_data("materiality")
    return (float(d["abs_nok"]), float(d["pct"]))
```

Then make these changes:
- Remove `DEFAULT_MATERIALITY = (5_000_000.0, 0.05)`. Replace each use in this file with `default_materiality()`. For the dataclass field default, use `field(default_factory=default_materiality)`.
- Add `def scenarios(self) -> dict: ...` to the `Resolver` Protocol.
- Add to `DictResolver`: `def scenarios(self) -> dict: return rule_data("scenarios")`.
- Update the imports of `DEFAULT_MATERIALITY` elsewhere: `grep -rn DEFAULT_MATERIALITY engine/smartass engine/tests`, and replace each with `default_materiality()`.

In `engine/smartass/knowledge/resolver.py`:
- Make `scenarios()` return `self.rule_of_kind("scenarios").data`.
- Make `materiality()` for a raw member return `tuple(self.rule_of_kind("materiality").data[k] for k in ("abs_nok", "pct"))`.

Replace the body of `scenario` in `engine/smartass/model/series.py`:

```python
def scenario(cube: str, series: str, r: Resolver) -> tuple[str, str]:
    """(scenario, version) of a series in a cube, from the `scenarios` rule (knowledge/semantic/rules)."""
    if series not in ("value", "ly", "budget", "forecast"):
        raise ValueError(f"unknown series {series!r}")
    table = r.scenarios().get(cube, {})
    if series not in table:
        raise ValueError(f"{series} exists only in FinPlan")
    scen, version = table[series]
    return (r.latest_forecast() if scen == "latest_forecast" else scen, version)
```

Also update the module docstring: "the table lives in the `scenarios` rule".

- [ ] **Step 4: Run all tests**

Run: `uv run python -B -m pytest -q tests < /dev/null | tail -3`
Expected: all pass. Some tests may pin the old 5 % default on raw members (pct 0.05). The rule is Tom's approved 10 %; update those pins and name them in the commit body. Fakes that implement the Resolver protocol without `scenarios` fail with `AttributeError`. Add `def scenarios(self): return rule_data("scenarios")` to each fake.

- [ ] **Step 5: Goldens**

Run: `uv run python -B -m smartass.testset run < /dev/null | tail -12`
Expected: 9/9. If a golden differs (the 10 % default can flag a raw-member variance as material), STOP. Report the diff to the controller for Tom and do not re-pin it.

- [ ] **Step 6: Commit**

```bash
git checkout engine/uv.lock
git add engine/smartass engine/tests
git commit -m "refactor(model): scenarios and default materiality come from the rules (defined once)"
```

---

### Task 9: Freshness said out loud

**Files:**
- Modify: `engine/smartass/knowledge/resolver.py` (`open_resolver`)
- Modify: `engine/smartass/server.py` (`fin_status` output)
- Test: `engine/tests/v2/test_v2_knowledge_a.py`

- [ ] **Step 1: Write the failing test** (append)

```python
def test_a_stale_atlas_is_a_warning_on_the_answer(monkeypatch):
    from smartass.knowledge import resolver as kr
    monkeypatch.setattr(kr.atlas, "status", lambda db: {"oldest_age_days": 41.0, "stale": True, "built": []})
    assert kr.freshness_notes(object()) == [
        "The atlas is 41 days old (over 30): new or renamed members may be missing - run `smartass atlas build`."]
    monkeypatch.setattr(kr.atlas, "status", lambda db: {"oldest_age_days": 3.0, "stale": False, "built": []})
    assert kr.freshness_notes(object()) == []
```

- [ ] **Step 2: Run it to see it fail**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_a.py -k stale < /dev/null`
Expected: `AttributeError: ... freshness_notes`

- [ ] **Step 3: Implement**

In `engine/smartass/knowledge/resolver.py`:

```python
def freshness_notes(db) -> list[str]:
    """Warnings for the answer when the knowledge may be out of date (spec 2026-11-10 §6: warn, never stop)."""
    st = atlas.status(db)
    if st["stale"]:
        age = st["oldest_age_days"]
        return [f"The atlas is {age:.0f} days old (over {atlas.MAX_AGE_DAYS}): new or renamed members may be "
                "missing - run `smartass atlas build`."] if age is not None else [
                "The atlas has never been built - run `smartass atlas build`."]
    return []
```

In `open_resolver`, after `db = atlas.connect()`, add `if notes is not None: notes.extend(freshness_notes(db))`.

`run_prepare` already puts `notes` first in `pre.warnings`, and they reach the reply's `warnings`.

In `fin_status` (server.py), add `"knowledge_layer": {k: kn.get(k) for k in ("rules", "entities", "kpis", "vocab", "notes", "built_at", "atlas_built_at")}` to the returned dict.

- [ ] **Step 4: Run tests + commit**

Run: `uv run python -B -m pytest -q tests < /dev/null | tail -1`
Expected: all pass.

```bash
git add engine/smartass engine/tests
git commit -m "feat(knowledge): a stale atlas is a warning on every answer; fin_status shows the knowledge layer"
```

---

### Task 10: Always-on rules: short instructions, skill points to `fin_lookup`

**Files:**
- Modify: `engine/smartass/server.py` (`INSTRUCTIONS`)
- Modify: the smartass skill `SKILL.md` and `reference/cubes.md` if they are in the repo (`grep -rln "name: smartass" C:/dev/smartass --include=SKILL.md`). The installed copy under `~/.claude/skills/smartass` is Tom's setup: change it only through the repo's installer, never by hand.

- [ ] **Step 1: INSTRUCTIONS**

In `INSTRUCTIONS`, replace line `1) fin_lookup: ...` with:

```
1) fin_lookup first when a word is not an exact member: metrics, entities (business areas, FinPlan division totals),
KPIs, vague words (-> the options to ask), grains ('per customer' -> what is answerable), rules (kind 'rule').
Every hit carries caveats; 'via' = inherited from that member. Pass them on.
```

- [ ] **Step 2: Skill**

If the repo has the skill source, replace each restated rule in `SKILL.md` §6 with "(fin_lookup kind 'rule' has the full rules)". The rules in question are the P&L chain, the signs and the cube choice. Keep each one as one line, because they are the always-on set. In `reference/cubes.md`, delete stale example months such as "Sep 26" and point to `fin_lookup`.

- [ ] **Step 3: Test the server instructions are still under the limit**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_server.py < /dev/null`
Expected: pass. If a test pins INSTRUCTIONS text, update it.

- [ ] **Step 4: Commit**

```bash
git add -A engine/smartass/server.py engine/tests skills 2>/dev/null; git add -A engine/smartass/server.py engine/tests
git commit -m "docs(server): always-on rules short; everything else via fin_lookup"
```

---

### Task 11: Delete the folded sources

**Files:**
- Delete: `knowledge/catalog.toml`, `knowledge/capabilities.toml`, `knowledge/accounts_map.json`, `knowledge/org_map.json`, `knowledge/finplan_map.json`, `knowledge/build_accounts_map.py`, `knowledge/build_org_map.py`, `knowledge/build_finplan_map.py`, `engine/smartass/knowledge/seed.py`, `engine/tests/v2/test_v2_seed.py`, `scripts/fold_knowledge.py`
- Modify: `knowledge/review.py` (`JSON_OF` keeps only `products`)

- [ ] **Step 1: Confirm nothing else reads them**

Run: `grep -rln "catalog\.toml\|capabilities\.toml\|accounts_map\|org_map\|finplan_map\|knowledge\.seed\|from \.seed\|import seed" --include=*.py --include=*.ps1 --include=*.toml . | grep -v "docs/\|\.venv"`
Expected: only the files on the delete list, `knowledge/review.py`, and `build_products_map.py` (if it imports `build_accounts_map`). If `build_products_map.py` imports `OUT_DIR` from `build_accounts_map`, inline that constant in `build_products_map.py` before deleting.

- [ ] **Step 2: Delete and trim**

```bash
git rm knowledge/catalog.toml knowledge/capabilities.toml knowledge/accounts_map.json knowledge/org_map.json \
  knowledge/finplan_map.json knowledge/build_accounts_map.py knowledge/build_org_map.py knowledge/build_finplan_map.py \
  engine/smartass/knowledge/seed.py engine/tests/v2/test_v2_seed.py scripts/fold_knowledge.py
```

In `knowledge/review.py`, set `JSON_OF = {"products": ("products_map.json", "products", "member")}`. Change the docstring's area list to `products` only.

- [ ] **Step 2b: Check the installer and the packs**

Run: `grep -rn "catalog\|capabilities\|_map.json" installer scripts engine/smartass/ops 2>/dev/null`
Expected: no copy step for the deleted files. If `installer\install.ps1` copies them, remove those lines.

- [ ] **Step 3: Full tests, lint, goldens**

Run: `uv run ruff check . < /dev/null | tail -1 && uv run python -B -m pytest -q tests < /dev/null | tail -1 && uv run python -B -m smartass.testset run < /dev/null | tail -3`
Expected: `All checks passed!`, all tests pass, 9/9.

- [ ] **Step 4: Commit**

```bash
git checkout engine/uv.lock
git add -A
git commit -m "chore(knowledge): catalog, capabilities and the three business maps are folded into knowledge/semantic - deleted (one source)"
```

---

### Task 12: Docs, version, merge, CI

**Files:**
- Modify: `docs/DESIGN.md` (knowledge section), `CHANGELOG.md`, `docs/specs/2026-11-10-knowledge-layer-design.md` (status line: "phase A done <date>"; the §8 Phase A row now names notes and bundle metrics, as the refinements at the top of this plan describe)

- [ ] **Step 1: Docs**

1. **`docs/DESIGN.md`:** in the knowledge section, list the semantic folders (`metrics`, `facts` with `inherit`, `concepts`, `rules`, `entities`, `kpis`, `vocab`, `notes`) and say that `knowledge/semantic` is the only source.
2. **`CHANGELOG.md`:** add an entry for the version under it:
   - one source;
   - inherited caveats with `via`;
   - new `fin_lookup` kinds;
   - scenarios and materiality from rules (raw-member default 10 %);
   - stale-atlas warning;
   - FTE materiality 10 FTE / 5 %;
   - FinPlan division totals.

- [ ] **Step 2: Version bump**

Run (PowerShell, repo root): `scripts\bump.ps1` (it bumps 0.2.21 → 0.2.22 and commits; read the script's usage line first if it needs an argument)
Expected: a commit `chore(v2): version 0.2.22`.

- [ ] **Step 3: Merge and push**

```bash
cd engine && uv run ruff check . < /dev/null | tail -1 && uv run python -B -m pytest -q tests < /dev/null | tail -1 && cd ..
git checkout engine/uv.lock
git checkout main && git merge --no-ff v2/knowledge-a -m "Merge v2/knowledge-a: one knowledge source, inherited caveats, freshness"
git push -q && gh run watch $(gh run list --limit 1 --json databaseId -q '.[0].databaseId') --exit-status > /dev/null; gh run list --limit 1
```

Expected: `completed success`.

- [ ] **Step 4: Reload and smoke-test**

After the version bump the MCP server reloads (or call `fin_status(restart=True)`). Then:
- `fin_lookup("bedrift", kind="entity")`: the Business entity with `finplan_total` and the children.
- `fin_lookup("kostnader")`: the vague-word options.
- `fin_lookup("C40S", kind="member", cube="FinMax")`: the C40510 trap with `via`.

Report the results, plus the two open points for Tom: the FTE materiality (10 FTE / 5 %), and which FinPlan level matches FinMax (from Task 6 Step 5).
