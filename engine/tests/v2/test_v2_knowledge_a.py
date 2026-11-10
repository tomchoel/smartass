"""Knowledge layer phase A (docs/plans/2026-11-10-knowledge-a-one-source.md)."""

import sqlite3

import pytest

from smartass.knowledge import atlas
from smartass.knowledge.build import build
from smartass.knowledge.expand import expand
from smartass.knowledge.resolver import KnowledgeResolver
from smartass.knowledge.semantic import Entity, Fact, Kpi, MemberNote, Rule, SemanticError, Vocab, load_all

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


def test_direct_members_are_never_via_even_below_another_named_member(adb):
    got = expand(_fact("down", ["un.SY.2CON", "CC1"]), adb, stops={})
    assert got[("FinMax", "CC1")] == "CC1" and got[("FinMax", "CC2")] == "un.SY.2CON"


def test_the_atlas_spelling_is_stored(adb):
    assert set(expand(_fact("self", ["UN.sy.2con"]), adb, stops={})) == {("FinMax", "un.SY.2CON")}


def test_a_self_parent_row_cannot_loop(adb):
    adb.execute("INSERT INTO member VALUES ('FinMax','All Entities','LOOP','Loop','LOOP',2,9)")
    got = expand(_fact("up", ["LOOP"]), adb, stops={})
    assert set(got) == {("FinMax", "LOOP")}


def test_a_two_member_cycle_cannot_loop(adb):
    adb.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)",
                    [("FinMax", "All Entities", "A", "A", "B", 2, 8), ("FinMax", "All Entities", "B", "B", "A", 2, 9)])
    assert set(expand(_fact("up", ["A"]), adb, stops={})) == {("FinMax", "A"), ("FinMax", "B")}


def test_a_shared_member_climbs_every_parent(adb):
    adb.execute("INSERT INTO member VALUES ('FinMax','All Entities','CC1','CC1 unit','un.SY',3,3)")
    got = expand(_fact("up", ["CC1"]), adb, stops={"FinMax": {"un.SY"}})
    assert {m for _c, m in got} == {"CC1", "un.SY.2CON", "un.SY"}


def test_ancestors_stay_in_the_dimension_of_the_placement(adb):
    adb.execute("INSERT INTO member VALUES ('FinMax','Other','CC1','CC1','OTHERP',2,1)")
    got = expand(_fact("up", ["ac.C40510"]), adb, stops={})
    assert {m for _c, m in got} == {"ac.C40510", "ac.C405"}


def test_a_large_tree_expands_without_per_member_sql(adb):
    adb.executemany("INSERT INTO member VALUES ('FinMax','All Entities',?,?,'CC2',4,?)",
                    [(f"X{i}", f"X{i}", i) for i in range(20000)])
    calls = []
    adb.set_trace_callback(calls.append)
    import time
    t0 = time.perf_counter()
    got = expand(_fact("down", ["un.SY"]), adb, stops={})
    assert len(got) > 20000 and time.perf_counter() - t0 < 1.5
    assert len(calls) <= 2


def test_the_build_warns_when_a_fact_reaches_too_many_members(tmp_path, adb, monkeypatch):
    from smartass.knowledge import build as b
    monkeypatch.setattr(b, "MAX_SCOPE", 2)
    sem = _write(tmp_path / "semantic", {"metrics/cogs.toml": COGS, "facts/reorg.toml": REORG})
    rep = build(sem, tmp_path / "atlas.db", tmp_path / "k.db")
    assert rep.ok and any("expands to 3 FinMax members" in w for w in rep.warnings)


def test_a_fact_that_names_the_touched_member_and_an_ancestor_says_no_via(tmp_path, adb):
    sem = _write(tmp_path / "semantic", {"metrics/cogs.toml": COGS, "facts/reorg.toml": REORG})
    build(sem, tmp_path / "atlas.db", tmp_path / "k.db")
    r = KnowledgeResolver(sqlite3.connect(tmp_path / "k.db"), adb, 26)
    (c,) = r.caveats({"FinMax": {"CC2", "un.SY.2CON"}})
    assert "via" not in c


def test_ensure_built_rebuilds_on_a_newer_atlas_only(tmp_path, monkeypatch):
    from kutil import knowledge

    from smartass.knowledge.build import summary
    from smartass.knowledge.resolver import REBUILT_NOTE, ensure_built
    a, _, k = knowledge(tmp_path, monkeypatch)
    con = sqlite3.connect(a)
    con.execute("DELETE FROM build")
    con.execute("INSERT INTO build VALUES ('FinMax','All Accounts',1790000000.0,9,1.0,NULL)")
    con.commit()
    notes: list[str] = []
    ensure_built(notes)                              # first call: the stamp differs from the one built without a build row
    notes.clear()
    ensure_built(notes)
    assert notes == []
    con.execute("UPDATE build SET built_at=1790000500.5")
    con.commit()
    ensure_built(notes)
    con.close()
    assert notes == [REBUILT_NOTE] and summary(k)["atlas_built_ts"] == "1790000500.5"



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
    adb.executemany("INSERT OR IGNORE INTO member VALUES (?,?,?,?,?,?,?)", PLAN_TREE)
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


# ---- review fixes (entity dimensions, division totals, one term list, every kind approvable)

def test_an_entity_member_must_sit_in_the_entity_dimension(tmp_path, adb):
    ent = 'label = "Rev"\nfinmax = "ac.C30"\nsource = "s"\n'          # an account, not an entity
    rep = _build(tmp_path, adb, {"entities/rev_unit.toml": ent})
    assert any("ac.C30" in e and "not in the atlas" in e for e in rep.errors)


PLAN_TREE_DEEP = [("FinPlan", "Entity", "1X Division", "1X Division", "BANORWAY", 2, 5),
                  ("FinPlan", "Entity", "2X Mid", "2X Mid", "1X Division", 3, 5),
                  ("FinPlan", "Entity", "3X Leaf", "3X Leaf", "2X Mid", 4, 5),
                  ("FinPlan", "Entity", "OTHERROOT", "OTHERROOT", "Entity", 1, 6),
                  ("FinPlan", "Entity", "9Y Stray", "9Y Stray", "OTHERROOT", 2, 6)]


def _deep(tmp_path, adb, finplan, total=""):
    adb.executemany("INSERT OR IGNORE INTO member VALUES (?,?,?,?,?,?,?)", PLAN_TREE_DEEP)
    ent = f'label = "Deep"\nfinplan = "{finplan}"\nsource = "s"\n' + (f'finplan_total = "{total}"\n' if total else "")
    return _build(tmp_path, adb, {"entities/deep.toml": ent})


def test_a_deep_finplan_entity_needs_the_level_one_total(tmp_path, adb):
    rep = _deep(tmp_path, adb, "3X Leaf")
    assert any("division total" in e and "1X Division" in e for e in rep.errors)
    assert _deep(tmp_path, adb, "3X Leaf", "1X Division").ok
    assert _deep(tmp_path, adb, "1X Division").ok                      # the division itself
    assert any("finplan_total" in e and "Nope" in e for e in _deep(tmp_path, adb, "3X Leaf", "Nope").errors)


def test_a_finplan_chain_that_never_reaches_banorway_is_an_error(tmp_path, adb):
    rep = _deep(tmp_path, adb, "9Y Stray")
    assert any("OTHERROOT" in e and "BANORWAY" in e for e in rep.errors)


def test_the_same_term_across_kinds_is_an_error(tmp_path, adb):
    ent = 'label = "COGS"\nfinmax = "un.SY.2BUS"\nsource = "s"\n'
    assert any("'cogs'" in e.casefold() for e in _build(tmp_path, adb, {"entities/c.toml": ent}).errors)
    voc = 'kind = "vague"\nwords = ["business"]\noptions = ["a", "b"]\nsource = "s"\n'
    rep = _build(tmp_path, adb, {"entities/business.toml": BUS.replace('aliases = ["bedrift"]\n', ''),
                                 "vocab/v.toml": voc})
    assert any("'business'" in e for e in rep.errors)


def test_pnl_rule_ids_are_checked_and_notes_duplicates_warn(tmp_path, adb):
    pnl = ('title = "P"\ntext = "t"\nkind = "pnl"\n[data]\nFinMax = ["revenue", "ghost"]\nFinPlan = ["revenue"]\n')
    notes = ('[[note]]\ncube = "FinMax"\nmember = "ac.C30"\ntext = "a"\n'
             '[[note]]\ncube = "FinMax"\nmember = "ac.c30"\ntext = "b"\n')
    rep = _build(tmp_path, adb, {"rules/pnl.toml": pnl})
    assert any("ghost" in e for e in rep.errors)
    rep = _build(tmp_path, adb, {"notes/m.toml": notes})
    assert any("2 notes" in w and "ac.C30" in w for w in rep.warnings)


def test_rule_data_is_strict():
    with pytest.raises(ValueError):
        Rule(id="m", title="M", text="t", kind="materiality", data={"abs_nok": True, "pct": 0.1})
    with pytest.raises(ValueError):
        Rule(id="m", title="M", text="t", kind="materiality", data={"abs_nok": 0, "pct": 0.1})
    sc = {"FinMax": {"value": ["Actual", ""], "ly": ["Actual", ""]},
          "FinPlan": {"value": ["Actual", "Working"], "ly": ["Actual", "Working"], "budget": ["Target", "Final"],
                      "forecast": ["latest_forecast", "Final"]}}
    with pytest.raises(ValueError):
        Rule(id="s", title="S", text="t", kind="scenarios", data={**sc, "FinXx": sc["FinMax"]})
    with pytest.raises(ValueError):
        Rule(id="s", title="S", text="t", kind="scenarios", data={**sc, "FinMax": {"value": ["A", 1], "ly": ["A", ""]}})
    with pytest.raises(ValueError):
        Vocab(id="v", kind="grain", words=[], source="s")


def test_items_default_to_draft():
    assert Entity(id="x", label="X", finmax="un.SY").status == "draft"


def test_approve_and_status_cover_every_kind(tmp_path, adb):
    from smartass.knowledge.refresh import approve, status
    voc = 'kind = "vague"\nwords = ["costs"]\noptions = ["a", "b"]\nsource = "s"\n'
    rule = 'title = "Signs"\ntext = "Costs negative."\n'
    sem = _write(tmp_path / "semantic", {"metrics/cogs.toml": COGS, "vocab/costs.toml": voc,
                                          "rules/signs.toml": rule, "entities/bus.toml": BUS})
    adb.executemany("INSERT OR IGNORE INTO member VALUES (?,?,?,?,?,?,?)", PLAN_TREE)
    adb.commit()
    assert status(sem, tmp_path / "k.db", tmp_path)["drafts"] == ["bus", "costs", "signs"]
    done, rep = approve(sem, tmp_path / "atlas.db", tmp_path / "k.db", ["costs", "signs"], "tom", "2026-11-10")
    assert done == ["costs", "signs"]
    assert 'status = "approved"' in (sem / "vocab" / "costs.toml").read_text(encoding="utf-8")
    assert load_all(sem).vocab[0].approved == {"by": "tom", "on": "2026-11-10"}


# ---- Task 7: the resolver serves the new kinds

def _resolver(tmp_path, adb, files):
    rep = _build(tmp_path, adb, files)
    assert rep.ok, rep.to_text()
    return KnowledgeResolver(sqlite3.connect(tmp_path / "knowledge.db"), adb, 26)


VAGUE = ('kind = "vague"\nwords = ["kostnader", "costs"]\noptions = ["opex", "opex + COGS"]\nsource = "s"\n'
         'status = "approved"\n')
NOTE = '[[note]]\ncube = "FinMax"\nmember = "ac.C30"\ntext = "Revenue incl. interconnect."\n'
SIGNS = 'title = "Signs"\ntext = "Costs are negative in FinMax."\nstatus = "approved"\n'


def test_search_finds_an_entity_by_alias_with_its_members(tmp_path, adb):
    r = _resolver(tmp_path, adb, {"entities/business.toml": BUS + 'finplan_total = "1BUS Business"\nstatus = "approved"\n'})
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
    (m,) = [x for x in r.search("C30 Revenue", kind="member", cube="FinMax")["members"] if x["name"] == "ac.C30"]
    assert m["note"] == "Revenue incl. interconnect."


def test_rules_are_listed_by_kind_rule(tmp_path, adb):
    r = _resolver(tmp_path, adb, {"rules/signs.toml": SIGNS})
    assert [x["id"] for x in r.search("", kind="rule")["rules"]] == ["signs"]
    assert r.search("", kind="any")["rules"] == []
