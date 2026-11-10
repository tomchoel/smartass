"""Knowledge layer phase A (docs/plans/2026-11-10-knowledge-a-one-source.md)."""

import sqlite3

import pytest

from smartass.knowledge import atlas
from smartass.knowledge.build import build
from smartass.knowledge.expand import expand
from smartass.knowledge.resolver import KnowledgeResolver
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
