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
