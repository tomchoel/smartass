"""KnowledgeResolver: the resolver-protocol tests of test_v2_atlas_resolver.py on the new class, plus caveats and
search."""
import pytest

from smartass.knowledge import atlas
from smartass.knowledge.build import build
from smartass.knowledge.resolver import KnowledgeResolver
from smartass.model.resolver import UnknownMetric

REVENUE = """
name = "Revenue"
aliases = ["revenue", "inntekter", "omsetning", "C30"]
description = "Revenue."
sign = "income"
materiality = [5000000, 0.1]
status = "approved"
[finmax]
members = ["ac.C30"]
"""
FTE = """
name = "FTE"
aliases = ["årsverk"]
description = "A level."
unit = "FTE"
balance = "Accumulated"
aggregation = "last"
status = "draft"
[finmax]
members = ["ac.STA20Y"]
"""
CONSUMER = """
text = "Consumer Norway was abolished on 1 Jun 2026."
valid_from = "2026-06"
severity = "warn"
source = "test"
status = "approved"
[scope_members]
FinMax = ["un.SY.2CON"]
"""
REVENUE_NOTE = """
text = "Revenue note."
scope_metrics = ["revenue"]
severity = "info"
source = "test"
status = "approved"
"""
BUS_BLOCK = """
text = "Business block until 2025."
valid_to = "2025-12"
severity = "block"
source = "test"
status = "approved"
[scope_members]
FinMax = ["un.SY.2BUS"]
"""
BUS_DRAFT = """
text = "Draft move of G910."
severity = "info"
source = "refresh"
status = "draft"
[scope_members]
FinMax = ["un.SY.G910"]
"""


@pytest.fixture
def dbs(tmp_path):
    con = atlas.connect(tmp_path / "atlas.db")
    rows = [("FinMax", "All Entities", "un.SY.2BUS", "2BUS Business Norway", "un.SY", 2, 1),
            ("FinMax", "All Entities", "un.SY.4BUS_B", "4BUS_B Business", "un.SY.2BUS", 3, 1),
            ("FinMax", "All Entities", "un.SY.G910", "G910 Talkmore Business", "un.SY.2BUS", 3, 2),
            ("FinMax", "All Entities", "un.SY.2CON", "2CON Consumer Norway", "un.SY", 2, 3),
            ("FinMax", "All Accounts", "ac.C30", "C30 REVENUES", "ac.C48", 3, 1),
            ("FinMax", "All Accounts", "ac.STA20Y", "STA20Y FTE", "ac.STA", 3, 1),
            ("FinMax", "All Accounts", "ac.X1", "X1 child of 2BUS name clash", "un.SY.2BUS", 4, 1),
            ("FinPlan", "Scenario", "Forecast Q226", "Forecast Q226", "Forecasts", 2, 1),
            ("FinPlan", "Scenario", "Forecast Q326", "Forecast Q326", "Forecasts", 2, 2)]
    con.executemany("INSERT INTO member (connection, dimension, name, label, parent, depth, ord) "
                    "VALUES (?,?,?,?,?,?,?)", rows)
    con.commit()
    sem = tmp_path / "semantic"
    for rel, text in {"metrics/revenue.toml": REVENUE, "metrics/fte.toml": FTE, "facts/abolished_consumer.toml":
                      CONSUMER, "facts/revenue_note.toml": REVENUE_NOTE, "facts/bus_block.toml": BUS_BLOCK,
                      "facts/bus_draft.toml": BUS_DRAFT}.items():
        (sem / rel).parent.mkdir(parents=True, exist_ok=True)
        (sem / rel).write_text(text, encoding="utf-8")
    from kutil import copy_required_rules
    copy_required_rules(sem)
    assert build(sem, tmp_path / "atlas.db", tmp_path / "knowledge.db").ok
    import sqlite3
    kdb = sqlite3.connect(tmp_path / "knowledge.db")
    yield kdb, con
    kdb.close()
    con.close()


def res(dbs, yy=26):
    kdb, con = dbs
    return KnowledgeResolver(kdb, con, forecast_year=yy)


# ---- the resolver protocol (as test_v2_atlas_resolver.py)

def test_children_stay_in_the_members_dimension(dbs):
    assert res(dbs).children("un.SY.2BUS", "FinMax") == ("un.SY.4BUS_B", "un.SY.G910")
    assert res(dbs).children("un.SY.G910", "FinMax") == ()


def test_exists_label_and_metrics(dbs):
    r = res(dbs)
    assert r.exists("un.sy.2bus", "FinMax") and not r.exists("un.SY.2BUS", "FinPlan")
    assert r.label("un.SY.2BUS", "FinMax") == "2BUS Business Norway"
    assert r.label("un.SY.ZZ", "FinMax") == "un.SY.ZZ"
    assert r.members("revenue", "FinMax") == ("ac.C30",)
    assert r.members("ac.C30", "FinMax") == ("ac.C30",)
    with pytest.raises(UnknownMetric):
        r.members("ac.NOPE", "FinMax")
    assert r.weight("un.SY.G910", "FinMax") == 1.0


def test_latest_forecast_of_the_year(dbs):
    assert res(dbs).latest_forecast() == "Forecast Q326"
    with pytest.raises(UnknownMetric, match="Forecast Q<q>25"):
        res(dbs, 25).latest_forecast()


def test_label_drops_a_trailing_parenthesis_suffix(dbs):
    _, con = dbs
    con.execute("INSERT INTO member (connection, dimension, name, label, parent, depth, ord) "
                "VALUES ('FinMax','All Entities','un.SY.2WHS','2WHS Wholesale (un.SY)','un.SY',2,9)")
    con.commit()
    r = res(dbs)
    assert r.label("un.SY.2WHS", "FinMax") == "2WHS Wholesale"
    assert r.label("un.SY.2BUS", "FinMax") == "2BUS Business Norway"


def test_canonical_spelling(dbs):
    r = res(dbs)
    assert r.canonical("un.sy.2bus", "FinMax") == "un.SY.2BUS"
    assert r.canonical("un.SY.ZZ", "FinMax") == "un.SY.ZZ"


# ---- metrics from knowledge.db

def test_metric_lookup_by_id_name_or_alias_any_case(dbs):
    r = res(dbs)
    for word in ("revenue", "Revenue", "OMSETNING", "c30", "Inntekter"):
        assert r.members(word, "FinMax") == ("ac.C30",), word
    assert r.sign("omsetning") == "income" and r.materiality("Revenue") == (5_000_000.0, 0.1)
    assert r.balance("årsverk") == "Accumulated" and r.aggregation("FTE") == "last"
    assert r.sign("ac.C30") == "neutral" and r.status("ac.C30") == "approved"   # raw member
    with pytest.raises(UnknownMetric, match="FinPlan"):
        r.members("revenue", "FinPlan")
    assert [m.name for m in r.metrics()] == ["FTE", "Revenue"]


def test_status_is_the_metrics_status(dbs):
    r = res(dbs)
    assert r.status("revenue") == "approved" and r.status("fte") == "draft"


# ---- caveats

def test_abolished_caveat_respects_valid_from(dbs):
    r = res(dbs)
    sep = r.caveats({"FinMax": {"un.SY.2CON"}}, set(), {(2026, 9)})
    assert [c["id"] for c in sep] == ["abolished_consumer"] and sep[0]["severity"] == "warn"
    assert "abolished" in sep[0]["text"]
    assert r.caveats({"FinMax": {"un.SY.2CON"}}, set(), {(2026, 1)}) == []
    assert r.caveats({"FinMax": {"un.sy.2con"}}, set(), {(2026, 1), (2026, 6)})[0]["id"] == "abolished_consumer"
    assert r.caveats({"FinPlan": {"un.SY.2CON"}}, set(), {(2026, 9)}) == []          # other cube


def test_caveats_sorted_deduplicated_and_by_metric(dbs):
    r = res(dbs)
    got = r.caveats({"FinMax": {"un.SY.2CON", "un.SY.2BUS", "ac.C30"}}, {"revenue"}, {(2025, 12), (2026, 9)})
    assert [c["id"] for c in got] == ["bus_block", "abolished_consumer", "revenue_note"]
    assert [c["severity"] for c in got] == ["block", "warn", "info"]
    assert r.caveats({}, {"revenue"}, {(2026, 9)})[0]["id"] == "revenue_note"
    assert r.caveats({"FinMax": {"un.SY.2BUS"}}, set(), {(2026, 1)}) == []          # valid_to 2025-12


def test_draft_fact_in_scope_adds_one_warn(dbs):
    got = res(dbs).caveats({"FinMax": {"un.SY.G910"}}, set(), {(2026, 9)})
    assert len(got) == 1 and got[0]["severity"] == "warn"
    assert "unapproved knowledge change touches this answer" in got[0]["text"]
    assert "Draft move" not in got[0]["text"]


# ---- search

def test_search_finds_a_metric_by_a_norwegian_alias(dbs):
    hits = res(dbs).search("omsetning")
    assert hits["metrics"][0]["id"] == "revenue"
    m = hits["metrics"][0]
    assert "inntekter" in m["aliases"] and m["description"] == "Revenue." and m["members"] == {"FinMax": ["ac.C30"]}
    assert m["caveats"][0]["id"] == "revenue_note"


def test_search_members_carry_caveats_without_month_filter(dbs):
    hits = res(dbs).search("2CON", kind="member")
    assert hits["metrics"] == []
    (m,) = [h for h in hits["members"] if h["name"] == "un.SY.2CON"]
    assert m["cube"] == "FinMax" and [c["id"] for c in m["caveats"]] == ["abolished_consumer"]
    assert res(dbs).search("revenue", kind="metric")["members"] == []


def test_metric_of_a_member_or_of_its_parent(dbs):
    _, con = dbs
    con.execute("INSERT INTO member (connection, dimension, name, label, parent, depth, ord) "
                "VALUES ('FinMax','All Accounts','ac.C3011','C3011 Subscriptions','ac.C30',4,1)")
    con.commit()
    r = res(dbs)
    assert r.metric_of("ac.C30", "FinMax") == "Revenue"
    assert r.metric_of("ac.c30", "FinMax") == "Revenue"
    assert r.metric_of("ac.C3011", "FinMax") == "Revenue"
    assert r.metric_of("ac.C48", "FinMax") is None
    assert r.metric_of("un.SY.2BUS", "FinMax") is None
