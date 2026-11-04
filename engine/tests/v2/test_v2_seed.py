import pytest

from smartass.core import config
from smartass.knowledge.interim import catalog_metrics
from smartass.knowledge.seed import seed
from smartass.knowledge.semantic import Fact, Metric, load_semantic, to_toml


@pytest.fixture(scope="module")
def seeded(tmp_path_factory):
    out = tmp_path_factory.mktemp("semantic")
    seed(config.KNOWLEDGE_DIR, out)
    return load_semantic(out)


def test_seed_covers_the_catalogue_and_loads(seeded):
    metrics, facts = seeded
    assert len(metrics) >= 11 and len(facts) >= 30
    names = {m.name for m in metrics}
    assert {m.name for m in catalog_metrics()} <= names
    assert all(m.status == "approved" and m.approved["by"] == "tom" for m in metrics)
    assert all(f.status == "approved" and f.source for f in facts)


def test_seeded_metrics_match_the_interim_metrics(seeded):
    by = {m.name: m for m in seeded[0]}
    for d in catalog_metrics():
        m = by[d.name]
        assert m.members() == d.members, d.name
        assert (m.sign, m.balance, m.aggregation, m.materiality) == (d.sign, d.balance, d.aggregation,
                                                                      d.materiality), d.name
    assert by["FTE"].unit == "FTE" and by["Opex"].unit == "NOK"


def test_opex_aliases_and_differs(seeded):
    opex = {m.id: m for m in seeded[0]}["opex"]
    assert {"opex", "C70", "driftskostnader", "driftskostnad"} <= set(opex.aliases)
    assert len({a.casefold() for a in opex.aliases}) == len(opex.aliases)       # FinPlan "OPEX" folds into "opex"
    gm = {m.id: m for m in seeded[0]}["gross_margin"]
    assert {"C48", "Gross_Profit", "bruttomargin"} <= set(gm.aliases)            # FinMax code + FinPlan member
    assert "FinPlan groups accounts differently" in opex.differs
    personnel = {m.id: m for m in seeded[0]}["personnel"]
    assert "lønnskostnader" in personnel.aliases and personnel.differs == ""     # no FinPlan member


def test_abolished_consumer_fact(seeded):
    facts = seeded[1]
    hits = [f for f in facts if "un.SY.2CON" in f.scope_members.get("FinMax", [])]
    assert len(hits) == 1
    f = hits[0]
    assert f.valid_from == "2026-06" and f.severity == "warn"
    assert "abolished on 1 Jun 2026" in f.text and "un.SY.2SSM" in f.text and "un.SY.2EI" in f.text


def test_fact_kinds(seeded):
    facts = {f.id: f for f in seeded[1]}
    reorg = [f for f in facts.values() if f.id.startswith("reorg_")]
    assert len(reorg) >= 10
    assert "un.SY.3KYS_C" in facts["reorg_cwk_kys"].scope_members["FinMax"]
    accounts = [f for f in facts.values() if f.id.startswith("account_")]
    assert len(accounts) == 14 and all(f.severity == "warn" for f in accounts)
    assert facts["account_c75"].scope_members["FinMax"] == ["ac.C75"]
    assert "Target" in facts["account_c75"].text and "avskrivninger" not in facts["account_c75"].text
    assert facts["finplan_banorway"].scope_members == {"FinPlan": ["BANORWAY"]}
    assert {"finplan_kpis_periodic", "finplan_no_actual_volumes", "finplan_target_stops_at_ebitda",
            "org_revenue_bearing_cost_centres", "org_adjustment_units"} <= set(facts)


def test_fact_texts_are_english(seeded):
    norwegian = ("Flyttet", "omarbeidet", "Ikke bokført", "Inntektsbærende", " og ", " ligger ")
    for f in seeded[1]:
        assert not any(w in f.text for w in norwegian), f.id


def test_to_toml_round_trips(tmp_path):
    m = Metric(id="opex", name="Opex", description='Say "hi"\nok', aliases=["opex", "lønn"],
               finmax={"members": ["ac.C70"]}, approved={"by": "tom", "on": "2026-10-15"}, status="approved")
    f = Fact(id="f1", text="t", source="s", scope_members={"FinMax": ["a", "b"]}, scope_metrics=["opex"])
    for item, kind in ((m, "metrics"), (f, "facts")):
        p = tmp_path / kind / f"{item.id}.toml"
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(to_toml(item), encoding="utf-8")
    (m2,), (f2,) = load_semantic(tmp_path)
    assert m2 == m and f2 == f
