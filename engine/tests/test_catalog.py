# engine/tests/test_catalog.py
import pytest

from smartview_mcp import catalog

OLD_PNL = [
    ("rev", "ac.C30", "line", None), ("cogs", "ac.C40S", "line", None), ("gm", "ac.C48", "sum", None),
    ("pers", "ac.C50S", "line", None), ("oth", "ac.C60", "line", None), ("opex", "ac.C70", "sum", None),
    ("boi", "ac.EBITDA_BOI", "sum", None), ("c73", "ac.C73", "line", None), ("ebitda", "ac.EBITDA", "grand", None),
    ("da", "ac.C75", "line", None), ("fte", "ac.STA20Y", "line", "tall1"),
]
OLD_PNL_CHECKS = [("C30 + C40S = C48", "rev+cogs-gm"), ("C50S + C60 = C70", "pers+oth-opex"),
                  ("C48 + C70 = EBITDA_BOI", "gm+opex-boi"), ("EBITDA_BOI + C73 = EBITDA", "boi+c73-ebitda")]
OLD_FP_PNL = [("rev", "Revenue", "line"), ("cogs", "COGS", "line"), ("gp", "Gross_Profit", "sum"),
              ("opex", "OPEX", "line"), ("boi", "EBITDA before Other Items", "sum"), ("oth", "Total Other Items", "line"),
              ("ebitda", "EBITDA", "grand")]
OLD_FP_CHECKS = [("Revenue + COGS = Gross profit", "rev+cogs-gp"),
                 ("Gross profit + OPEX = EBITDA before other items", "gp+opex-boi"),
                 ("EBITDA before other items + Other items = EBITDA", "boi+oth-ebitda")]


def test_repo_catalogue_carries_the_template_lists_unchanged():
    cat = catalog.load()
    assert [(x.key, x.account, x.kind, x.num) for x in cat.pnl["FinMax"]] == OLD_PNL
    assert [(c.label, c.expr) for c in cat.checks["FinMax"]] == OLD_PNL_CHECKS
    assert [(x.key, x.account, x.kind) for x in cat.pnl["FinPlan"]] == OLD_FP_PNL
    assert [(c.label, c.expr) for c in cat.checks["FinPlan"]] == OLD_FP_CHECKS


def test_repo_catalogue_has_areas_scenarios_kpis_materiality():
    cat = catalog.load()
    keys = {e.key for e in cat.entities}
    assert {"company", "wholesale", "business", "ssm", "ei", "technology", "it", "adm", "consumer"} <= keys
    assert cat.scenarios["budget"].scenario == "Target" and cat.scenarios["forecast"].rule == "latest_forecast"
    assert {k.key for k in cat.kpis} >= {"ebitda_margin", "opex_per_fte"}
    assert cat.materiality.pct == 0.10 and "ebitda" in cat.materiality.always


BASE = {
    "pnl": {"FinMax": [{"key": "rev", "account": "ac.C30", "kind": "line", "label": "Revenue"},
                       {"key": "ebitda", "account": "ac.EBITDA", "kind": "grand", "label": "EBITDA"}]},
    "checks": {"FinMax": [{"label": "x", "expr": "rev-ebitda"}]},
    "scenarios": {"actual": {"connection": "FinMax", "scenario": "Actual", "label": "Actual"}},
    "materiality": {"pct": 0.1, "abs_nok": 1, "always": ["rev"]},
    "entity": [{"key": "w", "label": "W", "finmax": "un.SY.2WHS", "aliases": ["wholesale"]}],
    "kpi": [{"key": "m", "label": "M", "connection": "FinMax", "numerator": "ebitda", "denominator": "rev",
             "unit": "pct"}],
}


def _with(path, value):
    import copy
    d = copy.deepcopy(BASE)
    node = d
    for k in path[:-1]:
        node = node[k]
    node[path[-1]] = value
    return d


@pytest.mark.parametrize("data, msg", [
    (_with(["pnl", "FinMax", 1, "kind"], "total"), "kind"),
    (_with(["pnl", "FinMax", 1, "key"], "rev"), "duplicate"),
    (_with(["checks", "FinMax", 0, "expr"], "rev-nope"), "nope"),
    (_with(["kpi", 0, "denominator"], "nope"), "nope"),
    (_with(["materiality", "always"], ["nope"]), "nope"),
    (_with(["entity"], BASE["entity"] * 2), "duplicate"),
    (_with(["scenarios", "actual"], {"connection": "FinMax", "label": "A"}), "scenario or rule"),
    (_with(["entity", 0, "successors"], ["nope"]), "successor 'nope'"),
    (_with(["entity", 0, "aliases"], "wholesale"), r"entity\[0\]: aliases must be a list"),
    (_with(["materiality", "always"], "rev"), "materiality: always must be a list"),
    (_with(["scenarios", "actual", "connection"], "Nope"), "scenario 'actual': connection 'Nope' is not a pnl cube"),
    (_with(["kpi", 0, "connection"], "Nope"), "kpi 'm': unknown connection 'Nope'"),
    (_with(["kpi", 0, "unit"], "eur"), "kpi 'm': unit 'eur'"),
    (_with(["kpi"], BASE["kpi"] * 2), "duplicate kpi key 'm'"),
    (_with(["pnl", "FinMax", 0, "better"], "up"), "FinMax line 'rev': better 'up'"),
    (_with(["entity", 0, "finmax"], None), "entity 'w' needs finmax or finplan"),
    (_with(["entity", 0, "abolished"], "1 June 2026"), "entity 'w': abolished '1 June 2026' is not an ISO date"),
    (_with(["pnl", "FinMax", 0, "kind"], "grand"), "FinMax has 2 grand lines"),
    (_with(["scenarios", "actual", "rule"], "latest_forecast"), "scenario 'actual' has both scenario and rule"),
    (_with(["materiality", "pct"], 0), r"materiality: pct 0.0 is not in \(0, 1\]"),
    (_with(["materiality", "pct"], 10), r"materiality: pct 10.0 is not in \(0, 1\]"),
    (_with(["materiality", "abs_nok"], -1), "materiality: abs_nok -1.0 is negative"),
    (_with(["materiality", "pct"], "ten"), "materiality: pct 'ten' is not a number"),
    (_with(["materiality", "abs_nok"], None), "materiality: abs_nok None is not a number"),
    (_with(["pnl", "FinMax", 0, "acount"], "x"), r"pnl.FinMax\[0\]: unknown field 'acount'"),
    (_with(["entity", 0, "alias"], ["x"]), r"entity\[0\]: unknown field 'alias'"),
    (_with(["kpi", 0, "nominator"], "rev"), r"kpi\[0\]: unknown field 'nominator'"),
    (_with(["scenarios", "actual", "versoin"], "x"), r"scenarios\[actual\]: unknown field 'versoin'"),
    (_with(["checks", "FinMax", 0, "exp"], "x"), r"checks.FinMax\[0\]: unknown field 'exp'"),
    (_with(["materiality", "pcts"], 1), "materiality: unknown field 'pcts'"),
    (_with(["kpi", 0], {"key": "m"}), r"kpi\[0\]: missing field"),
    (_with(["entity", 0], "w"), r"entity\[0\]: expected a table"),
    (_with(["checks", "FinMax", 0, "expr"], "rev*ebitda"), "FinMax check 'x': expression 'rev\*ebitda'"),
    (_with(["checks", "FinMax", 0, "expr"], "rev--ebitda"), "FinMax check 'x': expression"),
    (_with(["checks", "FinMax", 0, "expr"], "rev+1"), "FinMax check 'x': expression"),
])
def test_parse_rejects_broken_catalogues(data, msg):
    with pytest.raises(catalog.CatalogError, match=msg):
        catalog.parse(data)


@pytest.mark.parametrize("text, key", [
    ("P&L for Wholesale Sep 26", "wholesale"),
    ("hvordan går det med bedriftsmarkedet?", "business"),
    ("EBITDA i TV og internett hittil i år", "ei"),
    ("un.SY.2TEC", "technology"),
    ("2SSM Security Solutions & Mobile", "ssm"),
    ("it", "it"),
    ("Telenor Norge total", "company"),
])
def test_resolve_entity(text, key):
    assert catalog.resolve_entity(catalog.load(), text).key == key


@pytest.mark.parametrize("text", ["is it good?", "what about ei?", "something else entirely"])
def test_resolve_entity_needs_a_real_match(text):
    assert catalog.resolve_entity(catalog.load(), text) is None


@pytest.mark.parametrize("text, key", [
    ("total opex for SSM", "ssm"),
    ("Total for IT", "it"),
    ("total opex for Telenor Norge", "company"),
    ("Telenor Norge", "company"),
    ("kostnader i IT", "it"),
    ("kostnad per årsverk i IT", "it"),
    ("hva bruker bedriftsmarked på konsulenter?", "business"),
    ("EBITDA i grossistmarkedet", "wholesale"),
    ("kostnader i administrasjonen", "adm"),
    ("stabsenheter Q3", "adm"),
    ("admin opex", "adm"),
    ("tv & internett hittil i år", "ei"),
    ("opex 2NRD_N Sep 26", "nordics"),
    ("un.SY.2NRD_N", "nordics"),
    ("Consumer Q3 mot i fjor", "consumer"),
])
def test_resolve_entity_in_a_sentence(text, key):
    assert catalog.resolve_entity(catalog.load(), text).key == key


@pytest.mark.parametrize("text", ["mobil", "mobilkostnader i fjor", "Wholesale vs Business"])
def test_resolve_entity_none_when_not_one_area(text):
    assert catalog.resolve_entity(catalog.load(), text) is None


@pytest.mark.parametrize("text, keys", [
    ("Wholesale vs Business", ("wholesale", "business")),
    ("Business vs Wholesale vs Business", ("business", "wholesale")),
    ("IT-divisjonen mot Technology", ("it", "technology")),
    ("Security Solutions & Mobile og E&I", ("ssm", "ei")),
    ("Total for IT og SSM", ("it", "ssm")),
    ("total opex", ("company",)),
    ("nothing here", ()),
])
def test_resolve_entities_in_order_of_appearance(text, keys):
    assert tuple(e.key for e in catalog.resolve_entities(catalog.load(), text)) == keys


def test_abolished_consumer_names_its_successors():
    cat = catalog.load()
    con = catalog.resolve_entity(cat, "privatmarkedet")
    assert con.key == "consumer" and con.abolished == "2026-06-01" and con.successors == ("ssm", "ei")
    assert {e.key: e.successors for e in cat.entities}["ig"] == ()


def test_catalogue_tables_are_read_only():
    cat = catalog.parse(BASE)
    with pytest.raises(TypeError):
        cat.pnl["FinMax"] = ()
    with pytest.raises(TypeError):
        cat.scenarios["x"] = cat.scenarios["actual"]
    with pytest.raises(TypeError):
        cat.checks["FinMax"] = ()


def test_parse_accepts_the_base_and_a_toml_date():
    assert catalog.parse(BASE).entities[0].key == "w"
    import datetime
    cat = catalog.parse(_with(["entity", 0, "abolished"], datetime.date(2026, 6, 1)))
    assert cat.entities[0].abolished == "2026-06-01"
