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
