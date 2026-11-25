import pytest
from svxml_util import EchoTransport, atlas_with_dimensions
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.data.provider import SvHttpProvider
from smartass.epm import line_store, svhttp
from smartass.model.analyze import compute
from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.drivers import build_drivers
from smartass.model.blocks.table import build_table
from smartass.model.context import Ctx
from smartass.model.frame import Fact, Frame
from smartass.model.lines import LineSet
from smartass.model.requests import requests
from smartass.model.resolver import MetricDef, StaticResolver
from smartass.model.spec import AnalysisSpec
from smartass.render.layout import constant_columns


def _account_resolver():
    return StaticResolver(
        metrics=[MetricDef("COGS", {"FinMax": ("ac.C40S",)}, sign="cost")],
        tree={
            "un.SY.2WHS": ("un.SY.5BMT_T", "un.SY.5BMT_C"),
            "ac.C40S": ("ac.43200", "ac.41000"),
        },
        labels={
            "un.SY.2WHS": "2WHS Wholesale",
            "un.SY.5BMT_T": "5BMT_T Core",
            "un.SY.5BMT_C": "5BMT_C Carrier",
            "ac.C40S": "C40S COGS",
            "ac.43200": "43200 Power",
            "ac.41000": "41000 Interconnect",
        },
    )


def _account_facts():
    rows = [
        Fact("FinMax", "fm", "ac.43200", "un.SY.2WHS", "Sep 26", "Actual", "", -120.0),
        Fact("FinMax", "fm", "ac.41000", "un.SY.2WHS", "Sep 26", "Actual", "", -90.0),
        Fact("FinMax", "fm", "ac.C40S", "un.SY.2WHS", "Sep 26", "Actual", "", -210.0),
        Fact("FinMax", "fm", "ac.43200", "un.SY.2WHS", "Sep 25", "Actual", "", -100.0),
        Fact("FinMax", "fm", "ac.41000", "un.SY.2WHS", "Sep 25", "Actual", "", -80.0),
        Fact("FinMax", "fm", "ac.C40S", "un.SY.2WHS", "Sep 25", "Actual", "", -180.0),
    ]
    return Frame(rows)


def test_query_extra_dimensions_flow_through_requests_frame_and_constant_cols(tmp_path):
    from smartass.knowledge import atlas
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "title": "CP Costs", "metrics": ["Opex"]}])
    d["queries"][0]["dimensions"] = {"All Counterparts": ["cp.A", "cp.B"]}
    s = AnalysisSpec.model_validate(d)
    reqs = requests(s, resolver(), CLOSED)
    assert reqs[0].dimensions == (("All Counterparts", ("cp.A", "cp.B")),)

    db_path = atlas_with_dimensions(tmp_path / "a.db")
    with atlas.connect(db_path) as db:
        atlas.store(db, "FinMax", "All Counterparts", [
            {"name": "All Counterparts", "label": "All Counterparts", "parent": None, "depth": 0, "order": 0},
            {"name": "cp.A", "label": "cp.A", "parent": "All Counterparts", "depth": 1, "order": 1},
            {"name": "cp.B", "label": "cp.B", "parent": "All Counterparts", "depth": 1, "order": 2},
        ], 0.1)
    sv = svhttp.SvHttp(EchoTransport(lambda *a: 10.0), home="https://plan1.example", db_path=db_path)
    got = SvHttpProvider(sv).fetch(reqs[0])
    fr = Frame(got.facts)
    assert "All Counterparts" in fr.columns()
    assert fr.total({"All Counterparts": ("cp.A",)}) == pytest.approx(40.0)
    assert "All Counterparts" not in constant_columns(fr)


def test_table_accounts_and_entities_x_metrics():
    r = _account_resolver()
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "rows": "accounts", "metric": "COGS"}])
    d["queries"][0]["metrics"] = ["COGS"]
    d["queries"][0]["entity"] = "un.SY.2WHS"
    s = AnalysisSpec.model_validate(d)
    reqs = requests(s, r, CLOSED)
    assert set(reqs[0].accounts) >= {"ac.C40S", "ac.43200", "ac.41000"}
    out = build_table(s.blocks[0], Ctx(s, r, CLOSED, _account_facts()), "b0")
    assert [ln.key for ln in out.lines] == ["a0", "a1", "total", "check"]
    cells = evaluate([out], _account_facts())
    assert cells.get(ck("b0", "check", 0)) == 0.0

    # entities_x_metrics
    d2 = spec_dict(blocks=[{"type": "table", "query": "fm", "rows": "entities_x_metrics",
                            "metrics": ["Revenue", "EBITDA"], "columns": ["value", "ly", "var"]}])
    s2 = AnalysisSpec.model_validate(d2)
    out2 = build_table(s2.blocks[0], Ctx(s2, resolver(), CLOSED, facts()), "b0")
    assert any(ln.kind == "heading" and ln.entity == "un.SY.4BUS_B" for ln in out2.lines)
    assert any(ln.metric == "EBITDA" and ln.entity == "un.SY.4BUS_B" for ln in out2.lines)


def test_drivers_by_accounts_and_include_other():
    r = _account_resolver()
    d = spec_dict(blocks=[{"type": "drivers", "query": "fm", "metric": "COGS", "by": "accounts",
                            "cover": 1.0, "max_per_level": 5, "drill": True}])
    d["queries"][0]["metrics"] = ["COGS"]
    d["queries"][0]["entity"] = "un.SY.2WHS"
    s = AnalysisSpec.model_validate(d)
    out = build_drivers(s.blocks[0], Ctx(s, r, CLOSED, _account_facts()), "b0")
    assert [ln.metric for ln in out.lines] == ["COGS", "ac.43200", "ac.41000"]
    assert out.drills[0].accounts == ("ac.43200",)

    # include_other surfaces skipped offsetting entity drivers
    d2 = spec_dict(blocks=[{"type": "drivers", "query": "fm", "metric": "Opex", "depth": 1, "include_other": True}])
    s2 = AnalysisSpec.model_validate(d2)
    out2 = build_drivers(s2.blocks[0], Ctx(s2, resolver(), CLOSED, facts()), "b0")
    assert any("Other" in ln.label for ln in out2.lines)


def test_lines_block_filters_multi_group_and_drill_id_cache(tmp_path):
    rows = [
        {"Period": "Jan 26", "Entity": "un.SY.VB20", "CC_DESC": "Core", "ACCOUNT": "43200",
         "ACCOUNT_NAME": "Power", "AP_CUSTOMER_NAME": "Statkraft", "AMOUNT": -400.0},
        {"Period": "Jan 26", "Entity": "un.SY.VB20", "CC_DESC": "Core", "ACCOUNT": "43200",
         "ACCOUNT_NAME": "Power", "AP_CUSTOMER_NAME": "Statkraft", "AMOUNT": 400.0},
        {"Period": "Nov 26", "Entity": "un.SY.VB20", "CC_DESC": "Core", "ACCOUNT": "43200",
         "ACCOUNT_NAME": "Power", "AP_CUSTOMER_NAME": "OtherVendor", "AMOUNT": -100.0},
        {"Period": "Mar 26", "Entity": "un.SY.VB50", "CC_DESC": "Edge", "ACCOUNT": "41000",
         "ACCOUNT_NAME": "Interconnect", "AP_CUSTOMER_NAME": "Telia", "AMOUNT": -250.0},
    ]
    d = spec_dict(blocks=[{
        "type": "lines", "cell": {"All Accounts": "ac.C40S", "All Time Periods": "Yr 26"},
        "group_by": ["ACCOUNT", "Entity"],
        "where": {"period_from": "Jan 26", "period_to": "Oct 26"},
        "exclude_reversals": True,
    }])
    s = AnalysisSpec.model_validate(d)
    fr = Frame([])
    fr.lines["b0"] = LineSet("b0", rows, list(rows[0].keys()), cube_value=-350.0)
    res = compute(s, resolver(), CLOSED, fr)
    o = res.outs[0]
    assert o.lines[0].label == "41000 · un.SY.VB50"
    assert res.cells.get(ck("b0", "total", 0)) == -250.0
    assert res.key_values and len(res.key_values) >= 1

    # drill_id validated without cell + line_store round-trip
    db_path = tmp_path / "ls.db"
    did = line_store.store_drill({"All Accounts": "ac.C40S", "All Time Periods": "Yr 26"},
                                 rows, list(rows[0].keys()), cube_value=-350.0,
                                 reconciliation={"cube_value": -350.0, "lines_sum": -350.0, "gap": 0.0,
                                                 "status": "exact", "lines": 4},
                                 warnings=[], db_path=db_path)
    meta, loaded = line_store.load_rows(did, where={"period_to": "Oct 26"}, exclude_reversals=True, db_path=db_path)
    assert meta["drill_id"] == did and len(loaded) == 1 and loaded[0]["ACCOUNT"] == "41000"
