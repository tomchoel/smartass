# engine/tests/v2/test_v2_block_drivers.py
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.drivers import build_drivers, DrillRequest
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def run(**opts):
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "drivers", "query": "fm", "metric": "Opex", **opts}]))
    c = Ctx(s, resolver(), CLOSED, facts())
    return c, build_drivers(s.blocks[0], c, "b0")


def test_drivers_walk_the_biggest_contributors():
    c, out = run(depth=2)
    # 2BUS -8: 4BUS_B -7 covers 7/8 >= 80 % so G910 is not shown; 4BUS_B: B1 -10, then B2 +3 (10/13 < 80 %)
    assert [(ln.key, ln.label, ln.level) for ln in out.lines] == [
        ("top", "2BUS Business Norway", 0), ("d0", "un.SY.4BUS_B", 1),
        ("d0_0", "un.SY.B1", 2), ("d0_1", "un.SY.B2", 2)]
    cells = evaluate([out], c.frame)
    assert [cells.get(ck("b0", k, 2)) for k in ("top", "d0", "d0_0", "d0_1")] == [-8.0, -7.0, -10.0, 3.0]
    assert out.drills == []


def test_drivers_request_drills_at_the_leaves():
    _, out = run(depth=2, drill=True)
    assert out.drills == [
        DrillRequest("fm", "FinMax", "un.SY.B1", "Opex", ("ac.C50S", "ac.C60"), ("Sep 26",), ("Sep 25",), "ly"),
        DrillRequest("fm", "FinMax", "un.SY.B2", "Opex", ("ac.C50S", "ac.C60"), ("Sep 26",), ("Sep 25",), "ly")]


def test_depth_one_stops_at_children():
    _, out = run(depth=1)
    assert [ln.key for ln in out.lines] == ["top", "d0"]


def test_drill_requests_of_a_stock_metric_carry_only_the_last_month():
    from v2util import monthly_facts
    d = spec_dict(blocks=[{"type": "drivers", "query": "fm", "metric": "FTE", "depth": 1, "drill": True}])
    d["queries"][0].update(metrics=["FTE"], periods="Jul 26..Sep 26")
    s = AnalysisSpec.model_validate(d)
    out = build_drivers(s.blocks[0], Ctx(s, resolver(), CLOSED, monthly_facts()), "b0")
    assert out.drills
    for dr in out.drills:
        assert dr.periods == ("Sep 26",) and dr.base_periods == ("Sep 25",)


def test_walk_accounts_other_respects_negative_weights_and_computes_variance():
    from smartass.model.frame import Fact, Frame
    from smartass.model.resolver import MetricDef, StaticResolver
    # ac.P = +ac.POS - ac.NEG + ac.REM => AC: 150 - 80 + 30 = 100, LY: 100 - 20 + 10 = 90
    r = StaticResolver(
        metrics=[MetricDef("Custom", {"FinMax": ("ac.P",)})],
        tree={"ac.P": ("ac.POS", "ac.NEG", "ac.REM")},
        weights={("ac.P", "ac.NEG"): -1.0},
        extra_members={"ac.P", "ac.POS", "ac.NEG", "ac.REM", "un.SY.2BUS"},
    )
    fr = Frame([
        Fact("FinMax", "fm", a, "un.SY.2BUS", per, "Actual", "", v)
        for a, ac_v, ly_v in (("ac.P", 100.0, 90.0), ("ac.POS", 150.0, 100.0), ("ac.NEG", 80.0, 20.0),
                              ("ac.REM", 30.0, 10.0))
        for per, v in (("Sep 26", ac_v), ("Sep 25", ly_v))
    ])
    s = AnalysisSpec.model_validate(spec_dict(blocks=[
        {"type": "drivers", "query": "fm", "metric": "Custom", "by": "accounts", "depth": 1, "include_other": True},
    ]))
    s.queries[0].metrics = ["Custom"]
    c = Ctx(s, r, CLOSED, fr)
    out = build_drivers(s.blocks[0], c, "b0")
    cells = evaluate([out], fr)
    other = next(ln for ln in out.lines if ln.key == "dother")
    # ac.NEG (|var|=60) and ac.POS (|var|=50) cover >= 80%; ac.REM left in Other accounts (1)
    assert [cells.get(ck("b0", other.key, ci)) for ci in range(4)] == [30.0, 10.0, 20.0, 2.0]
