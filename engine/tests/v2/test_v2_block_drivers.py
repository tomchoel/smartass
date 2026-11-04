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
