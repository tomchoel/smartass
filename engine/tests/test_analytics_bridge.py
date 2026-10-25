"""analytics.bridge and analytics.landing: bridges close exactly; the landing and its run-rate alternative."""
import pytest

from smartview_mcp import catalog
from smartview_mcp.analytics.bridge import by_line, by_unit, components
from smartview_mcp.analytics.facts import Fact, missing
from smartview_mcp.analytics.landing import landing
from smartview_mcp.analytics.variance import variance

CAT = catalog.load()


def f(id, value):
    return Fact(id, value, "nok", id)


def test_components_follow_the_catalogue_chain():
    fm = components(CAT.checks["FinMax"], CAT.pnl["FinMax"], "ebitda")
    assert fm == [(1, "rev"), (1, "cogs"), (1, "pers"), (1, "oth"), (1, "c73")]
    fp = components(CAT.checks["FinPlan"], CAT.pnl["FinPlan"], "ebitda")
    assert [k for _, k in fp] == ["rev", "cogs", "opex", "oth"]
    with pytest.raises(ValueError, match="no identity check defines 'rev'"):
        components(CAT.checks["FinMax"], CAT.pnl["FinMax"], "rev")


def test_line_bridge_closes_when_the_identities_hold():
    lines = {x.key: x for x in CAT.pnl["FinMax"]}
    cur = {"rev": 1000.0, "cogs": -400.0, "pers": -200.0, "oth": -100.0, "c73": -10.0}
    ly = {"rev": 900.0, "cogs": -350.0, "pers": -210.0, "oth": -80.0, "c73": 0.0}
    vs = {k: variance(lines[k], f(f"c.{k}", cur[k]), f(f"l.{k}", ly[k]), k) for k in cur}
    b = by_line(f("ly", sum(ly.values())), f("cur", sum(cur.values())), vs,
                components(CAT.checks["FinMax"], CAT.pnl["FinMax"], "ebitda"), {"rev": "Revenue"})
    assert [s.key for s in b.steps] == ["rev", "cogs", "pers", "oth", "c73"] and b.steps[0].label == "Revenue"
    assert b.residual() == 0.0 and b.closes()
    broken = by_line(f("ly", 0.0), f("cur", 999.0), vs, [(1, "rev")], {})
    assert not broken.closes()


def test_unit_bridge_top_n_plus_other_sums_exactly():
    units = [(f"u{i}", f"Unit {i}", f(f"c{i}", c), f(f"l{i}", p))
             for i, (c, p) in enumerate([(10.0, 0.0), (-50.0, 0.0), (5.0, 4.0), (30.0, 0.0), (1.0, 1.0)])]
    units.append(("ux", "Unknown", missing("cx", "nok", "", "not read"), f("lx", 1.0)))
    b = by_unit(f("start", 100.0), f("end", 103.0), units, top_n=2)
    assert [s.key for s in b.steps] == ["u1", "u3", "other"]
    assert b.steps[-1].fact.value == 3.0 - (-50.0 + 30.0) and b.residual() == 0.0
    assert by_unit(missing("s", "nok", "", "x"), f("end", 1.0), units).steps[-1].fact.value is None


def test_landing_forecast_and_seasonal_run_rate():
    ld = landing(f("ytd", 900.0), f("rem", 250.0), f("tgt", 1200.0), f("ly_ytd", 750.0), f("ly_fy", 1000.0), 9)
    assert ld.forecast.value == 1150.0 and ld.gap_forecast.value == -50.0
    assert ld.run_rate.value == pytest.approx(900.0 / 0.75) and ld.seasonal and "75.0%" in ld.method
    assert ld.gap_run_rate.value == pytest.approx(1200.0 - 1200.0)


def test_landing_falls_back_to_a_plain_run_rate():
    ld = landing(f("ytd", 900.0), f("rem", 250.0), f("tgt", 1200.0), f("ly_ytd", 750.0), f("ly_fy", 0.0), 9)
    assert ld.run_rate.value == 1200.0 and not ld.seasonal and "YTD / 9 x 12" in ld.method
    neg = landing(f("ytd", 900.0), f("rem", 0.0), f("tgt", 1.0), f("ly_ytd", 750.0), f("ly_fy", -10.0), 9)
    assert not neg.seasonal                                       # a negative share is no seasonality profile
    gone = landing(missing("ytd", "nok", "", "not read"), f("rem", 1.0), f("tgt", 1.0), f("a", 1.0), f("b", 2.0), 9)
    assert gone.forecast.value is None and gone.run_rate.value is None and gone.gap_forecast.value is None
