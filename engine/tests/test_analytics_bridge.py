"""analytics.bridge and analytics.landing: bridges close exactly; the landing and its run-rate alternative."""
import pytest
from analytics_util import E, closes, frame

from smartview_mcp import catalog
from smartview_mcp.analytics.bridge import Bridge, by_line, by_unit, components
from smartview_mcp.analytics.facts import Fact, missing
from smartview_mcp.analytics.landing import landing, remaining_forecast
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
    assert b.residual() == 0.0 and closes(b)
    broken = by_line(f("ly", 0.0), f("cur", 999.0), vs, [(1, "rev")], {})
    assert closes(broken) and broken.steps[-1].key == "unexplained"       # the gap is shown, not hidden
    assert broken.steps[-1].label == "Unexplained"
    assert broken.steps[-1].fact.value == 999.0 - vs["rev"].abs.value
    assert [x.key for x in b.steps] == ["rev", "cogs", "pers", "oth", "c73"]   # nothing to explain: no extra step


def test_line_bridge_negative_leaf_and_missing_variance():
    lines = {x.key: x for x in CAT.pnl["FinMax"]}
    vs = {"rev": variance(lines["rev"], f("c", 10.0), f("p", 4.0), "rev")}
    b = by_line(f("s", 0.0), f("e", 0.0), vs, [(-1, "rev"), (1, "cogs")], {})
    assert b.steps[0].fact.value == -6.0 and b.steps[0].fact.definition.startswith("-(")     # the .neg path
    assert b.steps[1].fact.value is None and b.residual() is None and not closes(b)


def test_bridge_tolerance_is_one_nok():
    assert closes(by_line(f("s", 0.0), f("e", 10.0), {}, [], {}))      # a 10 gap becomes the rounding step
    assert closes(Bridge(f("s", 0.0), (), f("e", 0.9)))
    assert not closes(Bridge(f("s", 0.0), (), f("e", 1.1)))


def test_unit_bridge_top_n_plus_other_sums_exactly():
    units = [(f"u{i}", f"Unit {i}", f(f"c{i}", c), f(f"l{i}", p))
             for i, (c, p) in enumerate([(10.0, 0.0), (-50.0, 0.0), (5.0, 4.0), (30.0, 0.0), (1.0, 1.0)])]
    units.append(("ux", "Unknown", missing("cx", "nok", "", "not read"), f("lx", 1.0)))
    b = by_unit(f("start", 100.0), f("end", 103.0), units, top_n=2)
    assert [s.key for s in b.steps] == ["u1", "u3", "other"] and b.steps[-1].label == "Other units and eliminations"
    assert "Unknown" in b.steps[-1].fact.definition
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


def test_unit_bridge_omits_a_zero_other_step():
    units = [("a", "A", f("ca", 5.0), f("pa", 0.0)), ("b", "B", f("cb", 2.0), f("pb", 0.0))]
    b = by_unit(f("s", 0.0), f("e", 7.0), units, top_n=5)
    assert [x.key for x in b.steps] == ["a", "b"] and closes(b)


def test_landing_seasonal_share_is_bounded():
    # n = 9: the accepted share lies in 0.375-1.125
    over = landing(f("ytd", 900.0), f("rem", 250.0), f("tgt", 1200.0), f("a", 1200.0), f("b", 1000.0), 9)
    assert not over.seasonal and over.run_rate.value == 1200.0 and "outside" in over.method
    near0 = landing(f("ytd", 900.0), f("rem", 250.0), f("tgt", 1200.0), f("a", 10.0), f("b", 1000.0), 9)
    assert not near0.seasonal and "1.0%" in near0.method
    nofy = landing(f("ytd", 900.0), f("rem", 250.0), f("tgt", 1200.0), f("a", 750.0),
                   missing("b", "nok", "", "not read"), 9)
    assert not nofy.seasonal and "missing" in nofy.method and nofy.run_rate.value == 1200.0
    edge = landing(f("ytd", 900.0), f("rem", 0.0), f("tgt", 1.0), f("a", 375.0), f("b", 1000.0), 9)
    assert edge.seasonal                                                   # exactly 0.5 x 9/12 is accepted


def test_landing_months_done_zero_has_no_run_rate():
    z = landing(f("ytd", 0.0), f("rem", 100.0), f("tgt", 100.0), f("a", 1.0), f("b", 2.0), 0)
    assert z.run_rate.value is None and not z.seasonal and z.forecast.value == 100.0


def _stale(name, closed):
    return landing(f("ytd", 900.0), f("rem", 250.0), f("tgt", 1200.0), f("a", 750.0), f("b", 1000.0), 9,
                   name, closed)


def test_landing_names_the_forecast_and_warns_only_two_quarters_behind():
    old = _stale("Forecast Q126", "Sep 26")                              # Q1 against a Q3 close: two behind
    assert old.stale and old.warning == "Forecast Q126 is older than the closed month Sep 26"
    assert "Forecast Q126" in old.forecast.definition and "Forecast Q126" in old.method
    assert old.run_rate.value is not None                                  # the run rate stays alongside
    assert not _stale("Forecast Q226", "Sep 26").stale                     # one behind is normal
    assert not _stale("Forecast Q326", "Sep 26").stale and _stale("Forecast Q326", "Sep 26").warning is None
    assert _stale("Forecast Q226", "Oct 26").stale                         # Q2 with October (Q4) closed
    assert not _stale("Forecast Q326", "Oct 26").stale                     # Q3 with October closed is normal
    assert not _stale("Forecast Q425", "Jan 26").stale and _stale("Forecast Q325", "Jan 26").stale
    assert not _stale("not a forecast", "Sep 26").stale and not _stale("Forecast Q226", "nonsense").stale


def test_landing_seasonal_share_for_cost_units_with_negative_ebitda():
    ld = landing(f("ytd", -900.0), f("rem", -250.0), f("tgt", -1200.0), f("a", -750.0), f("b", -1000.0), 9)
    assert ld.seasonal and ld.run_rate.value == pytest.approx(-1200.0)
    mixed = landing(f("ytd", -900.0), f("rem", -250.0), f("tgt", -1200.0), f("a", 750.0), f("b", -1000.0), 9)
    assert not mixed.seasonal and "sign" in mixed.method
    zero = landing(f("ytd", 900.0), f("rem", 1.0), f("tgt", 1.0), f("a", 750.0), f("b", 0.0), 9)
    assert not zero.seasonal
    out = landing(f("ytd", -900.0), f("rem", -250.0), f("tgt", -1200.0), f("a", -1200.0), f("b", -1000.0), 9)
    assert not out.seasonal and "outside" in out.method


def test_bridge_unexplained_step_and_rounding_label():
    big = by_line(f("s", 0.0), f("e", 10.0), {}, [], {})
    assert big.steps[-1].key == "unexplained" and big.steps[-1].label == "Unexplained" and big.unexplained == 10.0
    tiny = by_line(f("s", 0.0), f("e", 0.4), {}, [], {})
    assert tiny.steps[-1].key == "rounding" and tiny.steps[-1].label == "Rounding" and tiny.unexplained == 0.0
    assert by_line(f("s", 1.0), f("e", 1.0), {}, [], {}).steps == ()


def test_unit_bridge_keeps_a_zero_other_when_some_units_are_unknown():
    units = [("a", "A", f("ca", 5.0), f("pa", 0.0)), ("x", "X", missing("cx", "nok", "", "gone"), f("px", 1.0))]
    b = by_unit(f("s", 0.0), f("e", 5.0), units)
    assert [x.key for x in b.steps] == ["a", "other"] and b.steps[-1].fact.value == 0.0


def test_remaining_forecast_blank_is_missing_not_zero():
    months = ["Oct 26", "Nov 26", "Dec 26"]
    blank = frame(values={(E, "ebitda", m, "forecast"): None for m in months})
    assert remaining_forecast(blank, E, "ebitda", "forecast", months).value is None
    part = frame(values={(E, "ebitda", "Oct 26", "forecast"): None, (E, "ebitda", "Nov 26", "forecast"): 5.0,
                         (E, "ebitda", "Dec 26", "forecast"): 7.0})
    assert remaining_forecast(part, E, "ebitda", "forecast", months).value == 12.0
    assert remaining_forecast(blank, E, "ebitda", "forecast", []).value == 0.0
    assert remaining_forecast(frame(), E, "ebitda", "forecast", months).value is None
