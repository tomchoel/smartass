"""analytics.measures and analytics.kpis: flows sum, stocks take levels, the n+1 average, annualising."""
import dataclasses
import math

import pytest
from analytics_util import E, frame, months

from smartview_mcp import catalog
from smartview_mcp.analytics import periods
from smartview_mcp.analytics.facts import Fact, total
from smartview_mcp.analytics.frame import from_rows
from smartview_mcp.analytics.kpis import kpi
from smartview_mcp.analytics.measures import amount, average_level, flow, level


def _fte_frame():
    v = months({}, "fte", "actual", "Dec 25", [10.0, 12.0, 14.0, 16.0])      # Dec 25 opening, Jan..Mar 26
    months(v, "opex", "actual", "Jan 26", [-300.0, -300.0, -300.0])
    months(v, "rev", "actual", "Jan 26", [1000.0, 0.0, 1000.0])
    months(v, "ebitda", "actual", "Jan 26", [100.0, -50.0, 150.0])
    return frame(values=v)


def test_flow_sums_months_and_names_its_inputs():
    f = flow(_fte_frame(), E, "opex", "actual", ["Jan 26", "Feb 26", "Mar 26"])
    assert f.value == -900.0 and f.unit == "nok" and len(f.inputs) == 3
    assert flow(_fte_frame(), E, "opex", "actual", []).value == 0.0
    gone = flow(_fte_frame(), E, "opex", "actual", ["Mar 26", "Apr 26"])
    assert gone.value is None and "Apr 26" in gone.reason


def test_stock_level_and_the_n_plus_1_average():
    fr = _fte_frame()
    assert level(fr, E, "fte", "actual", "Mar 26").value == 16.0
    avg = average_level(fr, E, "fte", "actual", ["Jan 26", "Feb 26", "Mar 26"])
    assert avg.value == (10 + 12 + 14 + 16) / 4 and avg.unit == "fte" and len(avg.inputs) == 4
    assert "4 levels (Dec 25 opening + 3 month ends)" in avg.definition
    assert amount(fr, E, "fte", "actual", ["Jan 26", "Feb 26", "Mar 26"]).value == 16.0   # stock: period end
    assert average_level(fr, E, "fte", "actual", ["Jan 25"]).value is None               # Dec 24 never read


def test_kpis_margin_and_annualised_cost_per_fte():
    cat = catalog.load()
    k = {x.key: x for x in cat.kpis}
    fr = _fte_frame()
    q1 = ["Jan 26", "Feb 26", "Mar 26"]
    m = kpi(fr, k["ebitda_margin"], E, "actual", q1)
    assert m.value == 200.0 / 2000.0 and m.unit == "pct"
    c = kpi(fr, k["opex_per_fte"], E, "actual", q1)
    assert c.value == 900.0 / 13.0 * 12 / 3 and c.unit == "nok_per_fte" and "x 12/3" in c.definition
    zero = kpi(fr, k["ebitda_margin"], E, "actual", ["Feb 26"])
    assert zero.value is None and zero.reason == "division by zero"


def test_blank_point_in_the_average_counts_as_zero_and_stays_a_point():
    v = months({}, "fte", "actual", "Dec 25", [None, 12.0, 14.0, 16.0])     # opening read as #Missing
    avg = average_level(frame(values=v), E, "fte", "actual", ["Jan 26", "Feb 26", "Mar 26"])
    assert avg.value == (0 + 12 + 14 + 16) / 4 and len(avg.inputs) == 4


# --- review fixes -----------------------------------------------------------------------------------------------


def _kpis():
    return {x.key: x for x in catalog.load().kpis}


def test_average_opens_at_the_month_before_the_window():
    v = months({}, "fte", "actual", "Sep 25", [float(i) for i in range(13)])     # Sep 25 .. Sep 26
    ltm = periods.ltm("Sep 26")
    avg = average_level(frame(values=v), E, "fte", "actual", ltm)
    assert len(ltm) == 12 and len(avg.inputs) == 13 and avg.value == sum(range(13)) / 13
    assert "Sep 25 opening + 12 month ends" in avg.definition
    jan = average_level(frame(values=months({}, "fte", "actual", "Dec 25", [10.0, 20.0])), E, "fte", "actual",
                        ["Jan 26"])
    assert jan.value == 15.0 and len(jan.inputs) == 2


def test_ltm_membership():
    assert len(periods.ltm("Sep 26")) == 12
    assert periods.ltm("Dec 26") == periods.year(26)


def test_kpi_factors_by_window():
    k = _kpis()
    v = months({}, "fte", "actual", "Dec 25", [10.0] * 13)
    months(v, "fte", "actual", "Jan 26", [10.0] * 12)
    months(v, "opex", "actual", "Jan 26", [-100.0] * 12)
    months(v, "ebitda", "actual", "Jan 26", [10.0] * 12)
    months(v, "rev", "actual", "Jan 26", [100.0] * 12)
    fr = frame(values=v)
    assert kpi(fr, k["opex_per_fte"], E, "actual", ["Jan 26"]).value == 100.0 / 10.0 * 12
    assert kpi(fr, k["opex_per_fte"], E, "actual", periods.ytd("Dec 26")).value == 1200.0 / 10.0
    assert kpi(fr, k["opex_per_fte"], E, "actual", periods.ltm("Dec 26")).value == 1200.0 / 10.0


def test_kpi_carries_parts_unit_and_values_in_its_definition():
    fr = _fte_frame()
    c = kpi(fr, _kpis()["opex_per_fte"], E, "actual", ["Jan 26", "Feb 26", "Mar 26"])
    assert c.unit == "nok_per_fte" and [p.value for p in c.parts] == [-900.0, 13.0]
    assert "-900" in c.definition and "13.0" in c.definition and "4 levels" in c.definition


def test_negative_revenue_margin_is_computed_not_hidden():
    v = months({}, "ebitda", "actual", "Jan 26", [-50.0]) | months({}, "rev", "actual", "Jan 26", [-100.0])
    assert kpi(frame(values=v), _kpis()["ebitda_margin"], E, "actual", ["Jan 26"]).value == 0.5


def test_default_kpi_id_is_unique_per_entity_and_span():
    fr, k = _fte_frame(), _kpis()["ebitda_margin"]
    other = "un.SY.2BUS"
    ids = {kpi(fr, k, E, "actual", ["Jan 26", "Feb 26"]).id, kpi(fr, k, E, "actual", ["Feb 26", "Mar 26"]).id,
           kpi(fr, k, other, "actual", ["Jan 26", "Feb 26"]).id}
    assert len(ids) == 3


def test_empty_and_noncontiguous_months():
    fr, k = _fte_frame(), _kpis()
    assert average_level(fr, E, "fte", "actual", []).value is None
    assert amount(fr, E, "fte", "actual", []).value is None
    assert kpi(fr, k["opex_per_fte"], E, "actual", []).value is None
    assert kpi(fr, k["ebitda_margin"], E, "actual", []).value is None
    assert flow(fr, E, "opex", "actual", []).value == 0.0
    with pytest.raises(ValueError, match="contiguous"):
        average_level(fr, E, "fte", "actual", ["Jan 26", "Mar 26"])
    with pytest.raises(ValueError, match="contiguous"):
        amount(fr, E, "opex", "actual", ["Mar 26", "Feb 26"])


def test_non_finite_values_are_missing():
    v = months({}, "rev", "actual", "Jan 26", [math.inf, math.nan])
    f = flow(frame(values=v), E, "rev", "actual", ["Jan 26"])
    assert f.value is None and "non-finite" in f.reason
    assert level(frame(values=v), E, "rev", "actual", "Feb 26").value is None


def test_fact_unit_is_validated():
    with pytest.raises(ValueError, match="unit"):
        Fact("x", 1.0, "eur", "bad")


def test_frame_fallback_key_has_seven_parts_and_from_rows_refuses_overwrites():
    assert len(_fte_frame().get(E, "rev", "Jan 26", "actual")[1]) == 7
    lines = catalog.load().pnl["FinMax"]
    rows = {("FinMax", "Actual", "", "E", "ac.C30", "Sep 26", "Periodic"): 1.0,
            ("FinMax", "Plan", "", "E", "ac.C30", "Sep 26", "Periodic"): 2.0}
    with pytest.raises(ValueError, match="both map to"):
        from_rows("FinMax", lines, rows, {("Actual", ""): "actual", ("Plan", ""): "actual"})
    dup = (*lines, lines[0].__class__(**{**lines[0].__dict__, "key": "rev2"}))
    with pytest.raises(ValueError, match="account"):
        from_rows("FinMax", dup, {}, {("Actual", ""): "actual"})


def test_cost_per_fte_is_shown_positive_and_keeps_better_semantics():
    import dataclasses

    from smartview_mcp.analytics.variance import variance
    k = _kpis()
    assert (k["opex_per_fte"].sign, k["opex_per_fte"].better) == (-1, "lower")
    assert (k["personnel_per_fte"].sign, k["personnel_per_fte"].better) == (-1, "lower")
    assert (k["ebitda_margin"].sign, k["ebitda_margin"].better) == (1, "higher")
    c = kpi(_fte_frame(), k["opex_per_fte"], E, "actual", ["Jan 26", "Feb 26", "Mar 26"])
    assert c.value == 900.0 / 13.0 * 12 / 3 and "shown positive: cost per FTE" in c.definition
    assert [p.value for p in c.parts] == [-900.0, 13.0]                  # the parts keep the FinMax sign
    lower = dataclasses.replace(c, id="cur", value=100.0)
    base = dataclasses.replace(c, id="cmp", value=120.0)
    assert variance(k["opex_per_fte"], lower, base, "x").favourable is True    # lower cost per FTE is better


def test_ltm_per_fte_average_uses_13_points():
    v = months({}, "fte", "actual", "Sep 25", [10.0] * 13)
    months(v, "opex", "actual", "Oct 25", [-100.0] * 12)
    c = kpi(frame(values=v), _kpis()["opex_per_fte"], E, "actual", periods.ltm("Sep 26"))
    assert c.value == 1200.0 / 10.0 and "13 levels" in c.definition and len(c.parts[1].inputs) == 13


def test_kpi_with_an_unknown_unit_raises():
    k = dataclasses.replace(_kpis()["ebitda_margin"], unit="bogus")
    with pytest.raises(ValueError, match="bogus"):
        kpi(frame(), k, E, "actual", ["Jan 26"])


def test_total_keeps_its_parts():
    a, b = Fact("a", 1.0, "nok", ""), Fact("b", 2.0, "nok", "")
    assert total("t", [a, b], "nok", "").parts == (a, b)
    gone = Fact("g", None, "nok", "", (), "x")
    assert total("t", [a, gone], "nok", "").parts == (a, gone)
