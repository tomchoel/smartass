"""analytics.measures and analytics.kpis: flows sum, stocks take levels, the n+1 average, annualising."""
from analytics_util import E, frame, months

from smartview_mcp import catalog
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
    assert c.value == -900.0 / 13.0 * 12 / 3 and c.unit == "nok_per_fte" and "x 12/3" in c.definition
    zero = kpi(fr, k["ebitda_margin"], E, "actual", ["Feb 26"])
    assert zero.value is None and zero.reason == "division by zero"


def test_blank_point_in_the_average_counts_as_zero_and_stays_a_point():
    v = months({}, "fte", "actual", "Dec 25", [None, 12.0, 14.0, 16.0])     # opening read as #Missing
    avg = average_level(frame(values=v), E, "fte", "actual", ["Jan 26", "Feb 26", "Mar 26"])
    assert avg.value == (0 + 12 + 14 + 16) / 4 and len(avg.inputs) == 4
