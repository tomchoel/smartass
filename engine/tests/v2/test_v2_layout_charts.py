# engine/tests/v2/test_v2_layout_charts.py
from v2util import CLOSED, facts, monthly_facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.spec import AnalysisSpec
from smartass.render.layout import layout


def run(blocks, frame, periods=None):
    d = spec_dict(blocks=blocks)
    if periods:
        d["queries"][0]["periods"] = periods
    s = AnalysisSpec.model_validate(d)
    return layout(1, s, compute(s, resolver(), CLOSED, frame), stock=set(), subtitle="x")


def test_bridge_gets_a_waterfall_anchor():
    p = run([{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}], facts())
    (c,) = p.charts
    assert c.kind == "waterfall" and c.title == "EBITDA bridge"
    a, b = c.cat_range.split(":")
    assert a[0] == b[0] == "C" and int(b[1:]) - int(a[1:]) == 4          # start, 2 steps, other, end
    assert p.cells[(int(a[1:]) - 1, 4)] == "Amount"                      # the bridge header right above
    assert [r for _, _, r in c.series][-1] == "total"
    assert (c.anchor_col, c.anchor_row) == ("B", dict(p.zones)["charts"])     # plan 5: in the Charts zone


def test_trend_gets_a_line_anchor_without_total_row():
    p = run([{"type": "trend", "query": "fm", "metrics": ["Opex"]}], monthly_facts(), "Jul 26..Sep 26")
    (c,) = p.charts
    assert c.kind == "line" and [rng[0] for _, rng, _ in c.series] == ["D", "E"]
    a, b = c.cat_range.split(":")
    assert int(b[1:]) - int(a[1:]) == 2                                  # three months, no Total
    assert c.anchor_col == "B" and c.anchor_row < int(a[1:])


def test_without_chart_blocks_the_implicit_chart_is_drawn():
    p = run([{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}], facts())
    assert [(c.kind, bool(c.series)) for c in p.charts] == [("column", True)]
    p = run([{"type": "trend", "query": "fm", "metrics": ["Opex"], "chart": False}], monthly_facts(),
            "Jul 26..Sep 26")
    assert [(c.kind, c.source) for c in p.charts] == [("line", "b0")]       # the implicit chart on the trend
    assert p.charts[0].series


def run2(blocks, frame, periods=None, stock=()):
    d = spec_dict(blocks=blocks)
    d["queries"][0]["metrics"].append("FTE")
    if periods:
        d["queries"][0]["periods"] = periods
    s = AnalysisSpec.model_validate(d)
    return layout(1, s, compute(s, resolver(), CLOSED, frame), stock=set(stock), subtitle="x")


def test_two_charts_share_a_zone_row_side_by_side():
    from smartass.render.layout import CHART_ROWS, CHART_WIDTH, width_pt
    blocks = [{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]},
              {"type": "trend", "query": "fm", "metrics": ["Opex"]}]
    p = run2(blocks, monthly_facts(), "Jul 26..Sep 26")
    first, second = p.charts
    assert first.kind == "waterfall" and second.kind == "line"
    assert first.anchor_row == second.anchor_row == dict(p.zones)["charts"]
    assert first.height_pt < CHART_ROWS * 15
    cols = [chr(c) for c in range(ord("B"), ord(second.anchor_col))]
    assert sum(width_pt(p.widths.get(c, 8.43)) for c in cols) >= CHART_WIDTH      # never over the first chart


def test_a_chart_far_below_keeps_its_own_header_row():
    blocks = [{"type": "trend", "query": "fm", "metrics": ["Opex"]},
              {"type": "table", "query": "fm", "title": "P&L"} ] + [
              {"type": "table", "query": "fm", "title": f"T{i}"} for i in range(6)] + [
              {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}]
    p = run2(blocks, monthly_facts(), "Jul 26..Sep 26")
    top_table = int(p.charts[0].cat_range.split(":")[0][1:])
    assert p.charts[-1].anchor_row < top_table                               # every chart above the tables
    assert p.charts[-1].anchor_row == dict(p.zones)["charts"]
def test_millions_axis_only_for_nok_amounts():
    from test_v5_layout import scaled
    flow = run2([{"type": "trend", "query": "fm", "metrics": ["Opex"]}], scaled(monthly_facts(), 1e6),
                "Jul 26..Sep 26")
    assert flow.charts[0].millions is True and flow.charts[0].whole is False
    small = run2([{"type": "trend", "query": "fm", "metrics": ["Opex"]}], monthly_facts(), "Jul 26..Sep 26")
    assert small.charts[0].millions is False and small.charts[0].whole is True       # NOK under 1 m: whole NOK
    fte = run2([{"type": "trend", "query": "fm", "metrics": ["FTE"]}], monthly_facts(), "Jul 26..Sep 26",
               stock={"FTE"})
    assert fte.charts[0].millions is False
    bridge = run2([{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}],
                  scaled(facts(), 1e6), stock={"FTE"})
    assert bridge.charts[0].millions is True
