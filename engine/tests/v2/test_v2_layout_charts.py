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
    assert c.type == "waterfall" and c.title == "EBITDA bridge"
    assert c.series_cols == ["C"] and len(c.rows) == 5 and c.totals == [1, 5]
    assert c.rows == list(range(c.rows[0], c.rows[0] + 5))
    assert c.header_row == c.rows[0] - 1
    assert c.anchor_col == "E"                                 # widest table is 1 value column: chart sits right of it


def test_trend_gets_a_line_anchor_without_total_row():
    p = run([{"type": "trend", "query": "fm", "metrics": ["Opex"]}], monthly_facts(), "Jul 26..Sep 26")
    (c,) = p.charts
    assert c.type == "line" and c.series_cols == ["C", "D"] and len(c.rows) == 3
    assert c.anchor_col == "F"


def test_no_chart_without_chart_blocks_and_anchor_right_of_table():
    p = run([{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}], facts())
    assert p.charts == []
    p = run([{"type": "trend", "query": "fm", "metrics": ["Opex"], "chart": False}], monthly_facts(),
            "Jul 26..Sep 26")
    assert p.charts == []


def run2(blocks, frame, periods=None, stock=()):
    d = spec_dict(blocks=blocks)
    d["queries"][0]["metrics"].append("FTE")
    if periods:
        d["queries"][0]["periods"] = periods
    s = AnalysisSpec.model_validate(d)
    return layout(1, s, compute(s, resolver(), CLOSED, frame), stock=set(stock), subtitle="x")


def test_a_chart_starts_below_the_previous_chart():
    import math
    blocks = [{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]},
              {"type": "trend", "query": "fm", "metrics": ["Opex"]}]
    p = run2(blocks, monthly_facts(), "Jul 26..Sep 26")
    first, second = p.charts
    assert first.type == "waterfall" and second.type == "line"
    assert first.anchor_row == first.header_row
    assert first.height_pt == max(180.0, 15.0 * (len(first.rows) + 3))
    bottom = first.anchor_row + math.ceil(first.height_pt / 15)
    assert second.anchor_row >= bottom + 1
    assert second.anchor_row >= second.header_row


def test_a_chart_far_below_keeps_its_own_header_row():
    blocks = [{"type": "trend", "query": "fm", "metrics": ["Opex"]},
              {"type": "table", "query": "fm", "title": "P&L"} ] + [
              {"type": "table", "query": "fm", "title": f"T{i}"} for i in range(6)] + [
              {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}]
    p = run2(blocks, monthly_facts(), "Jul 26..Sep 26")
    assert p.charts[-1].anchor_row >= p.charts[-1].header_row
def test_millions_axis_only_for_nok_amounts():
    flow = run2([{"type": "trend", "query": "fm", "metrics": ["Opex"]}], monthly_facts(), "Jul 26..Sep 26")
    assert flow.charts[0].millions is True
    fte = run2([{"type": "trend", "query": "fm", "metrics": ["FTE"]}], monthly_facts(), "Jul 26..Sep 26",
               stock={"FTE"})
    assert fte.charts[0].millions is False
    bridge = run2([{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}], facts(),
                  stock={"FTE"})
    assert bridge.charts[0].millions is True
