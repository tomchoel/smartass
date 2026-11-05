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
    assert c.series_cols == ["D"] and len(c.rows) == 5 and c.totals == [1, 5]
    assert c.rows == list(range(c.rows[0], c.rows[0] + 5))
    assert c.header_row == c.rows[0] - 1
    assert (c.anchor_col, c.anchor_row) == ("B", dict(p.zones)["charts"])     # plan 5: in the Charts zone


def test_trend_gets_a_line_anchor_without_total_row():
    p = run([{"type": "trend", "query": "fm", "metrics": ["Opex"]}], monthly_facts(), "Jul 26..Sep 26")
    (c,) = p.charts
    assert c.type == "line" and c.series_cols == ["D", "E"] and len(c.rows) == 3
    assert c.anchor_col == "B" and c.anchor_row < c.header_row


def test_without_chart_blocks_the_implicit_chart_is_reserved_but_not_drawn_yet():
    p = run([{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}], facts())
    assert [(c.type, c.rows) for c in p.charts] == [("column", [])]          # plan 5 Task 7 draws it
    p = run([{"type": "trend", "query": "fm", "metrics": ["Opex"], "chart": False}], monthly_facts(),
            "Jul 26..Sep 26")
    assert [(c.type, c.source) for c in p.charts] == [("line", "b0")]       # the implicit chart on the trend
    assert p.charts[0].rows


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
    assert first.type == "waterfall" and second.type == "line"
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
    assert p.charts[-1].anchor_row < p.charts[0].header_row                  # every chart above the tables
    assert p.charts[-1].anchor_row == dict(p.zones)["charts"]
def test_millions_axis_only_for_nok_amounts():
    flow = run2([{"type": "trend", "query": "fm", "metrics": ["Opex"]}], monthly_facts(), "Jul 26..Sep 26")
    assert flow.charts[0].millions is True
    fte = run2([{"type": "trend", "query": "fm", "metrics": ["FTE"]}], monthly_facts(), "Jul 26..Sep 26",
               stock={"FTE"})
    assert fte.charts[0].millions is False
    bridge = run2([{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}], facts(),
                  stock={"FTE"})
    assert bridge.charts[0].millions is True
