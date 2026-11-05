# engine/tests/v2/test_v2_block_trend.py
from v2util import CLOSED, monthly_facts, resolver, spec_dict

from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.trend import build_matrix, build_trend
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def run(block):
    d = spec_dict(blocks=[block])
    d["queries"][0]["periods"] = "Jul 26..Sep 26"
    s = AnalysisSpec.model_validate(d)
    return Ctx(s, resolver(), CLOSED, monthly_facts()), s.blocks[0]


def test_trend_months_series_and_total():
    c, b = run({"type": "trend", "query": "fm", "metrics": ["Opex"]})
    out = build_trend(b, c, "b0")
    assert out.headers == ["Opex AC", "Opex PY"] and out.chart == "line"       # the sheet: short + scenario
    assert out.keys == ["Opex Actual", "Opex Last year"]                        # key_values keys unchanged
    assert out.period == "Jul–Sep 26"
    assert [ln.label for ln in out.lines] == ["Jul 26", "Aug 26", "Sep 26", "Total"]
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b0", "p0", 0)) == -24.0          # 3 entities x (-7 - 1)
    assert cells.get(ck("b0", "total", 0)) == -24.0 - 27.0 - 30.0
    assert cells.get(ck("b0", "p0", 1)) == -24.0          # Jul 25 has the same invented values


def test_trend_cumulative_has_no_total():
    c, b = run({"type": "trend", "query": "fm", "metrics": ["Opex"], "series": ["value"], "cumulative": True})
    out = build_trend(b, c, "b0")
    cells = evaluate([out], c.frame)
    assert [ln.key for ln in out.lines] == ["p0", "p1", "p2"]
    assert cells.get(ck("b0", "p2", 0)) == -81.0


def test_matrix_children_by_month_with_check():
    c, b = run({"type": "matrix", "query": "fm", "metric": "Opex"})
    out = build_matrix(b, c, "b0")
    assert out.headers == ["Jul 26", "Aug 26", "Sep 26"]
    assert [ln.key for ln in out.lines] == ["e0", "e1", "total", "check"]
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b0", "e0", 2)) == -20.0          # 4BUS_B Sep: 2 x (-9 - 1)
    assert all(cells.get(ck("b0", "check", j)) == 0.0 for j in range(3))


def fte_ctx(cumulative=False):
    from smartass.model.frame import Fact, Frame
    d = spec_dict(blocks=[{"type": "trend", "query": "fm", "metrics": ["FTE"], "series": ["value"],
                           "cumulative": cumulative}])
    d["queries"][0].update(metrics=["FTE"], periods="Jul 26..Sep 26")
    s = AnalysisSpec.model_validate(d)
    fr = Frame([Fact("FinMax", "fm", "ac.STA20Y", "un.SY.2BUS", f"{m} 26", "Actual", "", v, "Accumulated")
                for m, v in (("Jul", 100.0), ("Aug", 101.0), ("Sep", 102.0))])
    return Ctx(s, resolver(), CLOSED, fr), s.blocks[0]


def test_stock_metric_has_no_total():
    c, b = fte_ctx()
    out = build_trend(b, c, "b0")
    assert [ln.key for ln in out.lines] == ["p0", "p1", "p2"]        # FTE never adds up over months: no Total


def test_stock_metric_does_not_accumulate():
    c, b = fte_ctx(cumulative=True)
    cells = evaluate([build_trend(b, c, "b0")], c.frame)
    assert [cells.get(ck("b0", f"p{i}", 0)) for i in range(3)] == [100.0, 101.0, 102.0]


def test_trend_runs_over_value_months_and_plan_months():
    from smartass.model.frame import Fact, Frame
    d = spec_dict(blocks=[{"type": "trend", "query": "fp", "metrics": ["EBITDA"], "series": ["value", "forecast"]}])
    d["queries"] = [{"id": "fp", "cube": "FinPlan", "metrics": ["EBITDA"], "entity": "un.SY.2BUS",
                     "periods": "Jan 26..Sep 26", "plan_periods": "Oct 26..Dec 26", "compare": ["forecast"]}]
    s = AnalysisSpec.model_validate(d)
    mons = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    fr = Frame([Fact("FinPlan", "fp", "EBITDA", "un.SY.2BUS", f"{m} 26", sc, "Final" if sc != "Actual" else "Working",
                     float(i + 1))
                for i, m in enumerate(mons) for sc in (("Actual",) if i < 9 else ("Forecast Q326",))])
    c = Ctx(s, resolver(), CLOSED, fr)
    out = build_trend(s.blocks[0], c, "b0")
    cells = evaluate([out], fr)
    assert [ln.label for ln in out.lines][:1] == ["Jan 26"] and len(out.lines) == 13      # 12 months + Total
    assert [out.lines[i].cells[0] is None for i in range(12)] == [False] * 9 + [True] * 3
    assert [out.lines[i].cells[1] is None for i in range(12)] == [True] * 9 + [False] * 3
    assert cells.get(ck("b0", "total", 0)) == 45.0 and cells.get(ck("b0", "total", 1)) == 10.0 + 11.0 + 12.0
    assert cells.get(ck("b0", "p9", 1)) == 10.0 and cells.get(ck("b0", "p9", 0)) is None
