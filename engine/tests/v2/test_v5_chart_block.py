# engine/tests/v2/test_v5_chart_block.py
"""Plan 5 Task 4: chart blocks (spec, guard on new questions only, implicit charts elsewhere, own chart data) and
the Material marker as an OR over every variance of the last group. Every figure is invented."""
import asyncio
import json

import pytest
from pydantic import ValidationError
from test_v2_pipeline import FakeProvider
from test_v2_redraw import wire
from v2util import CLOSED, facts, resolver, spec_dict

from smartass import pipeline, server
from smartass.model.analyze import compute
from smartass.model.blocks.chart import DEFAULT_SERIES, IMPLICIT_KEY, default_series
from smartass.model.blocks.common import BlockOut, Line, add_material, ck, evaluate
from smartass.model.errors import GuardError
from smartass.model.expr import Const
from smartass.model.findings import material
from smartass.model.frame import Frame
from smartass.model.requests import requests
from smartass.model.spec import AnalysisSpec, ChartBlock
from smartass.model.validate import guard_chart
from smartass.render.layout import layout
from smartass.testset.runner import load_cases

TABLE = {"type": "table", "query": "fm", "title": "P&L", "metrics": ["Revenue", "Opex"], "chain": False}


def spec(*blocks, depth="standard"):
    d = spec_dict(blocks=list(blocks))
    d["plan"]["depth"] = depth
    return AnalysisSpec.model_validate(d)


def chart(**kw):
    return {"type": "chart", "kind": "column", "title": "Opex rose vs last year", "source": {"block": 0}, **kw}


def outs_of(s):
    return compute(s, resolver(), CLOSED, facts()).outs


# --- the guard: new standard/deep questions only --------------------------------------------------------------

def test_a_standard_spec_without_a_chart_is_refused():
    with pytest.raises(GuardError) as e:
        guard_chart(spec(TABLE))
    assert e.value.reason == "the answer has no chart"
    assert "type:'chart'" in e.value.fix and "waterfall" in e.value.fix


def test_quick_passes_the_guard_and_gets_an_implicit_column_chart():
    s = spec(TABLE, depth="quick")
    guard_chart(s)
    outs = outs_of(s)
    c = outs[-1]
    assert c.key == IMPLICIT_KEY and c.type == "chart" and c.chart == "column" and c.chart_of == "b0"
    assert c.lines == [] and c.headers == []
    assert not any(isinstance(b, ChartBlock) for b in s.blocks)      # the spec stays as given


def test_a_chart_block_passes_and_gets_no_implicit_chart():
    s = spec(TABLE, chart())
    guard_chart(s)
    outs = outs_of(s)
    assert [o.key for o in outs if o.type == "chart"] == ["b1"]
    assert outs[-1].chart_of == "b0" and outs[-1].title == "Opex rose vs last year"


def test_a_bridge_with_chart_true_passes_and_chart_false_does_not():
    bridge = {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}
    d = spec_dict(blocks=[bridge])
    guard_chart(AnalysisSpec.model_validate(d))                      # chart defaults to true (as before)
    d["blocks"] = [dict(bridge, chart=True)]
    guard_chart(AnalysisSpec.model_validate(d))
    d["blocks"] = [dict(bridge, chart=False)]
    s = AnalysisSpec.model_validate(d)
    with pytest.raises(GuardError):
        guard_chart(s)
    outs = outs_of(s)
    assert outs[0].chart is None                                     # no chart drawn on the bridge itself
    assert outs[-1].type == "chart" and outs[-1].chart == "waterfall" and outs[-1].chart_of == "b0"


def test_a_trend_with_chart_true_passes():
    guard_chart(spec({"type": "trend", "query": "fm", "metrics": ["Opex"]}))
    with pytest.raises(GuardError):
        guard_chart(spec({"type": "trend", "query": "fm", "metrics": ["Opex"], "chart": False}))


def test_implicit_chart_prefers_bridge_then_trend_then_table():
    bridge = {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"], "chart": False}
    trend = {"type": "trend", "query": "fm", "metrics": ["Opex"], "chart": False}
    c = outs_of(spec(TABLE, trend, bridge, depth="quick"))[-1]
    assert (c.chart, c.chart_of) == ("waterfall", "b2")
    c = outs_of(spec(TABLE, trend, depth="quick"))[-1]
    assert (c.chart, c.chart_of) == ("line", "b1")
    kpi = {"type": "kpi_tiles", "query": "fm", "metrics": ["Opex"]}
    c = outs_of(spec(kpi, TABLE, depth="quick"))[-1]
    assert (c.chart, c.chart_of) == ("column", "b1")
    c = outs_of(spec(kpi, depth="quick"))[-1]                        # no table: the first block with lines
    assert (c.type, c.chart, c.chart_of) == ("chart", "column", "b0")


def test_implicit_chart_keeps_the_key_values(monkeypatch):
    from smartass.model import analyze
    s = spec(TABLE, depth="quick")
    res = compute(s, resolver(), CLOSED, facts())
    monkeypatch.setattr(analyze, "implicit_chart", lambda outs: None)
    bare = compute(s, resolver(), CLOSED, facts())
    assert res.key_values == bare.key_values and res.material == bare.material and res.params == bare.params
    assert len(res.outs) == len(bare.outs) + 1


def test_a_chart_titled_like_a_table_does_not_rename_the_table():
    s = spec(chart(source={"block": 1}, title="P&L"), TABLE)
    res = compute(s, resolver(), CLOSED, facts())
    assert any(k.startswith("P&L · ") for k in res.key_values)
    assert not any("P&L (2)" in k for k in res.key_values)


# --- the spec ------------------------------------------------------------------------------------------------

def test_source_block_out_of_range_is_refused():
    with pytest.raises(ValidationError, match="source.block 5"):
        spec(TABLE, chart(source={"block": 5}))


def test_source_block_must_not_be_a_chart_or_commentary():
    with pytest.raises(ValidationError, match="commentary"):
        spec({"type": "commentary"}, TABLE, chart(source={"block": 0}))
    with pytest.raises(ValidationError, match="chart"):
        spec(TABLE, chart(), chart(source={"block": 1}))


def test_waterfall_needs_a_bridge_or_own_data_with_amount():
    with pytest.raises(ValidationError, match="waterfall"):
        spec(TABLE, chart(kind="waterfall"))
    bridge = {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}
    spec(bridge, chart(kind="waterfall"))
    own = {"query": "fm", "metrics": ["Opex"]}
    with pytest.raises(ValidationError, match="waterfall"):
        spec(TABLE, chart(kind="waterfall", source=own))
    spec(TABLE, chart(kind="waterfall", source=own, series=["amount"]))


def test_source_needs_exactly_one_of_block_and_query():
    with pytest.raises(ValidationError, match="exactly one"):
        spec(TABLE, chart(source={}))
    with pytest.raises(ValidationError, match="exactly one"):
        spec(TABLE, chart(source={"block": 0, "query": "fm", "metrics": ["Opex"]}))


def test_empty_title_is_refused():
    for t in ("", "   "):
        with pytest.raises(ValidationError, match="title"):
            spec(TABLE, chart(title=t))
    with pytest.raises(ValidationError):
        spec(TABLE, {k: v for k, v in chart().items() if k != "title"})


def test_bar_defaults_to_sort_desc():
    s = spec(TABLE, chart(kind="bar"))
    assert s.blocks[1].sort == "desc"
    assert spec(TABLE, chart(kind="bar", sort="asc")).blocks[1].sort == "asc"
    assert spec(TABLE, chart(kind="column")).blocks[1].sort == "none"


def test_plan_title_defaults_to_empty():
    s = spec(TABLE, chart())
    assert s.plan.title == ""
    d = spec_dict(blocks=[TABLE, chart()])
    d["plan"]["title"] = "Opex up 4 % on consultants"
    assert AnalysisSpec.model_validate(d).plan.title == "Opex up 4 % on consultants"


def test_own_data_metrics_must_be_in_the_query():
    with pytest.raises(ValidationError, match="not in query"):
        spec(TABLE, chart(source={"query": "fm", "metrics": ["FTE"]}))
    with pytest.raises(ValidationError, match="unknown query"):
        spec(TABLE, chart(source={"query": "zz", "metrics": ["Opex"]}))
    with pytest.raises(ValidationError, match="compare"):
        spec(TABLE, chart(source={"query": "fm", "metrics": ["Opex"], "columns": ["value", "budget"]}))


def test_default_series_by_kind():
    assert DEFAULT_SERIES["waterfall"] == ["amount"] and DEFAULT_SERIES["variance"] == ["var"]
    assert default_series("line", ["value", "ly", "var"]) == ["value", "ly"]
    assert default_series("column", ["value", "ly", "var"]) == ["value", "ly"]
    assert default_series("column", ["value", "budget", "var"]) == ["value", "budget"]
    assert default_series("bar", ["value", "ly"]) == ["value"]
    assert default_series("combo", ["value", "ly", "var_pct"]) == ["value", "var_pct"]


# --- own chart data --------------------------------------------------------------------------------------------

def test_own_data_chart_equals_a_table_with_the_same_query_and_metrics():
    own = {"query": "fm", "metrics": ["Revenue", "Opex"], "columns": ["value", "ly"]}
    s = spec(TABLE, chart(source=own))
    res = compute(s, resolver(), CLOSED, facts())
    o = next(o for o in res.outs if o.type == "chart_data")
    assert o.key == "b1" and o.title == "Chart data: Opex rose vs last year" and o.chart == "column"
    assert "material" not in o.cols and o.cols == ["value", "ly"]
    t = spec({"type": "table", "query": "fm", "metrics": ["Revenue", "Opex"], "columns": ["value", "ly"],
              "chain": False, "material": False}, depth="quick")
    tres = compute(t, resolver(), CLOSED, facts())
    tab = tres.outs[0]
    assert [ln.label for ln in o.lines] == [ln.label for ln in tab.lines]
    for ln, tl in zip(o.lines, tab.lines):
        for j in range(len(o.cols)):
            assert res.cells.get(ck("b1", ln.key, j)) == tres.cells.get(ck("b0", tl.key, j))
    assert not any(x.type == "chart" for x in res.outs)              # an own-data chart is the chart


def test_own_data_entity_rows_fetch_the_children():
    own = {"query": "fm", "rows": "entities", "metric": "Opex", "columns": ["value", "ly"]}
    s = spec(TABLE, chart(kind="bar", source=own))
    reqs = requests(s, resolver(), {"FinMax": CLOSED, "FinLite": CLOSED, "FinPlan": CLOSED})
    assert {"un.SY.4BUS_B", "un.SY.G910"} <= set(reqs[0].entities)
    prov = FakeProvider(facts())
    pre = pipeline.prepare(s.model_dump(), resolver(), prov)
    assert any("un.SY.G910" in r.entities for r in prov.asked)
    o = next(o for o in pre.result.outs if o.type == "chart_data")
    g910 = next(ln for ln in o.lines if ln.entity == "un.SY.G910")
    assert pre.result.cells.get(ck(o.key, g910.key, 0)) == -10.0         # Sep 26: C50S -8 + C60 -2


# --- layout: chart outs do not break the sheet -----------------------------------------------------------------

def test_layout_skips_chart_outs_and_lays_out_chart_data_as_a_table():
    own = {"query": "fm", "metrics": ["Opex"], "columns": ["value", "ly"]}
    s = spec(TABLE, chart(), chart(title="Own data", source=own))
    res = compute(s, resolver(), CLOSED, facts())
    p = layout(3, s, res, stock=set(), subtitle="NOK")
    texts = [v for v in p.cells.values() if isinstance(v, str)]
    assert "Opex rose vs last year" not in texts
    assert "Chart data: Own data" in texts
    o = next(o for o in res.outs if o.type == "chart_data")
    assert all(ck(o.key, ln.key, j) in p.addr for ln in o.lines for j, e in enumerate(ln.cells) if e is not None)
    assert p.charts == []                                            # chart blocks are drawn from Task 7 on


def test_layout_still_charts_trend_and_bridge():
    bridge = {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}
    s = spec(TABLE, bridge)
    res = compute(s, resolver(), CLOSED, facts())
    p = layout(3, s, res, stock=set(), subtitle="NOK")
    assert [c.type for c in p.charts] == ["waterfall"]


# --- paths: guard on new questions only, implicit chart everywhere else -----------------------------------------

def _server(monkeypatch):
    seen = {}

    async def user(fn, what):
        return fn(None)

    def prep(raw, provider, lines=None, **kw):
        seen["raw"] = raw
        return "PRE"
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "_provider", lambda: "provider")
    monkeypatch.setattr(server, "run_prepare", prep)
    monkeypatch.setattr(server, "excel_job", lambda h, pre, target="session": {"q": 1})
    return seen


def test_new_question_without_chart_is_refused_before_any_work(monkeypatch):
    seen = _server(monkeypatch)
    with pytest.raises(GuardError, match="no chart"):
        asyncio.run(server._new_question(spec(TABLE), "session"))
    assert "raw" not in seen
    asyncio.run(server._new_question(spec(TABLE, chart()), "session"))
    assert seen["raw"]["blocks"][1]["type"] == "chart"
    asyncio.run(server._new_question(spec(TABLE, depth="quick"), "session"))


def test_fin_drill_is_quick_and_needs_no_chart(monkeypatch):
    from types import SimpleNamespace
    seen = _server(monkeypatch)

    async def drill(cell):
        return SimpleNamespace(rows=[{"AMOUNT": -1.0}], columns=["AMOUNT"], cell_value=-1.0,
                               reconciliation={"gap": 0.0}, warnings=[])
    monkeypatch.setattr(server, "_drill", drill)
    asyncio.run(server.fin_drill({"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC",
                                  "All Time Periods": "Sep 26"}))
    assert seen["raw"]["blocks"][0]["type"] == "lines"


def test_a_stored_pre_v5_spec_prepares_and_redraws_with_an_implicit_chart(monkeypatch):
    stored = spec_dict(blocks=[TABLE])                              # standard, no chart: stored before plan 5
    pre =pipeline.prepare(json.loads(json.dumps(stored)), resolver(), FakeProvider(facts()))
    assert pre.result.outs[-1].type == "chart" and pre.result.outs[-1].chart_of == "b0"
    assert [b["type"] for b in pre.raw["blocks"]] == ["table"]       # the stored spec stays as given
    log = []
    wire(monkeypatch, log)
    out = pipeline.redraw_job(None, pre, q=2, owned=None, target="session")
    assert out["q"] == 2 and "write_charts" in log
    # a refresh re-prepares the stored spec without complete and without the guard
    again = pipeline.prepare(pre.raw, resolver(), FakeProvider(facts()))
    assert again.result.outs[-1].type == "chart"


def test_test_set_cases_stay_without_own_data_charts():
    for case in load_cases():
        s = AnalysisSpec.model_validate(case["spec"])
        assert not any(isinstance(b, ChartBlock) and b.source.query for b in s.blocks), case["id"]


# --- the Material marker over every base of the last group ----------------------------------------------------

class _R:
    def materiality(self, metric):
        return 100.0, 0.05


def _mat_out(value, ly, budget):
    cols = ["value", "ly", "var", "var_pct", "budget", "var", "var_pct"]
    o = BlockOut("b0", "table", "T", [""] * len(cols), cols=list(cols), query="fm")
    from smartass.model.blocks.common import derived
    cells = [Const(value), Const(ly), None, None, Const(budget), None, None]
    for j, c in enumerate(cols):
        if c in ("var", "var_pct"):
            cells[j] = derived(c, cols, "b0", "m0", j=j)
    o.lines.append(Line("m0", "Opex", cells, metric="Opex", entity="e"))
    add_material(o, _R())
    return o


def test_material_is_one_when_only_the_budget_variance_is_material():
    o = _mat_out(1000.0, 990.0, 700.0)                              # ΔPY 10 (1 %): no; ΔBU 300: yes
    cells = evaluate([o], Frame([]))
    hits = material([o], cells, _R())
    assert [h["base"] for h in hits] == ["budget"]
    assert cells.get(ck("b0", "m0", len(o.cols) - 1)) == 1.0
    f = o.lines[0].cells[-1].to_excel("T", {ck("b0", "m0", 2): "E5", ck("b0", "m0", 3): "F5",
                                            ck("b0", "m0", 5): "H5", ck("b0", "m0", 6): "I5",
                                            "mat_abs.Opex": "$D$40", "mat_pct.Opex": "$D$41",
                                            "mat_small.Opex": "$D$42"}.__getitem__)
    assert f == ("IF(OR(OR(ABS(E5)>=$D$40,AND(ABS(N(F5))>=$D$41,ABS(E5)>=$D$42)),"
                 "OR(ABS(H5)>=$D$40,AND(ABS(N(I5))>=$D$41,ABS(H5)>=$D$42))),1,0)")


def test_material_is_zero_when_no_base_is_material_and_one_for_ly():
    o = _mat_out(1000.0, 990.0, 995.0)
    cells = evaluate([o], Frame([]))
    assert material([o], cells, _R()) == [] and cells.get(ck("b0", "m0", len(o.cols) - 1)) == 0.0
    o = _mat_out(1000.0, 800.0, 995.0)
    cells = evaluate([o], Frame([]))
    assert cells.get(ck("b0", "m0", len(o.cols) - 1)) == 1.0
