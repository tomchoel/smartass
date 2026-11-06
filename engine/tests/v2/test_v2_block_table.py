# engine/tests/v2/test_v2_block_table.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.table import build_kpi, build_table
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def run(blocks):
    s = AnalysisSpec.model_validate(spec_dict(blocks=blocks))
    c = Ctx(s, resolver(), CLOSED, facts())
    return c, s.blocks


def test_table_metric_rows_default_to_query_metrics():
    c, (b,) = run([{"type": "table", "query": "fm", "title": "P&L"}])
    out = build_table(b, c, "b0")
    assert out.headers == ["AC Sep 26", "PY Sep 25", "ΔPY", "ΔPY %", "Material"]
    assert out.keys == ["Actual", "Last year", "Variance", "Variance %", ""]
    assert [ln.label for ln in out.lines] == ["Revenue", "Opex", "EBITDA"]
    cells = evaluate([out], c.frame)
    assert [cells.get(ck("b0", "m2", j)) for j in range(3)] == [65.0, 63.0, 2.0]


def test_table_entity_rows():
    c, (b,) = run([{"type": "table", "query": "fm", "rows": "entities", "metric": "Opex",
                    "columns": ["value", "ly", "var", "share"], "sort": "abs"}])
    out = build_table(b, c, "b0")
    assert [ln.key for ln in out.lines] == ["e0", "e1", "total", "check"]
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b0", "e0", 2)) == -7.0
    assert cells.get(ck("b0", "e1", 3)) == pytest.approx(10 / 85)


def test_kpi_tiles():
    c, (b,) = run([{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}])
    out = build_kpi(b, c, "b0")
    cells = evaluate([out], c.frame)
    assert out.type == "kpi_tiles"
    assert [cells.get(ck("b0", "k0", j)) for j in range(3)] == [65.0, 63.0, 2.0]


def plan_ctx(columns):
    from smartass.model.frame import Fact, Frame
    d = spec_dict(blocks=[{"type": "table", "query": "fp", "columns": columns, "metrics": ["EBITDA"]}])
    d["queries"] = [{"id": "fp", "cube": "FinPlan", "metrics": ["EBITDA"], "entity": "un.SY.2BUS",
                     "periods": "closed", "compare": ["budget", "forecast"]}]
    s = AnalysisSpec.model_validate(d)
    fr = Frame([Fact("FinPlan", "fp", "EBITDA", "un.SY.2BUS", "Sep 26", sc, ver, v)
                for sc, ver, v in (("Actual", "Working", 70.0), ("Target", "Final", 80.0),
                                   ("Forecast Q326", "Final", 75.0))])
    return Ctx(s, resolver(), CLOSED, fr), s.blocks[0]


def test_variance_compares_against_the_nearest_base_to_its_left():
    c, b = plan_ctx(["value", "budget", "var", "forecast", "var", "var_pct"])
    out = build_table(b, c, "b0")
    assert out.headers == ["AC Sep 26", "BU Sep 26", "ΔBU", "FC Sep 26 (Q3 26)", "ΔFC", "ΔFC %", "Material"]
    assert out.keys == ["Actual", "Budget", "Variance vs Budget", "Forecast", "Variance vs Forecast",
                        "Variance % vs Forecast", ""]
    cells = evaluate([out], c.frame)
    assert [cells.get(ck("b0", "m0", j)) for j in range(5)] == [70.0, 80.0, -10.0, 75.0, -5.0]
    assert cells.get(ck("b0", "m0", 5)) == pytest.approx(-5.0 / 75.0)


def test_single_base_keeps_plain_variance_headers():
    c, b = plan_ctx(["value", "budget", "var", "var_pct"])
    out = build_table(b, c, "b0")
    assert out.headers == ["AC Sep 26", "BU Sep 26", "ΔBU", "ΔBU %", "Material"]
    assert out.keys == ["Actual", "Budget", "Variance", "Variance %", ""]


def test_blocks_expose_column_codes_and_line_entities():
    c, (b,) = run([{"type": "table", "query": "fm", "rows": "entities", "metric": "Opex",
                    "columns": ["value", "ly", "var", "var_pct", "share"], "top": 1}])
    out = build_table(b, c, "b0")
    assert out.cols == ["value", "ly", "var", "var_pct", "share", "material"] and len(out.cols) == len(out.headers)
    assert [(ln.key, ln.entity) for ln in out.lines] == [("e0", "un.SY.4BUS_B"), ("other", None),
                                                         ("total", "un.SY.2BUS"), ("check", None)]
    c, (b,) = run([{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}])
    out = build_kpi(b, c, "b0")
    assert out.cols == ["value", "ly", "var", "var_pct"] and out.lines[0].entity == "un.SY.2BUS"


def test_material_reads_every_variance_column_by_code():
    from smartass.model.findings import material
    c, b = plan_ctx(["value", "budget", "var", "forecast", "var", "var_pct"])
    c.resolver._m['ebitda'].materiality = (7.0, 0.5)         # EBITDA: |-10| vs budget is material, |-5| is not
    out = build_table(b, c, "b0")
    hits = material([out], evaluate([out], c.frame), c.resolver)
    assert [(h["key"], h["variance"]) for h in hits] == [("b0.m0", -10.0)]
