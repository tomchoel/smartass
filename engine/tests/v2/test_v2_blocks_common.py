# engine/tests/v2/test_v2_blocks_common.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.common import BlockOut, ck, entity_lines, evaluate, row_cells
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def ctx():
    return Ctx(AnalysisSpec.model_validate(spec_dict()), resolver(), CLOSED, facts())


def test_row_cells_value_ly_var_pct():
    c = ctx()
    q = c.query("fm")
    out = BlockOut("b0", "table", "t", ["Actual", "Last year", "Variance", "Variance %"])
    from smartass.model.blocks.common import Line
    out.lines.append(Line("m0", "Opex", row_cells(c, q, "Opex", ["value", "ly", "var", "var_pct"], "b0", "m0")))
    cells = evaluate([out], c.frame)
    assert [cells.get(ck("b0", "m0", j)) for j in range(3)] == [-85.0, -77.0, -8.0]
    assert cells.get(ck("b0", "m0", 3)) == pytest.approx(-8.0 / 77.0)


def test_entity_lines_top_other_total_and_check():
    c = ctx()
    q = c.query("fm")
    cols = ["value", "share"]
    out = BlockOut("b1", "table", "t", ["Actual", "Share %"], lines=entity_lines(c, q, "Opex", cols, "b1",
                                                                               top=1, sort="abs"))
    assert [ln.key for ln in out.lines] == ["e0", "other", "total", "check"]
    assert out.lines[0].label == "un.SY.4BUS_B"
    assert out.lines[2].label == "2BUS Business Norway" and out.lines[2].kind == "grand"
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b1", "e0", 0)) == -75.0
    assert cells.get(ck("b1", "other", 0)) == -10.0
    assert cells.get(ck("b1", "total", 0)) == -85.0
    assert cells.get(ck("b1", "e0", 1)) == pytest.approx(75 / 85)
    assert cells.get(ck("b1", "check", 0)) == 0.0


def test_entity_lines_without_children_has_no_check():
    c = ctx()
    q = c.query("fm").model_copy(update={"entity": "un.SY.B1"})
    lines = entity_lines(c, q, "Opex", ["value"], "b2")
    assert [ln.key for ln in lines] == ["total"]


def weighted_ctx(weights, parent_value, block=None):
    from smartass.model.frame import Fact, Frame
    from smartass.model.resolver import MetricDef, StaticResolver
    d = spec_dict(blocks=[block or {"type": "matrix", "query": "fm", "metric": "Revenue"}])
    d["queries"][0].update(entity="P", periods="Sep 26")
    r = StaticResolver([MetricDef("Revenue", {"FinMax": ("ac.C30",)})], tree={"P": ("A", "B")}, weights=weights)
    fr = Frame([Fact("FinMax", "fm", "ac.C30", e, "Sep 26", "Actual", "", v)
                for e, v in (("A", 100.0), ("B", 30.0), ("P", parent_value))])
    return Ctx(AnalysisSpec.model_validate(d), r, CLOSED, fr)


def test_subtracted_child_reconciles_with_its_parent():
    from smartass.model.blocks.trend import build_matrix
    c = weighted_ctx({"B": -1.0}, 70.0)                       # P = A - B
    q = c.query("fm")
    lines = entity_lines(c, q, "Revenue", ["value"], "b0")
    cells = evaluate([Line_out("b0", lines)], c.frame)
    assert cells.get(ck("b0", "check", 0)) == 0.0
    out = build_matrix(c.spec.blocks[0], c, "b1")
    assert evaluate([out], c.frame).get(ck("b1", "check", 0)) == 0.0


def test_ignored_child_is_left_out_of_the_check():
    from smartass.model.blocks.trend import build_matrix
    c = weighted_ctx({"B": 0.0}, 100.0)                       # P = A, B is '~'
    lines = entity_lines(c, c.query("fm"), "Revenue", ["value"], "b0")
    assert evaluate([Line_out("b0", lines)], c.frame).get(ck("b0", "check", 0)) == 0.0
    out = build_matrix(c.spec.blocks[0], c, "b1")
    assert evaluate([out], c.frame).get(ck("b1", "check", 0)) == 0.0


def test_children_sum_with_only_subtracted_children():
    from smartass.model.blocks.common import children_sum
    c = weighted_ctx({"A": -1.0, "B": -1.0}, -130.0)
    assert c.value(children_sum(c, c.query("fm"), "Revenue", ["A", "B"])) == -130.0


def Line_out(key, lines):  # noqa: N802
    return BlockOut(key, "table", "t", ["Actual"], lines=lines)
