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
    assert out.headers == ["Actual", "Last year", "Variance", "Variance %"]
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
