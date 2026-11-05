# engine/tests/v2/test_v2_block_ranking.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.ranking import build_ranking
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def test_ranking_top_other_total_share_and_check():
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "ranking", "query": "fm", "metric": "Opex",
                                                       "top": 1}]))
    c = Ctx(s, resolver(), CLOSED, facts())
    out = build_ranking(s.blocks[0], c, "b0")
    assert out.headers == ["AC Sep 26", "PY Sep 25", "ΔPY", "ΔPY %", "Share %"]
    assert out.keys == ["Actual", "Last year", "Variance", "Variance %", "Share %"]
    assert [ln.key for ln in out.lines] == ["e0", "other", "total", "check"]
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b0", "e0", 0)) == -75.0
    assert cells.get(ck("b0", "other", 2)) == -1.0          # G910: -10 vs -9
    assert cells.get(ck("b0", "other", 4)) == pytest.approx(10 / 85)
    assert cells.get(ck("b0", "check", 0)) == 0.0


def test_ranking_without_compare():
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "ranking", "query": "fm", "metric": "Opex",
                                                       "compare": None}]))
    c = Ctx(s, resolver(), CLOSED, facts())
    assert build_ranking(s.blocks[0], c, "b0").headers == ["AC Sep 26", "Share %"]
