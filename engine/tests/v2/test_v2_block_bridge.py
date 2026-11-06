# engine/tests/v2/test_v2_block_bridge.py
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.bridge import build_bridge
from smartass.model.blocks.common import ck, evaluate
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def test_bridge_from_ly_to_actual_with_other():
    s = AnalysisSpec.model_validate(spec_dict(blocks=[
        {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}]))
    c = Ctx(s, resolver(), CLOSED, facts())
    out = build_bridge(s.blocks[0], c, "b0")
    assert out.chart == "waterfall" and out.headers == ["Amount"]
    assert [ln.key for ln in out.lines] == ["start", "s0", "s1", "other", "end"]
    cells = evaluate([out], c.frame)
    got = {ln.key: cells.get(ck("b0", ln.key, 0)) for ln in out.lines}
    assert got == {"start": 63.0, "s0": 10.0, "s1": -8.0, "other": 0.0, "end": 65.0}
    addr = {ck("b0", k, 0): f"D{i + 4}" for i, k in enumerate(got)}.__getitem__
    assert out.lines[3].cells[0].to_excel("Data_Q1", addr) == "ROUND((D8-D4)-SUM(D5:D6),0)"   # whole kroner: never "(0)"
