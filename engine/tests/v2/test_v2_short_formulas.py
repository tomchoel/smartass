from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.expr import pick
from smartass.model.spec import AnalysisSpec
from smartass.render.layout import constant_columns, layout


def test_pick_drops_criteria_that_match_the_whole_table():
    e = pick({"Cube": ("FinMax",), "Account": ("ac.C30",), "Period": ("Sep 26", "Aug 26")})
    full = e.to_excel("T", str)
    short = e.to_excel("T", str, const={"Cube": "FinMax"})
    assert "T[Cube]" in full and "T[Cube]" not in short and "T[Account]" in short
    assert e.to_excel("T", str, const={"Cube": "FinLite"}) == full     # a different constant is kept


def test_layout_uses_the_frames_constant_columns():
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}]))
    fr = facts()
    assert {"Cube": "FinMax", "Query": "fm", "Scenario": "Actual"}.items() <= constant_columns(fr).items()
    p = layout(1, s, compute(s, resolver(), CLOSED, fr), stock=set(), subtitle="x", const=constant_columns(fr))
    f = next(v for v in p.cells.values() if isinstance(v, str) and "SUMIFS" in v)
    assert "[Cube]" not in f and "[Query]" not in f and "[Account]" in f
