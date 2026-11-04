# engine/tests/v2/test_v2_block_lines.py
import pytest
from pydantic import ValidationError
from v2util import CLOSED, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.blocks.common import ck
from smartass.model.frame import Frame
from smartass.model.lines import LineSet
from smartass.model.spec import AnalysisSpec

ROWS = [{"AP_CUSTOMER_NAME": "Accenture AS", "AMOUNT": -100.0},
        {"AP_CUSTOMER_NAME": "accenture as", "AMOUNT": -20.0},
        {"AP_CUSTOMER_NAME": "Bouvet", "AMOUNT": -30.0},
        {"AP_CUSTOMER_NAME": "Sopra Steria", "AMOUNT": -10.0},
        {"AP_CUSTOMER_NAME": "", "AMOUNT": -5.0}]
CELL = {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Yr 26"}


def run(top=2, cube=-170.0):
    d = spec_dict(blocks=[{"type": "lines", "cell": CELL, "group_by": "supplier", "top": top}])
    s = AnalysisSpec.model_validate(d)
    fr = Frame([])
    fr.lines["b0"] = LineSet("b0", ROWS, ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=cube)
    res = compute(s, resolver(), CLOSED, fr)
    return res, res.outs[0]


def test_ranking_other_blank_total_cube_gap():
    res, o = run()
    keys = [ln.key for ln in o.lines]
    assert keys == ["g0", "g1", "other", "blank", "total", "cube", "gap"]
    assert [ln.label for ln in o.lines][:2] == ["Accenture AS", "Bouvet"]
    v = {ln.key: res.cells.get(ck(o.key, ln.key, 0)) for ln in o.lines}
    assert v == {"g0": -120.0, "g1": -30.0, "other": -10.0, "blank": -5.0, "total": -165.0,
                 "cube": -170.0, "gap": -5.0}
    assert o.headers == ["Amount", "Share %"]


def test_gap_is_a_warning_finding():
    res, _ = run(cube=-170.0)
    assert any("gap" in f.text.lower() for f in res.findings)
    res, _ = run(cube=-165.0)
    assert not any("gap" in f.text.lower() for f in res.findings)


def test_cell_needs_account_and_period():
    with pytest.raises(ValidationError, match="All Accounts"):
        AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "lines", "cell": {"All Time Periods": "Yr 26"}}]))
