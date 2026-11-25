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
    assert o.headers == ["Amount", "Share %", "Lines"]


def test_gap_is_a_warning_finding():
    res, _ = run(cube=-170.0)
    assert any("gap" in f.text.lower() for f in res.findings)
    res, _ = run(cube=-165.0)
    assert not any("gap" in f.text.lower() for f in res.findings)


def test_cell_needs_account_and_period():
    with pytest.raises(ValidationError, match="All Accounts"):
        AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "lines", "cell": {"All Time Periods": "Yr 26"}}]))


def _run_custom(rows, *, group_by="supplier", top=10, cols=None, cube=None):
    d = spec_dict(blocks=[{"type": "lines", "cell": CELL, "group_by": group_by, "top": top}])
    s = AnalysisSpec.model_validate(d)
    fr = Frame([])
    cnames = cols or sorted({k for r in rows for k in r})
    fr.lines["b0"] = LineSet("b0", rows, cnames, cube_value=cube)
    res = compute(s, resolver(), CLOSED, fr)
    return res, res.outs[0]


def test_group_by_account_uses_code_dominant_name_and_fixes_grammar():
    rows = [
        {"ACCOUNT": "43200", "ACCOUNT_NAME": "Consultancy Old", "AMOUNT": -10.0},
        {"ACCOUNT": "43200", "ACCOUNT_NAME": "Consultancy Main", "AMOUNT": -90.0},
        {"ACCOUNT": "41000", "ACCOUNT_NAME": "", "AMOUNT": -40.0},
        {"ACCOUNT": "", "ACCOUNT_NAME": "", "AMOUNT": -5.0},
    ]
    res, o = _run_custom(rows, group_by="account", top=5)
    by_key = {ln.key: ln for ln in o.lines}
    assert (by_key["g0"].code, by_key["g0"].label) == ("43200", "Consultancy Main")
    assert (by_key["g1"].code, by_key["g1"].label) == ("41000", "41000")
    assert by_key["blank"].label == "Lines without an account"
    assert res.cells.get(ck(o.key, "g0", 2)) == 2.0


def test_group_by_project_and_supplier_codes():
    rows = [
        {"AP_CUSTOMER_NAME": "Vendor A", "INITIATIVE": "Proj Alpha", "PROJECT": "00000", "ACCOUNT": "41000", "AMOUNT": -20.0},
        {"AP_CUSTOMER_NAME": "Vendor A", "INITIATIVE": "Proj Alpha", "PROJECT": "P100", "ACCOUNT": "42000", "AMOUNT": -80.0},
        {"AP_CUSTOMER_NAME": "Vendor A", "INITIATIVE": "Proj Alpha", "PROJECT": "P200", "ACCOUNT": "43000", "AMOUNT": -10.0},
    ]
    _, o_sup = _run_custom(rows, group_by="supplier", top=5)
    assert o_sup.lines[0].code == "42000 +2"
    _, o_proj = _run_custom(rows, group_by="project", top=5)
    assert (o_proj.lines[0].code, o_proj.lines[0].label) == ("P100", "Proj Alpha")
    assert all(ln.key != "blank" for ln in o_proj.lines)   # 0-line blank row suppressed


def test_zero_blank_rows_suppressed_with_other():
    rows = [
        {"AP_CUSTOMER_NAME": "A", "AMOUNT": -50.0},
        {"AP_CUSTOMER_NAME": "B", "AMOUNT": -30.0},
        {"AP_CUSTOMER_NAME": "C", "AMOUNT": -10.0},
    ]
    res, o = _run_custom(rows, group_by="supplier", top=2)
    keys = [ln.key for ln in o.lines]
    assert "blank" not in keys and "other" in keys
    assert res.cells.get(ck(o.key, "other", 0)) == -10.0
    assert res.cells.get(ck(o.key, "other", 2)) == 1.0


def test_mixed_signs_switches_to_debit_credit_lines():
    rows = [
        {"AP_CUSTOMER_NAME": "Vendor X", "AMOUNT": -120.0},
        {"AP_CUSTOMER_NAME": "Vendor X", "AMOUNT": 30.0},
        {"AP_CUSTOMER_NAME": "Vendor Y", "AMOUNT": -50.0},
    ]
    res, o = _run_custom(rows, group_by="supplier", top=5)
    assert o.headers == ["Amount", "Debit", "Credit", "Lines"]
    assert o.cols == ["amount", "debit", "credit", "lines"]
    assert res.cells.get(ck(o.key, "g0", 0)) == -90.0
    assert res.cells.get(ck(o.key, "g0", 1)) == -120.0
    assert res.cells.get(ck(o.key, "g0", 2)) == 30.0
    assert res.cells.get(ck(o.key, "g0", 3)) == 2.0
    assert res.cells.get(ck(o.key, "total", 1)) == -170.0
    assert res.cells.get(ck(o.key, "total", 2)) == 30.0
    assert res.cells.get(ck(o.key, "total", 3)) == 3.0


def test_no_supplier_children_top_five_journals():
    rows = [
        {"AP_CUSTOMER_NAME": "Named AS", "JE_SOURCE": "AP", "JOURNAL_NAME": "AP1", "ACCOUNT": "40000", "AMOUNT": -50.0},
        {"AP_CUSTOMER_NAME": "", "JE_SOURCE": "Spreadsheet", "JOURNAL_NAME": "HG 12 Sep 26", "ACCOUNT": "43100", "AMOUNT": -90.0},
        {"AP_CUSTOMER_NAME": "", "JE_SOURCE": "Spreadsheet", "JOURNAL_NAME": "Reverses HG 12 Aug 26", "ACCOUNT": "43100", "AMOUNT": -10.0},
        {"AP_CUSTOMER_NAME": "", "JE_SOURCE": "ABA", "JOURNAL_NAME": "ABA batch", "ACCOUNT": "43200", "AMOUNT": -80.0},
        {"AP_CUSTOMER_NAME": "", "JE_SOURCE": "Spreadsheet", "JOURNAL_NAME": "DATABRONN2609", "ACCOUNT": "43300", "AMOUNT": -70.0},
        {"AP_CUSTOMER_NAME": "", "JE_SOURCE": "SRC4", "JOURNAL_NAME": "", "ACCOUNT": "43400", "AMOUNT": -60.0},
        {"AP_CUSTOMER_NAME": "", "JE_SOURCE": "SRC5", "JOURNAL_NAME": "", "ACCOUNT": "43500", "AMOUNT": -50.0},
        {"AP_CUSTOMER_NAME": "", "JE_SOURCE": "", "JOURNAL_NAME": "", "ACCOUNT": "43600", "AMOUNT": -10.0},
    ]
    res, o = _run_custom(rows, group_by="supplier", top=5)
    children = [ln for ln in o.lines if ln.level == 1]
    assert len(children) == 5
    assert [(ln.label, ln.code, res.cells.get(ck(o.key, ln.key, 0)), res.cells.get(ck(o.key, ln.key, 2)))
            for ln in children] == [
        ("HG 12", "43100", -100.0, 2.0),
        ("ABA", "43200", -80.0, 1.0),
        ("DATABRONN", "43300", -70.0, 1.0),
        ("SRC4", "43400", -60.0, 1.0),
        ("SRC5", "43500", -50.0, 1.0),
    ]


def test_no_supplier_spreadsheet_journal_not_double_counted_with_other_source():
    rows = [
        {"AP_CUSTOMER_NAME": "", "JE_SOURCE": "Spreadsheet", "JOURNAL_NAME": "HG 12 Sep 26", "ACCOUNT": "43100", "AMOUNT": -100.0},
        {"AP_CUSTOMER_NAME": "", "JE_SOURCE": "Manual", "JOURNAL_NAME": "HG 12 Sep 26", "ACCOUNT": "43200", "AMOUNT": -50.0},
    ]
    res, o = _run_custom(rows, group_by="supplier", top=5)
    children = {ln.label: (res.cells.get(ck(o.key, ln.key, 0)), res.cells.get(ck(o.key, ln.key, 2)))
                for ln in o.lines if ln.level == 1}
    assert children == {"HG 12": (-100.0, 1.0), "Manual": (-50.0, 1.0)}
    assert res.cells.get(ck(o.key, "blank", 0)) == -150.0


def test_no_supplier_only_journal_column_never_references_missing_je_source():
    rows = [
        {"AP_CUSTOMER_NAME": "", "JOURNAL_NAME": "HG 12 Sep 26", "AMOUNT": -100.0},
        {"AP_CUSTOMER_NAME": "", "JOURNAL_NAME": "", "AMOUNT": -10.0},
    ]
    res, o = _run_custom(rows, group_by="supplier", top=5, cols=["AP_CUSTOMER_NAME", "JOURNAL_NAME", "AMOUNT"])
    for ln in o.lines:
        if ln.level == 1:
            f = ln.cells[0].to_excel("Data_Q1", str)
            assert "JE_SOURCE" not in f
    children = {ln.label: res.cells.get(ck(o.key, ln.key, 0)) for ln in o.lines if ln.level == 1}
    assert children == {"HG 12": -100.0, "(no journal/source)": -10.0}
