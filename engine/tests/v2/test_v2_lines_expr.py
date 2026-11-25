from smartass.model.expr import Cells, CubeRef, LinePick
from smartass.model.frame import Frame
from smartass.model.lines import LineSet

ROWS = [{"AP_CUSTOMER_NAME": "Accenture AS", "AMOUNT": -100.0},
        {"AP_CUSTOMER_NAME": "accenture as", "AMOUNT": -20.0},
        {"AP_CUSTOMER_NAME": "Bouvet", "AMOUNT": -30.0},
        {"AP_CUSTOMER_NAME": "", "AMOUNT": -5.0}]


def frame():
    fr = Frame([])
    fr.lines["b0"] = LineSet("b0", ROWS, ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=-160.0)
    return fr


def ev(e, fr):
    return e.evaluate(fr, Cells(fr, {}))


def test_line_pick_sums_amount_case_insensitively():
    fr = frame()
    assert ev(LinePick("b0", "AP_CUSTOMER_NAME", ("Accenture AS", "accenture as")), fr) == -120.0
    assert ev(LinePick("b0", "AP_CUSTOMER_NAME", ("",)), fr) == -5.0
    assert ev(LinePick("b0", None, None), fr) == -155.0


def test_line_pick_formula_on_the_lines_table():
    e = LinePick("b0", "AP_CUSTOMER_NAME", ("Accenture AS",))
    assert e.to_excel("Data_Q3", str) == 'SUMIFS(Lines_Q3_b0[AMOUNT],Lines_Q3_b0[AP_CUSTOMER_NAME],"Accenture AS")'
    assert LinePick("b0", None, None).to_excel("Data_Q3", str) == "SUM(Lines_Q3_b0[AMOUNT])"
    blank = LinePick("b0", "AP_CUSTOMER_NAME", ("",)).to_excel("Data_Q3", str)
    assert blank.endswith(',"")')


def test_cube_ref():
    fr = frame()
    assert ev(CubeRef("b0"), fr) == -160.0
    assert CubeRef("b0").to_excel("Data_Q3", str) == "'Q3 Lines b0'!$D$3"
    assert CubeRef("b0").to_excel("Lines_Q4_b0", str) == "'Q4 Lines b0'!$D$3"


def test_line_pick_count_debit_credit_and_extra():
    rows = [
        {"AP_CUSTOMER_NAME": "Accenture", "JE_SOURCE": "ABA", "AMOUNT": -100.0},
        {"AP_CUSTOMER_NAME": "Accenture", "JE_SOURCE": "ABA", "AMOUNT": 25.0},
        {"AP_CUSTOMER_NAME": "Accenture", "JE_SOURCE": "MANUAL", "AMOUNT": -40.0},
    ]
    fr = Frame([])
    fr.lines["b0"] = LineSet("b0", rows, ["AP_CUSTOMER_NAME", "JE_SOURCE", "AMOUNT"], cube_value=-115.0)
    extra = (("JE_SOURCE", ("ABA",)),)
    assert ev(LinePick("b0", "AP_CUSTOMER_NAME", ("Accenture",), agg="count", extra=extra), fr) == 2.0
    assert ev(LinePick("b0", "AP_CUSTOMER_NAME", ("Accenture",), agg="debit", extra=extra), fr) == -100.0
    assert ev(LinePick("b0", "AP_CUSTOMER_NAME", ("Accenture",), agg="credit", extra=extra), fr) == 25.0
    assert ev(LinePick("b0", None, None, agg="count"), fr) == 3.0
    assert ev(LinePick("b0", None, None, agg="debit"), fr) == -140.0
    assert ev(LinePick("b0", None, None, agg="credit"), fr) == 25.0

    assert LinePick("b0", None, None, agg="count").to_excel("Lines_Q3_b0", str) == "ROWS(Lines_Q3_b0[AMOUNT])"
    assert (LinePick("b0", None, None, agg="debit").to_excel("Lines_Q3_b0", str)
            == 'SUMIFS(Lines_Q3_b0[AMOUNT],Lines_Q3_b0[AMOUNT],"<0")')
    assert (LinePick("b0", "AP_CUSTOMER_NAME", ("Accenture",), agg="credit", extra=extra).to_excel("Lines_Q3_b0", str)
            == 'SUMIFS(Lines_Q3_b0[AMOUNT],Lines_Q3_b0[AMOUNT],">0",'
               'Lines_Q3_b0[AP_CUSTOMER_NAME],"Accenture",Lines_Q3_b0[JE_SOURCE],"ABA")')


def test_line_pick_safe_sumifs_fallback():
    for unsafe in ("0047", "Sep-26", "*Co", "=SUM", "x" * 256):
        f_sum = LinePick("b0", "ACCOUNT", (unsafe,), agg="sum").to_excel("Data_Q1", str)
        f_cnt = LinePick("b0", "ACCOUNT", (unsafe,), agg="count").to_excel("Data_Q1", str)
        f_deb = LinePick("b0", "ACCOUNT", (unsafe,), agg="debit").to_excel("Data_Q1", str)
        assert f_sum.startswith("SUM(FILTER(")
        assert f_cnt == f'SUM(--(Lines_Q1_b0[ACCOUNT]="{unsafe}"))'
        assert "<0" in f_deb and f_deb.startswith("SUM(FILTER(")


def test_line_pick_filter_fallback_handles_blank_in_multi_values():
    f = LinePick("b0", "AP_CUSTOMER_NAME", ("0047",),
                 extra=(("JE_SOURCE", ("", "Spreadsheet")),)).to_excel("Data_Q1", str)
    assert f.startswith("SUM(FILTER(")
    assert 'Lines_Q1_b0[JE_SOURCE]=""' in f
    assert 'Lines_Q1_b0[JE_SOURCE]="Spreadsheet"' in f
    assert 'MATCH(Lines_Q1_b0[JE_SOURCE],{"","Spreadsheet"}' not in f
