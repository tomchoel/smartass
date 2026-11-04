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
    assert CubeRef("b0").to_excel("Data_Q3", str) == "'Q3 Lines b0'!$D$2"
