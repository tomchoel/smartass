"""Readable formulas (skill tn-formelstil): builders and lookups pass the checker."""
from smartview_mcp import formulas as fx
from smartview_mcp.grid import Axis, Layout
from smartview_mcp.userexcel import Pull, lookup_formula


def _pull(rows, cols, rl, cl, name="Q1 FinMax"):
    lay = Layout([("Entity", "un.SY")], rows, cols)
    return Pull(name, "FinMax", lay, None, 0, dict(zip(lay.row_tuples, rl)), dict(zip(lay.col_tuples, cl)))


def test_sums():
    assert fx.sum_rows("C", list(range(13, 22))) == "=SUM(C13:C21)"
    assert fx.sum_rows("C", [10, 20, 21, 22, 25]) == "=SUM(C10,C20:C22,C25)"
    assert fx.sum_rows("C", [30, 40]) == "=C30+C40"
    assert fx.sum_rows("C", [7]) == "=C7"
    assert fx.sum_cols(6, list(range(3, 12))) == "=SUM(C6:K6)"


def test_variance_pct_blank_on_zero_base():
    assert fx.variance_pct("C13", "D13") == '=IF(D13=0,"",(C13-D13)/ABS(D13))'


def test_plain_sum_expr():
    assert fx.expr_keys_if_plain_sum("a + b+c") == ["a", "b", "c"]
    assert fx.expr_keys_if_plain_sum("a+b") is None
    assert fx.expr_keys_if_plain_sum("a-b+c") is None


def test_lookup_single():
    p = _pull([Axis("Account", ["ac.C30", "ac.EBITDA"])], [Axis("Period", ["Jan 26", "Feb 26"])],
              [("C30 REVENUES (ac)",), ("EBITDA EBITDA (ac)",)], [("Jan 26",), ("Feb 26",)])
    assert lookup_formula(p, ["ac.EBITDA"], ["Feb 26"]) == (
        "=XLOOKUP(\"EBITDA EBITDA (ac)\",'Q1 FinMax'!$A:$A,"
        "XLOOKUP(\"Feb 26\",'Q1 FinMax'!$A$2:$C$2,'Q1 FinMax'!$A:$C))")


def test_lookup_multi_cols_and_rows():
    p = _pull([Axis("Entity", ["un.SY"])], [Axis("Account", ["ac.X", "ac.Y"]), Axis("Period", ["Sep 26", "Sep 25"])],
              [("SY (un)",)], [("X (ac)", "Sep 26"), ("X (ac)", "Sep 25"), ("Y (ac)", "Sep 26"), ("Y (ac)", "Sep 25")])
    f = lookup_formula(p, ["un.SY"], ["ac.Y", "Sep 25"])
    assert f.startswith("=XLOOKUP(\"SY (un)\",'Q1 FinMax'!$A:$A,XLOOKUP(1,(") and "$A$3:$E$3=\"Sep 25\"" in f
    p = _pull([Axis("Entity", ["un.A", "un.B"]), Axis("Account", ["ac.C30"])], [Axis("Period", ["Sep 26"])],
              [("A (un)", "C30 (ac)"), ("B (un)", "C30 (ac)")], [("Sep 26",)])
    assert lookup_formula(p, ["un.B", "ac.C30"], ["Sep 26"]) == (
        "=INDEX('Q1 FinMax'!$A:$C,XMATCH(1,('Q1 FinMax'!$A:.$A=\"B (un)\")*('Q1 FinMax'!$B:.$B=\"C30 (ac)\")),"
        "XMATCH(\"Sep 26\",'Q1 FinMax'!$A$2:$C$2))")


def test_generated_formulas_pass_checker():
    p = _pull([Axis("Account", ["ac.C30"])], [Axis("Period", ["Sep 26"])], [("C30 REVENUES (ac)",)], [("Sep 26",)])
    fs = [lookup_formula(p, ["ac.C30"], ["Sep 26"]), fx.sum_rows("C", list(range(13, 22))), fx.sum_rows("C", [30, 40]),
          fx.variance("C13", "D13"), fx.variance_pct("C13", "D13"), fx.sum_cols(6, list(range(3, 12)))]
    res = fx.check([("Q1 Svar", f"C{i + 1}", f) for i, f in enumerate(fs)])
    assert res["ok"] is True, res
    bad = fx.check([("Q1 Svar", "C1", "=C13+C14+C15")])
    assert bad["ok"] is False and bad["block"][0]["rule"] == "PLUS_CHAIN"


def test_avg_rows_is_an_average_over_ranges_and_passes_the_checker():
    assert fx.avg_rows("C", list(range(20, 30))) == "=AVERAGE(C20:C29)"
    assert fx.avg_rows("D", [5, 7, 8]) == "=AVERAGE(D5,D7:D8)"
    assert fx.check([("Q1 Svar", "C30", fx.avg_rows("C", list(range(20, 30))))])["ok"] is True
    import pytest
    with pytest.raises(ValueError):
        fx.avg_rows("C", [])
