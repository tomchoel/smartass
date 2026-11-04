import pytest

from smartass.model.expr import Cells, Pick, Ref, Sub, Sum, pick, share, var_pct
from smartass.model.frame import Fact, Frame

FR = Frame([Fact("FinMax", "fm", "ac.C50S", "E1", "Sep 26", "Actual", "", -40.0),
            Fact("FinMax", "fm", "ac.C60", "E1", "Sep 26", "Actual", "", -20.0),
            Fact("FinMax", "fm", "ac.C50S", "E1", "Sep 25", "Actual", "", -30.0)])
ADDR = {"a": "D4", "b": "D5", "c": "D6", "x": "F9"}.__getitem__


def ev(e, exprs=None):
    return e.evaluate(FR, Cells(FR, exprs or {}))


def test_pick_single_values():
    e = pick({"Account": ("ac.C50S",), "Period": ("Sep 26",)})
    assert ev(e) == -40.0
    assert e.to_excel("T", ADDR) == 'SUMIFS(T[Value],T[Account],"ac.C50S",T[Period],"Sep 26")'


def test_pick_one_and_two_multi_filters_use_row_then_column_arrays():
    one = pick({"Account": ("ac.C50S", "ac.C60"), "Period": ("Sep 26",)})
    assert ev(one) == -60.0
    assert one.to_excel("T", ADDR) == 'SUM(SUMIFS(T[Value],T[Account],{"ac.C50S","ac.C60"},T[Period],"Sep 26"))'
    two = pick({"Account": ("ac.C50S", "ac.C60"), "Period": ("Sep 26", "Sep 25")})
    assert ev(two) == -90.0
    assert two.to_excel("T", ADDR) == ('SUM(SUMIFS(T[Value],T[Account],{"ac.C50S","ac.C60"},'
                                       'T[Period],{"Sep 26";"Sep 25"}))')


def test_pick_splits_when_three_filters_have_several_values():
    e = pick({"Account": ("ac.C50S", "ac.C60"), "Entity": ("E1", "E2"), "Period": ("Sep 26", "Sep 25")})
    assert isinstance(e, Sum) and len(e.items) == 2
    assert ev(e) == -90.0


def test_pick_quotes_are_doubled_and_empty_filter_refused():
    assert '"a""b"' in pick({"Entity": ('a"b',)}).to_excel("T", ADDR)
    with pytest.raises(ValueError, match="no values"):
        pick({"Entity": ()})


def test_missing_is_none_not_zero():
    assert ev(pick({"Entity": ("E9",)})) is None


def test_refs_sum_and_sub():
    exprs = {"a": pick({"Account": ("ac.C50S",), "Period": ("Sep 26",)}),
             "b": pick({"Account": ("ac.C60",), "Period": ("Sep 26",)}),
             "c": Sum((Ref("a"), Ref("b")))}
    cells = Cells(FR, exprs)
    assert cells.get("c") == -60.0
    assert exprs["c"].to_excel("T", ADDR) == "SUM(D4:D5)"
    d = Sub(Ref("a"), Sub(Ref("b"), Ref("c")))
    assert d.evaluate(FR, cells) == -40.0 - (-20.0 - -60.0)
    assert d.to_excel("T", ADDR) == "D4-(D5-D6)"


def test_sum_of_non_adjacent_refs_lists_them():
    assert Sum((Ref("a"), Ref("c"), Ref("x"))).to_excel("T", ADDR) == "SUM(D4,D6,F9)"


def test_var_pct_and_share_blank_on_zero_base():
    exprs = {"a": pick({"Account": ("ac.C50S",), "Period": ("Sep 26",)}),
             "b": pick({"Account": ("ac.C50S",), "Period": ("Sep 25",)})}
    cells = Cells(FR, exprs)
    v = var_pct(Ref("a"), Ref("b"))
    assert v.evaluate(FR, cells) == pytest.approx((-40.0 - -30.0) / 30.0)
    assert v.to_excel("T", ADDR) == 'IF(D5=0,"",(D4-D5)/ABS(D5))'
    assert share(Ref("a"), Ref("b")).to_excel("T", ADDR) == 'IF(D5=0,"",(D4)/(D5))'
    zero = Cells(FR, {"z": pick({"Entity": ("E9",)}), "a": exprs["a"]})
    assert var_pct(Ref("a"), Ref("z")).evaluate(FR, zero) is None


def test_cells_detect_cycles_and_unknown_keys():
    cells = Cells(FR, {"a": Ref("b"), "b": Ref("a")})
    with pytest.raises(ValueError, match="circular"):
        cells.get("a")
    with pytest.raises(KeyError, match="unknown cell"):
        Cells(FR, {}).get("nope")


def test_pick_is_hashable_tree():
    assert isinstance(pick({"Entity": ("E1",)}), Pick)
    hash(pick({"Entity": ("E1",), "Period": ("Sep 26", "Sep 25")}))
    hash(pick({"A": ("1", "2"), "Account": ("3", "4"), "Entity": ("5", "6")}))


def test_share_brackets_compound_denominator():
    e = share(Ref("a"), Sub(Ref("b"), Ref("c")))
    assert e.to_excel("T", ADDR) == 'IF(D5-D6=0,"",(D4)/(D5-D6))'


def test_cells_not_stuck_busy_after_error():
    cells = Cells(FR, {"a": Ref("nope")})
    for _ in range(2):
        with pytest.raises(KeyError):
            cells.get("a")


def test_pick_rejects_string_and_dedupes():
    with pytest.raises(TypeError, match="filter Entity: give a tuple of values, not a string"):
        pick({"Entity": "E1"})
    dup = pick({"Account": ("ac.C50S", "ac.C50S"), "Period": ("Sep 26",)})
    assert dup.to_excel("T", str).count("ac.C50S") == 1
    assert ev(dup) == ev(pick({"Account": ("ac.C50S",), "Period": ("Sep 26",)}))


def test_pick_splits_on_fewest_valued_multi_column():
    e = pick({"Account": ("a", "b", "c"), "Entity": ("x", "y"), "Period": ("p", "q", "r", "s")})
    assert isinstance(e, Sum) and len(e.items) == 2
    assert all(len(dict(i.filters)["Entity"]) == 1 for i in e.items)


def test_criteria_wildcards_escaped():
    f = pick({"Entity": ("E*",)}).to_excel("T", str)
    assert '"E~*"' in f
    assert '"a~~b~?"' in pick({"Entity": ("a~b?",)}).to_excel("T", str)


def test_sum_needs_items():
    with pytest.raises(ValueError, match="Sum needs at least one item"):
        Sum(())


def test_neg_formula_and_value():
    from smartass.model.expr import Neg
    exprs = {"a": pick({"Account": ("ac.C50S",), "Period": ("Sep 26",)}), "b": pick({"Account": ("ac.C60",),
             "Period": ("Sep 26",)})}
    cells = Cells(FR, exprs)
    assert Neg(Ref("a")).evaluate(FR, cells) == 40.0 and Neg(Ref("a")).to_excel("T", ADDR) == "-D4"
    assert Neg(Sub(Ref("a"), Ref("b"))).to_excel("T", ADDR) == "-(D4-D5)"
    assert Neg(pick({"Entity": ("E9",)})).evaluate(FR, cells) is None
