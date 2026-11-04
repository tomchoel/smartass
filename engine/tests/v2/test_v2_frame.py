import pytest

from smartass.model.frame import COLUMNS, Fact, Frame


def f(acc, ent, per, v):
    return Fact("FinMax", "fm", acc, ent, per, "Actual", "", v)


FR = Frame([f("ac.C50S", "E1", "Sep 26", -40.0), f("ac.C60", "E1", "Sep 26", -20.0),
            f("ac.C50S", "E2", "Sep 26", -10.0)])


def test_total_sums_matching_facts():
    assert FR.total({"Account": ("ac.C50S", "ac.C60"), "Entity": ("E1",)}) == -60.0
    assert FR.total({"Account": ("ac.C50S",)}) == -50.0


def test_total_is_none_when_nothing_matches():
    assert FR.total({"Entity": ("E9",)}) is None


def test_unknown_filter_column_is_an_error():
    with pytest.raises(KeyError, match="Colour"):
        FR.total({"Colour": ("red",)})


def test_rows_follow_columns():
    assert COLUMNS[-1] == "Value"
    assert FR.rows()[0] == ("FinMax", "fm", "ac.C50S", "E1", "Sep 26", "Actual", "", -40.0)
