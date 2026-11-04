# engine/tests/v2/test_v2_readback.py
from smartass.render.readback import compare


class FakeCells:
    def __init__(self, d):
        self.d = d

    def get(self, k):
        return self.d[k]


def grid(rows, cols, values):
    """values: {(r, c): v} -> Value2-like tuple from A1 to (rows, cols)."""
    return tuple(tuple(values.get((r, c)) for c in range(1, cols + 1)) for r in range(1, rows + 1))


def test_equal_sheet_has_no_mismatch():
    addr = {"a": "C4", "b": "D4", "m": "C5"}
    cells = FakeCells({"a": -85.0, "b": None, "m": 0.1})
    sheet = grid(5, 4, {(4, 3): -85.0, (4, 4): 0.0, (5, 3): 0.1000000001})
    assert compare(addr, cells, sheet) == []


def test_mismatch_and_error_values_are_reported():
    addr = {"a": "C4", "b": "D4"}
    cells = FakeCells({"a": -85.0, "b": 2.0})
    sheet = grid(4, 4, {(4, 3): -84.0, (4, 4): -2146826281})        # #N/A comes back as an int
    out = compare(addr, cells, sheet)
    assert len(out) == 2 and out[0].startswith("C4: sheet -84")


def test_blank_and_empty_string_count_as_zero():
    addr = {"a": "C4"}
    assert compare(addr, FakeCells({"a": None}), grid(4, 3, {(4, 3): ""})) == []


def test_near_zero_amounts_are_equal():
    """Live 04.11.2026: 'C21: sheet -8.79e-07, model 1.94e-07' - a bridge 'Other' that is 0 up to float noise
    (SUMIFS over large NOK amounts vs the model's own sum)."""
    addr = {"a": "C4"}
    assert compare(addr, FakeCells({"a": 1.94e-07}), grid(4, 3, {(4, 3): -8.79e-07})) == []
    assert compare(addr, FakeCells({"a": -2.2e-07}), grid(4, 3, {(4, 3): -2.6e-06})) == []


def test_near_zero_fractions_are_still_compared():
    addr = {"p": "C4"}
    out = compare(addr, FakeCells({"p": 0.0001}), grid(4, 3, {(4, 3): 0.0005}), fractions={"p"})
    assert len(out) == 1 and out[0].startswith("C4:")


def test_an_amount_just_above_the_noise_is_still_a_mismatch():
    addr = {"a": "C4"}
    assert len(compare(addr, FakeCells({"a": 0.0}), grid(4, 3, {(4, 3): 0.002}))) == 1
