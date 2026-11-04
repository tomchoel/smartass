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
