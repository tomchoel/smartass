"""livexl snapshot logic that needs no Excel."""
from smartview_mcp.livexl import _coerced, _constants


def test_constants_skip_formula_cells_and_keep_types():
    f2 = [["Line", "2026"], ["=B2", ""]]          # Formula2 grid of B2:C3 ('2026' typed as text)
    v2 = [["Line", "2026"], [5.0, None]]          # Value2 grid
    assert _constants(f2, v2, 2, 2) == {"B2": "Line", "C2": "2026", "C3": None}


def test_coerced_finds_text_turned_number_and_leftovers():
    consts = {"B2": "Line", "C2": "2026", "D2": 80, "C3": None}
    now = [["Line", 2026.0, 80.0], [7.0, 1.0, None]]   # B2:D3 after the XML import
    assert _coerced(consts, now, 2, 2) == [("C2", "2026"), ("C3", None)]
