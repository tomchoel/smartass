# engine/tests/v2/test_v2_lines_initiative.py
"""add_initiative (lifted from v1 recipe lines_ranking): project grouping marks each line's booked project."""
from smartass.model import lines


def test_initiative_is_booked_project_else_derived_and_marked():
    rows = [{"PROJECT": "P123", "PROJECT_DESC": "5G rollout", "CC_DESC": "Radio", "AMOUNT": -1.0},
            {"PROJECT": "P124", "PROJECT_DESC": "", "AMOUNT": -1.0},
            {"PROJECT": "000000", "CC_DESC": "Core network", "AMOUNT": -1.0},
            {"PROJECT": "", "GL_DESCRIPTION_LINE": "KPMG 123456 audit fee 2026 Q3 extra work and much more text here",
             "AMOUNT": -1.0},
            {"AMOUNT": -1.0}]
    cols = lines.add_initiative(rows, ["AMOUNT"])
    assert cols == ["AMOUNT", "INITIATIVE", "INITIATIVE_DERIVED"]
    assert [(r["INITIATIVE"], r["INITIATIVE_DERIVED"]) for r in rows] == [
        ("5G rollout", ""), ("P124", ""), ("", "Core network"),
        ("", "KPMG audit fee Q3 extra work and much mo"), ("", lines.NO_DESC)]


def test_add_initiative_is_idempotent():
    rows = [{"PROJECT": "P1", "PROJECT_DESC": "X"}]
    cols = lines.add_initiative(rows, ["AMOUNT"])
    assert lines.add_initiative(rows, cols) == cols
