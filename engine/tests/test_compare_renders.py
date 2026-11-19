"""scripts/compare_renders.py's diff, the parity gate's verdict: pure data in, findings out (no Excel)."""
import copy
import importlib.util
from pathlib import Path

_p = Path(__file__).resolve().parents[2] / "scripts" / "compare_renders.py"
_spec = importlib.util.spec_from_file_location("compare_renders", _p)
cr = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cr)


def _sheet(q=1):
    return {"missing": False, "used": "$A$1:$C$3", "rect": [1, 1, 3, 3],
            "cells": {"font name": [(1, 1, 3, 3, "Arial")], "font size": [(1, 1, 1, 3, 14), (2, 1, 3, 3, 9)],
                      "style": [(1, 1, 3, 3, "Normal")], "top border": [(1, 1, 3, 3, "none")]},
            "contents": {"error": False, "masked": 1,
                         "cells": {"A1": ["t", "abc"], "B2": ["n", 10.001], "C3": ["f", f"=SUM(Data_Q{q}[Value])"]}},
            "row heights": {"1:3": 15}, "column widths": {"1:2": 16, "3:3": 9.5},
            "print area": "$A$1:$C$3", "orientation": 2, "tab colour": 0, "tables": [], "freeze panes": [True, 3, 0],
            "names": [["'Q1 Answer'!Print_Area", "='Q1 Answer'!$A$1:$C$3", True]], "merges": [], "chart count": 1,
            "charts": [{"top left": "$E$5", "type": 51, "overlap": -27, "title": "h1",
                        "series": [{"formula": "=SERIES(,'Q1 Answer'!$B$2,'Q1 Answer'!$C$2,1)", "fill": 255}]}]}


def _side(sheet):
    record = {"workbook": "x.xlsx", "cases": {"b02": {"ok": True, "q": 1, "sheets": ["Q1 Answer"],
                                                      "reply": {"keys": ["q"], "status": "OK"}}}}
    return record, {"workbook": "x.xlsx", "cases": {"b02": {"Answer": sheet}}}


def _diff(a, b):
    (ra, da), (rb, db) = _side(a), _side(b)
    return cr.diff({"com": ra, "file": rb}, {"com": da, "file": db})


def test_a_dump_against_itself_has_no_difference():
    findings, stats = _diff(_sheet(), _sheet())
    assert findings == [] and stats["cases"] == 1 and stats["masked"] == 2


def test_one_font_one_width_and_one_series_formula_changed_are_exactly_three_findings():
    b = _sheet()
    b["cells"]["font name"] = [(1, 1, 1, 1, "Calibri"), (1, 2, 3, 3, "Arial"), (2, 1, 3, 1, "Arial")]
    b["column widths"] = {"1:1": 16, "2:2": 17, "3:3": 9.5}
    b["charts"][0]["series"][0]["formula"] = "=SERIES(,'Q1 Answer'!$B$2,'Q1 Answer'!$D$2,1)"
    findings, _ = _diff(_sheet(), b)
    assert [(f.prop, f.where, f.a, f.b) for f in findings] == [
        ("font name", "A1", "Arial", "Calibri"), ("column widths", "B", 16, 17),
        ("series formula", "chart at $E$5 series 1", _sheet()["charts"][0]["series"][0]["formula"],
         b["charts"][0]["series"][0]["formula"])]
    text, bad = cr.report(findings, {"cases": 1, "masked": 0, "errors": {}, "cells": 9})
    assert bad == 3 and "DIFF | Answer | font name | com 'Arial' -> file 'Calibri' | 1 in b02: b02 A1" in text


def test_numbers_within_half_a_hundredth_are_equal_and_values_are_never_shown():
    b = _sheet()
    b["contents"]["cells"]["B2"] = ["n", 10.006]
    assert _diff(_sheet(), b)[0] == []
    b["contents"]["cells"]["B2"] = ["n", 10.007]
    b["contents"]["cells"]["A1"] = ["t", "def"]
    findings, _ = _diff(_sheet(), b)
    assert [(f.prop, f.where, f.show) for f in findings] == [("value", "A1", False), ("value", "B2", False)]
    text, _ = cr.report(findings, {"cases": 1, "masked": 0, "errors": {}, "cells": 9})
    assert "10.0" not in text and "| differs |" in text


def test_the_allowed_differences_are_listed_but_do_not_fail():
    a, b = _sheet(), _sheet()
    b["cells"]["style"] = [(1, 1, 3, 3, "TN Tekst")]
    b["names"] = a["names"] + [["'Q1 Answer'!smartass_pagesetup", "=1", False]]
    findings, _ = _diff(a, b)
    assert {f.allowed for f in findings} == {"the cell style name", "the hidden sheet name smartass_pagesetup"}
    assert cr.report(findings, {"cases": 1, "masked": 0, "errors": {}, "cells": 9})[1] == 0


def test_a_data_width_may_be_one_character_under_and_three_over_autofit():
    a, b = _sheet(), _sheet()
    a["column widths"] = {"1:4": 16, "5:5": 9.5}
    b["column widths"] = {"1:1": 15.0, "2:2": 14.9, "3:3": 19.0, "4:4": 19.1, "5:5": 9.5}
    (ra, da), (rb, db) = _side(a), _side(b)
    da["cases"]["b02"], db["cases"]["b02"] = {"Data": a}, {"Data": b}
    findings, _ = cr.diff({"com": ra, "file": rb}, {"com": da, "file": db})
    assert [(f.where, bool(f.allowed)) for f in findings] == [("A", True), ("B", False), ("C", True), ("D", False)]


def _chart_pair(**changes):
    a, b = _sheet(), _sheet()
    for side, (k, v) in ((s, kv) for s, kvs in changes.items() for kv in kvs.items()):
        {"com": a, "file": b}[side]["charts"][0][k] = v
    return _diff(a, b)[0]


def test_a_chart_size_within_half_a_point_is_allowed_beyond_it_not():
    assert [bool(f.allowed) for f in _chart_pair(com={"height": 260.0}, file={"height": 260.2})] == [True]
    assert [bool(f.allowed) for f in _chart_pair(com={"width": 480.0}, file={"width": 480.6})] == [False]


def test_only_a_line_chart_groups_overlap_is_allowed_to_differ():
    line = _chart_pair(com={"type": 4, "overlap": 100}, file={"type": 4, "overlap": 0})
    column = _chart_pair(com={"overlap": -27}, file={"overlap": 0})
    assert [bool(f.allowed) for f in line] == [True] and [bool(f.allowed) for f in column] == [False]


def test_a_value_axis_bound_is_never_shown():
    findings = _chart_pair(com={"value axis minimum": 12345.6}, file={"value axis minimum": 12000})
    text, bad = cr.report(findings, {"cases": 1, "masked": 0, "errors": {}, "cells": 9})
    assert bad == 1 and "12345" not in text and "12000" not in text and "value axis minimum | differs" in text


def test_a_case_that_failed_on_one_side_is_a_finding_and_its_sheets_are_not_compared():
    (ra, da), (rb, _) = _side(_sheet()), _side(_sheet())
    rb["cases"]["b02"] = {"ok": False, "error": "'boom'"}
    findings, stats = cr.diff({"com": ra, "file": rb}, {"com": da, "file": {"cases": {}}})
    assert [(f.sheet, f.prop, f.a, f.b) for f in findings] == [("draw", "drawn", "ok", "'boom'")]
    assert stats["cases"] == 0


def test_q_numbers_are_normalised_so_a_shifted_pair_compares():
    assert cr.normalise("=SUM(Data_Q3[Value])+'Q3 Lines b0'!G3", 3) == "=SUM(Data_Q#[Value])+'Q# Lines b0'!G3"
    assert cr.normalise("Q31 Answer", 3) == "Q31 Answer"


def test_charts_pair_by_their_top_left_cell():
    a, b = _sheet(), copy.deepcopy(_sheet())
    b["charts"][0]["top left"] = "$E$30"
    findings, _ = _diff(a, b)
    assert [(f.prop, f.where, f.a, f.b) for f in findings] == [
        ("chart", "chart at $E$30", "missing", "present"), ("chart", "chart at $E$5", "present", "missing")]
