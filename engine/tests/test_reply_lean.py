from smartview_mcp.core import reply


def _big():
    return {
        "workbook": "w.xlsx", "question": "Q1", "answer_sheet": "Q1 Answer", "sheets": ["Q1 Answer"],
        "data_sheets": [], "note": "n", "gate": "clear", "period": "Oct 26", "routes": {"Q1 FinMax": "http"},
        "key_points": ["k"], "followups": ["f"], "warnings": ["w"], "lines": {"d1": 3},
        "key_values": {"Revenue": {"Oct 26": 2247931399.1299996, "Change %": -0.5236831758915402}},
        "checks": {"a": "OK", "b": "OK", "c": "MISMATCH 12"},
        "review": {"ok": True, "findings": [{"severity": "Warning", "text": "x"}, {"severity": "Warning", "text": "x"}]},
        "material": [{"line": "s1", "label": "L", "cur": 1.0, "cmp": 2.0, "delta": -1.0, "pct": -0.5, "material": True,
                      "always": False}],
        "outline": {"recipe": "pnl", "pulls": []},
        "timings": {"_total": 10.0, "answer": 3.0, "answer.style": 2.0, "write source sheet": 2.5, "log sources": 1.0,
                    "open target": 0.1, "com calls.answer": 150, "first frame": 0.8, "drills": 8.0},
        "formula_style": {"ok": True, "checked": 9, "counts": {"LONG": 2}, "block": [], "warn": [{"cell": "A1"}] * 20},
        "style_lint": [f"l{i}" for i in range(9)],
        "table": "| Line |\n" + "\n".join(f"| r{i} |" for i in range(60)),
        "flags": [{"x": "y" * 100}] * 40, "landing": {"a": 1},
    }


def test_lean_keeps_what_the_answer_needs_and_trims_the_rest():
    out = reply.lean(_big(), "lean")
    assert out["key_values"]["Revenue"] == {"Oct 26": 2247931399, "Change %": -0.5237}
    assert out["checks"] == {"ok": 2, "c": "MISMATCH 12"}
    assert out["review"]["findings"] == [{"severity": "Warning", "text": "x"}]
    assert out["material"] == [{"line": "s1", "label": "L", "delta": -1.0, "pct": -0.5}]
    assert set(out["timings"]) == {"_total", "first frame", "drills", "answer", "write source sheet", "log sources"}
    assert out["formula_style"] == {"ok": True, "counts": {"LONG": 2}, "block": []}
    assert out["style_lint"] == [f"l{i}" for i in range(5)]
    assert "table" not in out                                  # 60 rows > 40: the numbers are in key_values
    assert out["flags"] == "40 items: ask with detail='full'" and out["landing"] == {"a": 1}
    for k in ("workbook", "question", "answer_sheet", "sheets", "routes", "key_points", "followups", "warnings",
              "outline", "gate", "period", "lines", "note"):
        assert out[k] == _big()[k]


def test_full_detail_and_small_tables_are_kept():
    assert reply.lean(_big(), "full") == _big()
    small = {**_big(), "table": "| a |\n| b |"}
    assert reply.lean(small, "lean")["table"] == "| a |\n| b |"
    assert reply.lean({"gate": "ambiguous", "questions": ["q"]}, "lean") == {"gate": "ambiguous", "questions": ["q"]}


def test_lean_keeps_ratios_and_fractions_but_rounds_money():
    kv = {"EBITDA margin": {"Sep 26": 0.2345678}, "Var": {"Variance": 0.01234}, "FTE": {"Sep 26": 0.5},
          "Revenue": {"Sep 26": 2247931399.13}}
    out = reply.lean({"key_values": kv}, "lean")["key_values"]
    assert out["EBITDA margin"] == {"Sep 26": 0.2346} and out["Var"] == {"Variance": 0.0123}
    assert out["FTE"] == {"Sep 26": 0.5} and out["Revenue"] == {"Sep 26": 2247931399}
    assert isinstance(out["Revenue"]["Sep 26"], int)


def test_lean_reports_lines_beyond_key_values():
    kv = {f"L{i}": {"Sep 26": 1.0} for i in range(40)}
    out = reply.lean({"key_values": kv, "key_values_total": 55}, "lean")
    assert out["more_lines"] == 15 and "detail='full'" in out["more_lines_note"]
    assert "more_lines" not in reply.lean({"key_values": kv, "key_values_total": 40}, "lean")


def test_line_count_counts_lines_with_values():
    spec = {"sections": [{"lines": [{"key": "a"}, {"key": "b"}, {}]}]}
    assert reply.line_count(spec, {"values": {"a": [1]}}) == 1


def test_gate_rejects_answer_keys_that_are_not_checklist_items():
    import pytest
    from smartview_mcp.core import scope
    from smartass.xl.host import InputError
    with pytest.raises(InputError, match="measure"):
        scope.gate("Opex this year", "Sep 26", learned=lambda i, p: None, answers={"bogus": "x"})
