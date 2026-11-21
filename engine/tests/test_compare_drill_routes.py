"""scripts/compare_drill_routes.py's judge and its output (Task 7's gate for deleting the browser route): rows
compared as sorted JSON, the tie, the cube value and the columns each fail the comparison, and nothing it prints
carries a figure, a business text, a token or a query value. The live run itself is the script's, on the owner's PC."""
import importlib.util
import json
from pathlib import Path

_p = Path(__file__).resolve().parents[2] / "scripts" / "compare_drill_routes.py"
_spec = importlib.util.spec_from_file_location("compare_drill_routes", _p)
cmp = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cmp)

CELL = {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"}
ROWS = [{"AMOUNTX": 123456.78, "Supplier": "Leverandør AS", "Account": "4010"},
        {"AMOUNTX": -98765.43, "Supplier": "Annen Leverandør", "Account": "4020"}]


def side(rows=ROWS, columns=("AMOUNTX", "Supplier", "Account"), status="exact", cube=24691.35, secs=12.34):
    return {"ok": True, "rows": cmp.row_keys(rows), "lines": len(rows), "columns": list(columns), "status": status,
            "cell_value": cube, "secs": secs}


def test_the_same_rows_in_another_order_are_the_same_lines():
    c = cmp.compare(side(), side(rows=list(reversed(ROWS)), secs=3.2))
    assert c["same_lines"] and c["tie_equal"] and c["cube_equal"] and c["ok"]


def test_a_changed_amount_or_a_missing_row_is_not_the_same_lines():
    changed = [dict(ROWS[0], AMOUNTX=123456.79), ROWS[1]]
    assert not cmp.compare(side(), side(rows=changed))["same_lines"]
    assert not cmp.compare(side(), side(rows=ROWS[:1]))["ok"]


def test_columns_tie_and_cube_value_each_fail_the_comparison():
    assert not cmp.compare(side(), side(columns=("AMOUNTX", "Account", "Supplier")))["same_lines"]
    c = cmp.compare(side(), side(status="gap"))
    assert c["same_lines"] and not c["tie_equal"] and not c["ok"]
    assert not cmp.compare(side(), side(cube=24691.36))["ok"]
    assert cmp.compare(side(), side(cube=24691.354))["cube_equal"]
    assert cmp.compare(side(cube=None), side(cube=None))["cube_equal"]
    assert not cmp.compare(side(cube=None), side())["cube_equal"]


def test_a_route_that_failed_fails_the_comparison():
    c = cmp.compare(side(), {"ok": False, "error": "DrillError: x"})
    assert not c["ok"] and not c["same_lines"]


def test_the_line_shows_counts_booleans_tie_and_seconds_only():
    text = cmp.line(CELL, cmp.compare(side(), side(secs=3.21)))
    assert text == ("ac.C60 / un.SY.2TEC / Sep 26: lines 2/2, same lines: True, tie browser/token: exact/exact, "
                    "cube value equal: True, seconds 12.3/3.2")
    assert not any(s in text for s in ("123456", "98765", "24691", "Leverand", "4010"))


def test_a_failed_route_shows_its_error_kind_in_the_line():
    text = cmp.line(CELL, cmp.compare(side(), {"ok": False, "error": "NotSignedIn: run smartass login"}))
    assert "token failed: NotSignedIn" in text and "same lines: False" in text


def test_the_view_diff_names_key_paths_and_types_never_values():
    a = {"columns": [{"name": "Supplier", "width": 120}, {"name": "Amount"}], "ts": 1795272710098, "title": "Kostnad"}
    b = {"columns": [{"name": "Supplier", "width": "120"}, {"name": "Amount"}], "ts": 1795272799999, "extra": True}
    d = cmp.view_diff(a, b)
    assert d["only_browser"] == ["$.title"] and d["only_token"] == ["$.extra"]
    assert d["type_differs"] == ["$.columns[0].width (int/str)"] and d["value_differs"] == ["$.ts"]
    shown = json.dumps(d) + "\n".join(cmp.view_lines(a, b))
    assert not any(s in shown for s in ("Kostnad", "1795272710098", "1795272799999", "Supplier"))


def test_a_link_that_differs_is_shown_by_its_parts_never_its_values():
    a = {"links": [{"href": "https://pod/aif/ui/model/view?page=DRILL_THROUGH&appId=52&_=1795272710098"}]}
    b = {"links": [{"href": "https://pod/aif/ui/model/view?page=DRILL_THROUGH&appId=52&_=1795272799999"}]}
    d = cmp.view_diff(a, b)
    assert d["value_differs"] == ["$.links[0].href (a link: same path True, query items that differ: _)"]
    assert "1795272" not in "\n".join(cmp.view_lines(a, b))


def test_the_view_lines_say_whether_the_columns_differ():
    a = {"columns": [{"name": "A"}]}
    text = "\n".join(cmp.view_lines(a, a))
    assert "same documents: True" in text


def test_an_unlisted_query_item_is_named_by_its_key_never_its_value():
    msg = "refused: GET https://pod/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?... with a query item this " \
          "request does not take"
    url = "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.drilldown&FORMAT=SUI&secretish=VALUE42"
    kind = cmp.refused_kind(msg, url, None)
    assert "secretish" in kind and "VALUE42" not in kind and "module" not in kind.split("item")[-1]


def test_a_non_ascii_query_item_is_named_by_its_key():
    kind = cmp.refused_kind("refused: a URL with a space, a control or a non-ASCII character",
                            "/aif/x.jsp?attribute=Lev%C3%B8r&name=Leverør", None)
    assert "name" in kind and "Lever" not in kind


def test_a_refused_drill_query_is_named_by_its_first_keyword_only():
    body = json.dumps({"QUERY": "WITH x AS (SELECT secret_col FROM t) SELECT * FROM x", "PARAMS": [], "PARAMSTYPE": []})
    kind = cmp.refused_kind("refused: the drill query is not one read-only SELECT",
                            "/aif/ui/model/drilldown/getData?limit=5000&offset=0", body)
    assert "starts with WITH" in kind and "secret_col" not in kind and "SELECT" not in kind.split("starts")[-1]


def test_an_error_text_is_scrubbed_of_amounts_and_queries():
    text = cmp.error_kind(RuntimeError("Lines differ by 1 234 567,89 NOK at https://pod/x?sso_token=abc&q=1"))
    assert text.startswith("RuntimeError:") and "567" not in text and "abc" not in text and "q=1" not in text
