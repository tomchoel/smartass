"""scripts/compare_drill_routes.py's judge and its output (Task 7's gate for deleting the browser route): rows
compared as sorted JSON with their types, the tie, the cube value and the columns each fail the comparison, the view
crosses between processes as a digest, the exit-code contract, the renewal probe's verdict, and nothing it prints
carries a figure, a business text, a token or a query value. The live runs themselves are the script's, on the
owner's PC."""
import importlib.util
import json
from decimal import Decimal
from pathlib import Path

import pytest

_p = Path(__file__).resolve().parents[2] / "scripts" / "compare_drill_routes.py"
_spec = importlib.util.spec_from_file_location("compare_drill_routes", _p)
cmp = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cmp)

CELL = {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"}
ROWS = [{"AMOUNTX": 123456.78, "Supplier": "Leverandør AS", "Account": "4010"},
        {"AMOUNTX": -98765.43, "Supplier": "Annen Leverandør", "Account": "4020"}]
VIEW = {"items": [{"viewName": "DrillFilter", "viewDefinition": json.dumps(
    [{"label": "Supplier", "value": "SUPPLIER"}, {"label": "Amount", "value": "AMOUNTX"}])}],
    "links": [{"href": "https://pod/aif/ui/model/view?page=DRILL_THROUGH&appId=52&_=1795272710098"}]}


def side(rows=ROWS, columns=("AMOUNTX", "Supplier", "Account"), status="exact", cube=24691.35, secs=12.34):
    return {"ok": True, "rows": cmp.row_keys(rows), "types": cmp.column_types(rows), "columns": list(columns),
            "status": status, "cell_value": cube, "secs": secs}


def test_the_same_rows_in_another_order_are_the_same_lines():
    c = cmp.compare(side(), side(rows=list(reversed(ROWS)), secs=3.2))
    assert c["same_lines"] and c["tie_equal"] and c["cube_equal"] and c["ok"]


def test_a_changed_amount_or_a_missing_row_is_not_the_same_lines():
    changed = [dict(ROWS[0], AMOUNTX=123456.79), ROWS[1]]
    c = cmp.compare(side(), side(rows=changed))
    assert not c["same_lines"] and (c["only_a"], c["only_b"]) == (1, 1)
    assert not cmp.compare(side(), side(rows=ROWS[:1]))["ok"]


def test_a_row_paged_twice_shows_as_a_duplicate():
    c = cmp.compare(side(), side(rows=[ROWS[0], ROWS[0]]))
    assert not c["same_lines"] and c["dupes_b"] == 1 and c["only_a"] == 1 and c["only_b"] == 1


def test_a_value_of_another_type_is_not_the_same_line_and_is_named_by_its_column():
    assert cmp.row_keys([{"x": Decimal("1.0")}]) != cmp.row_keys([{"x": "1.0"}])
    c = cmp.compare(side(), side(rows=[dict(ROWS[0], Account=4010), ROWS[1]]))
    assert not c["ok"] and c["type_differs"] == ["Account (str/int, str)"]


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
    assert "tie token/token:" in cmp.line(CELL, cmp.compare(side(), side()), ("token", "token"))


def test_a_failed_route_shows_its_error_kind_in_the_line():
    text = cmp.line(CELL, cmp.compare(side(), {"ok": False, "error": "NotSignedIn: sign-in"}))
    assert "token failed: NotSignedIn" in text and "same lines: False" in text


def test_the_view_crosses_as_a_digest_without_values():
    d = cmp.view_digest(VIEW)
    shown = json.dumps(d)
    assert not any(s in shown for s in ("Supplier", "SUPPLIER", "AMOUNTX", "1795272710098", "DRILL_THROUGH", "pod"))
    assert d["columns"] == 2


def test_the_view_diff_names_key_paths_and_types_never_values():
    a = {"columns": [{"name": "Supplier", "width": 120}, {"name": "Amount"}], "ts": 1795272710098, "title": "Kostnad"}
    b = {"columns": [{"name": "Supplier", "width": "120"}, {"name": "Amount"}], "ts": 1795272799999, "extra": True}
    d = cmp.view_diff(cmp.view_digest(a), cmp.view_digest(b))
    assert d["only_browser"] == ["$.title"] and d["only_token"] == ["$.extra"]
    assert d["type_differs"] == ["$.columns[0].width (int/str)"] and d["value_differs"] == ["$.ts"]
    shown = json.dumps(d) + "\n".join(cmp.view_lines(cmp.view_digest(a), cmp.view_digest(b))[0])
    assert not any(s in shown for s in ("Kostnad", "1795272710098", "1795272799999", "Supplier"))


def test_a_link_that_differs_is_shown_by_its_parts_never_its_values():
    b = json.loads(json.dumps(VIEW))
    b["links"][0]["href"] = b["links"][0]["href"].replace("710098", "799999")
    d = cmp.view_diff(cmp.view_digest(VIEW), cmp.view_digest(b))
    assert d["value_differs"] == ["$.links[0].href (a link: same path True, query items that differ: _)"]
    lines, ok = cmp.view_lines(cmp.view_digest(VIEW), cmp.view_digest(b))
    assert ok and "1795272" not in "\n".join(lines)


def test_the_view_lines_say_whether_the_columns_differ():
    lines, ok = cmp.view_lines(cmp.view_digest(VIEW), cmp.view_digest(VIEW))
    assert ok and "same documents: True" in lines[0] and "same labels and fields in the same order: True" in lines[1]
    other = json.loads(json.dumps(VIEW))
    other["items"][0]["viewDefinition"] = json.dumps([{"label": "Amount", "value": "AMOUNTX"},
                                                      {"label": "Supplier", "value": "SUPPLIER"}])
    lines, ok = cmp.view_lines(cmp.view_digest(VIEW), cmp.view_digest(other))
    assert not ok and "same labels and fields in the same order: False" in lines[1]


def test_views_the_drill_cannot_read_are_a_failure():
    lines, ok = cmp.view_lines(cmp.view_digest({"a": 1}), cmp.view_digest({"a": 1}))
    assert not ok and "not a drill view" in lines[1]


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


def test_an_error_is_its_class_and_a_fixed_kind_never_its_text():
    from smartass.epm.drillcore import CellTooWide, DrillError, SignedOut
    assert cmp.error_kind(RuntimeError("Lines differ by 1 234 567,89 NOK for Leverandør AS")) == \
        "RuntimeError: other"
    assert cmp.error_kind(CellTooWide("EPMFDM-140373: too wide for Leverandør")) == \
        "CellTooWide: EPM refused (EPMFDM-140373)"
    assert cmp.error_kind(DrillError("EPM refused the drill: Leverandør")) == "DrillError: EPM refused"
    assert cmp.error_kind(SignedOut("x")) == "SignedOut: the EPM session was lost"
    assert cmp.error_kind(TimeoutError()) == "TimeoutError: timeout"


def _got(a=None, b=None, started=False):
    return [{"route": "browser", "cells": {"1": a or {**side(), "view": cmp.view_digest(VIEW)}},
             "browser_started": True},
            {"route": "token", "cells": {"1": b or {**side(), "view": cmp.view_digest(VIEW)}},
             "browser_started": started}]


def test_report_exits_0_only_when_every_cell_and_the_view_are_the_same(capsys):
    assert cmp.report([1], _got()) == 0
    assert cmp.report([1], _got(b={**side(status="gap"), "view": cmp.view_digest(VIEW)})) == 1
    assert cmp.report([1], _got(b={"ok": False, "error": "DrillError: other"})) == 1
    assert cmp.report([1], _got(started=True)) == 1
    assert cmp.report([1], _got(b={**side(), "view": cmp.view_digest({"a": 1})})) == 1
    assert "PASS" in capsys.readouterr().out


def test_main_is_blocked_without_a_sign_in(monkeypatch, capsys):
    from smartass.epm import oauth

    def no(base, reject=None):
        raise oauth.NotSignedIn("not signed in")
    monkeypatch.setattr(oauth, "access_token", no)
    monkeypatch.setattr(cmp.subprocess, "run", lambda *a, **k: pytest.fail("no route may run"))
    assert cmp.main(["--cell", "1"]) == cmp.BLOCKED
    assert "BLOCKED" in capsys.readouterr().out


@pytest.mark.parametrize("token_cell, code", [
    ({**side(), "view": cmp.view_digest(VIEW)}, 0),
    ({**side(rows=ROWS[:1]), "view": cmp.view_digest(VIEW)}, 1),
    ({"ok": False, "error": "NotSignedIn: sign-in", "blocked": True}, cmp.BLOCKED),
])
def test_main_runs_one_process_per_route_and_keeps_the_exit_contract(monkeypatch, token_cell, code):
    from smartass.epm import oauth
    monkeypatch.setattr(oauth, "access_token", lambda base, reject=None: "at")
    seen = []

    class Done:
        returncode, stderr = 0, ""

    def run(argv, env, **kw):
        route = argv[argv.index("--route") + 1]
        seen.append((route, env["SV_DRILL_ROUTE"]))
        cell = {**side(), "view": cmp.view_digest(VIEW)} if route == "browser" else token_cell
        Path(argv[argv.index("--out") + 1]).write_text(json.dumps(
            {"route": route, "cells": {"1": cell}, "browser_started": route == "browser"}), encoding="utf-8")
        return Done()
    monkeypatch.setattr(cmp.subprocess, "run", run)
    assert cmp.main(["--cell", "1"]) == code
    assert seen == [("browser", "browser"), ("token", "token")]


def _probe(**over):
    p = {"ok": True, "renewals": 2, "same_rows": True, "tie": "exact", "cube_equal": True, "old_after_new_key": 0,
         "later_ok": True, "later_same_rows": True}
    return {**p, **over}


@pytest.mark.parametrize("over, ok", [
    ({}, True), ({"renewals": 0}, False), ({"same_rows": False}, False), ({"old_after_new_key": 1}, False),
    ({"later_ok": False}, False), ({"later_same_rows": False}, False), ({"tie": "gap"}, False),
    ({"ok": False}, False),
])
def test_the_renewal_probe_passes_only_with_a_renewal_the_same_lines_and_a_working_later_drill(over, ok):
    assert cmp.renewal_passed(_probe(**over)) is ok


def test_old_key_requests_are_sorted_by_when_they_left():
    reqs = [{"gen": 0, "t": 1.0}, {"gen": 0, "t": 2.5}, {"gen": 0, "t": 4.0}, {"gen": 1, "t": 5.0},
            {"gen": None, "t": 3}]
    assert cmp.old_key_timing(reqs, old=0, renew_at=2.0, new_key_at=3.0) == {"before": 1, "between": 1, "after": 1}
