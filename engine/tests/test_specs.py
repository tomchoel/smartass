"""specs.py: repeat/sum shorthands, missing months, Python evaluation, drafted key points, patch merge."""
import sys
from types import SimpleNamespace as NS

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from smartview_mcp import specs  # noqa: E402

MONTHS_25 = [f"{m} 25" for m in specs.MONTHS]
MONTHS_26 = [f"{m} 26" for m in specs.MONTHS[:9]]


def fake_pull(cols, values):
    """One row 'ac.EBITDA' over cols."""
    lay = NS(col_tuples=[(c,) for c in cols], row_tuples=[("ac.EBITDA",)], nr=1)
    parsed = NS(rows=[["ac.EBITDA", *values]])
    return NS(layout=lay, parsed=parsed, row_labels={("ac.EBITDA",): ("EBITDA",)},
              col_labels={(c,): (c,) for c in cols})


def monthly_spec():
    e = ["ac.EBITDA"]
    return {"title": "t", "columns": [{"header": "2026"}, {"header": "2025"}, {"header": "Variance"},
                                      {"header": "%", "kind": "prosent1"}],
            "sections": [{"title": "Per month", "repeat": {"over": "months:Jan..Dec", "line": {
                "key": "m{i}", "label": "{v}",
                "cells": [{"pull": "m", "row": e, "col": ["{v} 26"]}, {"pull": "m", "row": e, "col": ["{v} 25"]},
                          {"var": [0, 1]}, {"varpct": [0, 1]}]}},
                "lines": [{"key": "ytd", "label": "YTD", "kind": "grand",
                           "cells": [{"sum": "m[0-8]"}, {"sum": "m[0-8]"}, {"var": [0, 1]}, {"varpct": [0, 1]}]}]}],
            "chart": {"type": "column", "lines": "m*", "columns": [0, 1]}}


def pulls():
    p = fake_pull(MONTHS_25 + MONTHS_26, [100.0] * 12 + [110.0] * 8 + [90.0])
    return {"Q1 FinMax": p}, {"m": "Q1 FinMax"}


def test_repeat_expands_and_blanks_open_months():
    sess, ids = pulls()
    s = specs.expand(monthly_spec(), ids, sess)
    lines = s["sections"][0]["lines"]
    assert [ln["key"] for ln in lines] == [f"m{i}" for i in range(12)] + ["ytd"]   # months, then the total
    by = {ln["key"]: ln for ln in lines}
    assert by["m0"]["cells"][0]["pull"] == "Q1 FinMax" and by["m0"]["cells"][0]["col"] == ["Jan 26"]
    assert by["m9"]["cells"][0] is None                  # Oct 26 not pulled -> empty
    assert by["m9"]["cells"][1]["col"] == ["Oct 25"]
    assert by["m9"]["cells"][2] is None and by["m9"]["cells"][3] is None   # variance of an empty cell dropped
    assert by["ytd"]["cells"][0]["sum"] == [f"m{i}" for i in range(9)]
    assert s["chart"]["lines"] == [f"m{i}" for i in range(12)]


def test_evaluate_and_draft_points():
    sess, ids = pulls()
    s = specs.expand(monthly_spec(), ids, sess)
    v = specs.evaluate(s, sess)
    assert v["m0"][:3] == [110.0, 100.0, 10.0]
    assert v["m8"][2] == -10.0
    assert v["ytd"][0] == 8 * 110 + 90 and v["ytd"][1] == 900 and v["ytd"][2] == 70
    pts = specs.draft_points(s, v)
    assert pts[0].startswith("YTD: 970 (2026) vs 900 (2025), variance 70, +7,8 %")
    assert "largest decrease Sep −10" in pts[1] and pts[1][0].isupper()


def test_prepare_puts_drafted_points_first():
    sess, ids = pulls()
    spec = {**monthly_spec(), "key_points": ["Insight."]}
    s = specs.prepare(spec, ids, sess)
    assert s["key_points"][-1] == "Insight." and len(s["key_points"]) == 3
    s2 = specs.prepare({**spec, "auto_points": False}, ids, sess)
    assert s2["key_points"] == ["Insight."]


def test_plain_spec_unchanged_and_missing_member_kept_outside_repeat():
    sess, ids = pulls()
    spec = {"columns": [{"header": "a"}], "sections": [{"title": "x", "lines": [
        {"key": "k", "label": "K", "cells": [{"pull": "Q1 FinMax", "row": ["ac.EBITDA"], "col": ["Oct 26"]}]}]}]}
    s = specs.expand(spec, ids, sess)
    assert s["sections"][0]["lines"][0]["cells"][0]["col"] == ["Oct 26"]   # surfaces as a warning in build


def test_merge_patch():
    base = {"title": "a", "notes": ["x"], "chart": {"type": "column", "columns": [0, 1]}}
    m = specs.merge(base, {"notes": ["y"], "chart": {"type": "bar"}, "title": None})
    assert m == {"notes": ["y"], "chart": {"type": "bar", "columns": [0, 1]}}
    assert base["notes"] == ["x"]


def test_avg_pattern_expands_and_evaluates():
    sess, ids = pulls()
    spec = {"columns": [{"header": "FTE"}], "sections": [{"title": "levels", "repeat": {"over": ["Dec 25", "Jan 26"],
            "line": {"key": "lv{i}", "label": "{v}", "cells": [{"pull": "m", "row": ["ac.EBITDA"], "col": ["{v}"]}]}},
            "lines": [{"key": "favg", "label": "Average", "cells": [{"avg": "lv*"}]}]}]}
    s = specs.expand(spec, ids, sess)
    assert s["sections"][0]["lines"][-1]["cells"][0] == {"avg": ["lv0", "lv1"]}
    assert specs.evaluate(s, sess)["favg"] == [(100.0 + 110.0) / 2]


def test_avg_counts_a_missing_level_as_zero_and_an_unknown_key_blanks_it():
    p = fake_pull(["Dec 25", "Jan 26"], ["#Missing", 110.0])
    spec = {"columns": [{"header": "FTE"}], "sections": [{"title": "levels", "lines": [
        {"key": "a", "label": "a", "cells": [{"pull": "m", "row": ["ac.EBITDA"], "col": ["Dec 25"]}]},
        {"key": "b", "label": "b", "cells": [{"pull": "m", "row": ["ac.EBITDA"], "col": ["Jan 26"]}]},
        {"key": "avg", "label": "avg", "cells": [{"avg": ["a", "b"]}]},
        {"key": "bad", "label": "bad", "cells": [{"avg": ["a", "zz"]}]}]}]}
    v = specs.evaluate(spec, {"Q1 FinMax": p} | {"m": p})
    assert v["avg"] == [55.0] and v["bad"] == [None]
