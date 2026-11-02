"""Recipe lines_ranking (plan 2026-11-02 Batch B): invoiced ranking top 10 + Other, the group block, the reconciling
line, LY from a second drill, the unexplained gap. Pure: fixture rows mirror the live consultant question."""
from __future__ import annotations

from types import SimpleNamespace

import pytest

from smartview_mcp import catalog, drill_xl, specs
from smartview_mcp.core import contract, recipes
from smartview_mcp.core.recipes import lines

M = 1e6


def _r(supplier, amount, period, **kw):
    return {"AP_CUSTOMER_NAME": supplier, "AMOUNT": float(amount), "Period": period, **kw}


VENDORS = [(f"Vendor {i:02d} AS", -(1.2 - i * 0.1) * M) for i in range(1, 12)]     # 11 small suppliers
D1 = ([_r("KPMG AS", -10 * M, "Jan-26"), _r("KPMG AS", -20 * M, "Sep-26"),
       _r("Accenture Norge AS", -20 * M, "Mar-26"), _r("Sopra Steria AS", -15 * M, "Jun-26"),
       _r("TSS - Telenor Shared Services AS", -121.4 * M, "May-26"), _r("Telenor ASA", -15.2 * M, "Apr-26"),
       _r("Towers Norge AS", -2.9 * M, "Feb-26"),
       _r("", 25.4 * M, "Jun-26"), _r(None, 0.0, "Jul-26"),
       _r("", 12.7 * M, "Oct-26"), _r("KPMG AS", -99 * M, "Oct-26")]          # October: after the closed month
      + [_r(n, a, "Aug-26") for n, a in VENDORS])
D2 = [_r("KPMG AS", -25 * M, "Sep-25"), _r("Accenture Norge AS", -18 * M, "Mar-25"), _r("Telenor ASA", -10 * M, "Apr-25"),
      _r("", 5 * M, "Jun-25"), _r("Sopra Steria AS", -3 * M, "Dec-25")]                 # December: after Sep
PARAMS = {"accounts": ["ac.C60080", "ac.C60090"], "entity": "un.SY", "period": "YTD Sep 26", "by": "supplier",
          "projects": False}


def _ctx(lines_=None, closed="Sep 26"):
    return recipes.RecipeContext(closed=closed, cat=catalog.load(), label=lambda c, m: m, now="2026-10-05 10:00",
                                 lines={"d1": D1, "d2": D2} if lines_ is None else lines_)


def _ext_total(rows, months):
    return sum(r["AMOUNT"] for r in rows if r["Period"] in months and r["AP_CUSTOMER_NAME"])


def _cube(now: float, ly: float):
    """A fake cube pull: two accounts (rows) × [Sep 26, Sep 25] (Accumulated = YTD)."""
    rows = [["ac.C60080", now * 0.75, ly * 0.75], ["ac.C60090", now * 0.25, ly * 0.25]]
    return {"cube": SimpleNamespace(layout=SimpleNamespace(col_tuples=[("Sep 26",), ("Sep 25",)], nr=1),
                                    parsed=SimpleNamespace(rows=rows))}


def _values(d, cube_now, cube_ly, lines_=None):
    return specs.evaluate(d.layout, _cube(cube_now, cube_ly), lines_ if lines_ is not None else {"d1": D1, "d2": D2})


def _lines_total(rows, months):
    return sum(r["AMOUNT"] for r in rows if r["Period"] in months)


NOW = [f"{m}-26" for m in ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep")]
LY = [m.replace("-26", "-25") for m in NOW]


def _labels(d, section: str) -> list[str]:
    sec = next(s for s in d.layout["sections"] if s["title"].startswith(section))
    return [ln["label"] for ln in sec["lines"]]


def test_ranking_top10_excludes_group_and_orders_by_cost():
    d = lines.lines_ranking(dict(PARAMS), _ctx())
    ranked = _labels(d, "Invoiced by supplier")[:10]
    assert ranked[:3] == ["KPMG AS", "Accenture Norge AS", "Sopra Steria AS"]
    assert ranked[3:] == [n for n, _ in VENDORS[:7]]                       # biggest cost first
    assert not any("Telenor" in x or "TSS" in x or "Towers" in x for x in ranked)
    contract.validate({"version": 1, "question": "q", "gate": {"outcome": "clear"}, "pulls": d.pulls,
                       "layout": d.layout})


def test_ranking_has_other_external_group_and_reconciling_lines():
    d = lines.lines_ranking(dict(PARAMS), _ctx())
    labs = _labels(d, "Invoiced by supplier")
    assert labs[10:] == ["Other external suppliers", "Total external suppliers"]
    assert _labels(d, "Group companies") == ["TSS - Telenor Shared Services AS", "Telenor ASA", "Towers Norge AS",
                                             "Total group companies"]
    assert _labels(d, "Reconciliation") == [lines.NO_SUPPLIER, "Total (ledger lines)", lines.GAP, "ac.C60080",
                                            "ac.C60090", "Total (cube)"]
    total = _lines_total(D1, NOW)
    v = _values(d, total, _lines_total(D2, LY))
    assert v["nosupp"][0] == pytest.approx(25.4 * M) and v["lines"][0] == pytest.approx(total)
    assert v["grp"][0] == pytest.approx(-139.5 * M)
    assert v["other"][0] == pytest.approx(sum(a for _, a in VENDORS[7:]))
    assert v["s1"][1] == pytest.approx(-25 * M) and v["s1"][4] == pytest.approx(-30 * M / total)


def test_ranking_checks_tie_blocks_to_lines_and_lines_to_cube():
    d = lines.lines_ranking(dict(PARAMS), _ctx())
    labels = {c["label"]: c for c in d.layout["checks"]}
    assert set(labels) == {"Blocks = lines total", "Total = cube"}
    v = _values(d, -170 * M, -60 * M)
    for c in d.layout["checks"]:
        res = specs.evaluate({**d.layout, "sections": d.layout["sections"] + [
            {"title": "chk", "lines": [{"key": "chk", "cells": [{"expr": c["expr"]}] * 2}]}]},
            _cube(-170 * M, -60 * M), {"d1": D1, "d2": D2})["chk"]
        assert all(abs(x) < 1 for x in res[:2]), (c["label"], res)
    assert v["cube"][0] == pytest.approx(-170 * M)


def test_ranking_unexplained_gap_is_a_reconciling_line_and_a_warning_finding():
    d = lines.lines_ranking(dict(PARAMS), _ctx())
    total = _lines_total(D1, NOW)
    v = _values(d, total - 11.2 * M, _lines_total(D2, LY))
    assert v["gap"][0] == pytest.approx(-11.2 * M)
    f = recipes.value_findings(d.result["findings"], v)
    assert len(f) == 1 and f[0]["severity"] == "Warning" and "-11,200,000" in f[0]["text"]
    assert "Unexplained gap" in f[0]["text"] and f"{len([r for r in D1 if r['Period'] in NOW])} lines" in f[0]["text"]


def test_ranking_no_gap_finding_within_1_nok():
    d = lines.lines_ranking(dict(PARAMS), _ctx())
    v = _values(d, _lines_total(D1, NOW) + 0.6, 0.0)
    assert recipes.value_findings(d.result["findings"], v) == []


def test_ranking_ly_cells_use_a_list_where_on_ly_months():
    d = lines.lines_ranking(dict(PARAMS), _ctx())
    s1 = d.layout["sections"][0]["lines"][0]
    assert s1["cells"][0] == {"drill": "d1", "where": {"AP_CUSTOMER_NAME": "KPMG AS", "Period": NOW}}
    assert s1["cells"][1] == {"drill": "d2", "where": {"AP_CUSTOMER_NAME": "KPMG AS", "Period": LY}}
    v = _values(d, -1.0, -1.0)
    assert v["s3"][1] == 0.0                                    # Sopra Dec-25 is outside the LY months
    assert [x["alias"] for x in lines.drills(dict(PARAMS), _ctx())] == ["d1", "d2"]
    assert lines.drills(dict(PARAMS), _ctx())[1]["cells"][0] == {"All Accounts": "ac.C60080",
                                                                 "All Entities": "un.SY", "All Time Periods": "Yr 25"}


def test_ranking_without_ly_drill_falls_back_to_total_ly():
    d = lines.lines_ranking(dict(PARAMS), _ctx({"d1": D1}))
    s1 = d.layout["sections"][0]["lines"][0]
    assert s1["cells"][1] is None and s1["cells"][2] is None
    assert any("last year" in w.lower() for w in d.warnings)
    v = _values(d, -170 * M, -60 * M, {"d1": D1})
    assert v["cube"][1] == pytest.approx(-60 * M) and v["s1"][1] is None
    assert all(c["columns"] == [0] for c in d.layout["checks"])


def test_ranking_assumptions_pass_the_english_check():
    d = lines.lines_ranking(dict(PARAMS), _ctx())
    assert d.assumptions and not [a for a in d.assumptions if contract.non_english(a)]
    assert any("reconciling line" in a for a in d.assumptions)


def test_drill_formula_list_value_is_an_or():
    f = drill_xl.drill_formula("Lines_Q1_d1", {"AP_CUSTOMER_NAME": "KPMG AS", "Period": ["Jan-26", "Feb-26"]})
    assert f == ('=SUM(FILTER(Lines_Q1_d1[AMOUNT],(Lines_Q1_d1[AP_CUSTOMER_NAME]="KPMG AS")*'
                 '(ISNUMBER(MATCH(Lines_Q1_d1[Period],{"Jan-26","Feb-26"},0))),0))')
    with pytest.raises(Exception, match="empty"):
        drill_xl.drill_formula("T", {"Period": []})
    assert specs._drill_value(D1, {"AP_CUSTOMER_NAME": "kpmg as", "Period": ["Jan-26", "Oct-26"]}, "sum") == -109 * M


def test_group_company_patterns_match_telenor_tss_towers_not_kpmg():
    gc = catalog.load().group_companies
    assert gc.column == "" and gc.patterns
    for n in ("Telenor ASA", "TSS - Telenor Shared Services AS", "Telenor Shared Services", "Towers Norge AS"):
        assert lines.is_group(n, [], gc), n
    for n in ("KPMG AS", "Accenture Norge AS", "Statsbygg", "Stssen AS"):
        assert not lines.is_group(n, [], gc), n

def test_span_of_gate_periods():
    s = lines.span
    assert (s(None, "Sep 26").head(26), s("YTD Sep 26", "Oct 26").months(25)[-1]) == ("YTD Sep 26", "Sep-25")
    assert s("FY 26", "Sep 26").head(26) == "YTD Sep 26" and s("2025", "Sep 26").head(25) == "YTD Dec 25"
    q = s("Q3 26", "Oct 26")
    assert (q.head(26), q.opening(26), q.balance) == ("Q3 26", "Jun 26", "Accumulated")
    assert (s("Aug 26", "Sep 26").balance, s("Aug 26", "Sep 26").months(26)) == ("Periodic", ["Aug-26"])


def test_ranking_without_lines_is_a_valid_cube_only_draft_with_followups_filled():
    d = lines.lines_ranking(dict(PARAMS), _ctx({}))
    contract.validate({"version": 1, "question": "q", "gate": {"outcome": "clear"}, "pulls": d.pulls,
                       "layout": d.layout})
    assert d.warnings and d.layout["checks"]
    full = lines.lines_ranking(dict(PARAMS), _ctx())
    f = recipes.followups("lines_ranking", PARAMS, "YTD Sep 26", full.result["fill"])
    assert f[0] == "Show the invoice lines for KPMG AS" and len(f) == 3


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
    d = lines.lines_ranking({**PARAMS, "projects": True}, _ctx())
    titles = [s["title"] for s in d.extras[0]["layout"]["sections"]]
    assert d.extras[0]["topic"] == "By project" and "Initiative (derived, not booked)" in titles


def test_by_project_sheet_sums_to_lines_total():
    rows = [dict(r, PROJECT="P1" if r["AP_CUSTOMER_NAME"] == "KPMG AS" else "000000",
                 PROJECT_DESC="Audit", CC_DESC="Finance") for r in D1]
    d = lines.lines_ranking({**PARAMS, "projects": True}, _ctx({"d1": rows, "d2": D2}))
    contract.validate({"version": 1, "question": "q", "gate": {"outcome": "clear"}, "pulls": d.pulls,
                       "layout": d.layout, "extras": d.extras})
    lay = d.extras[0]["layout"]
    v = specs.evaluate(lay, {}, {"d1": rows})
    assert v["bp"][0] == pytest.approx(-30 * M) and v["bp1"][0] == pytest.approx(-30 * M)
    assert v["bp"][0] + v["dp"][0] == pytest.approx(_lines_total(rows, NOW)) == pytest.approx(v["plines"][0])
    chk = specs.evaluate({**lay, "sections": lay["sections"] + [
        {"title": "c", "lines": [{"key": "chk", "cells": [{"expr": lay["checks"][0]["expr"]}]}]}]}, {}, {"d1": rows})
    assert abs(chk["chk"][0]) < 1 and lay["checks"][0]["label"] == "By project = lines total"
    assert not lines.lines_ranking(dict(PARAMS), _ctx()).extras
