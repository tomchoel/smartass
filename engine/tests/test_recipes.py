"""Phase 5 recipes: pure functions (params, context) -> draft answer spec v1."""
from types import SimpleNamespace

import pytest

from smartview_mcp.core import contract, recipes
from smartview_mcp.core.recipes import Draft, RecipeContext
from smartview_mcp.grid import Axis
from smartview_mcp.templates import PullSpec

GATE = {"outcome": "clear", "checklist": {"period": "closed month", "scenario": "Actual", "entity": "un.SY",
                                          "measure": "C30", "currency_sign": "NOK", "granularity": "month"}}


def ctx(**over):
    base = dict(closed="Oct 26", cat=None, label=lambda conn, name: f"<{name}>", pnl_facts=None,
                now="2026-11-02 10:00")
    return RecipeContext(**{**base, **over})


def _gate_names_no_recipe(monkeypatch):
    """sv_answer re-runs the gate on the question; question 'q' is answered by a test recipe, so the gate is clear
    and names no standard recipe."""
    from smartview_mcp import tools_core
    monkeypatch.setattr(tools_core.scope, "gate", lambda q, closed, answers=None: {"suggested_outcome": "clear",
                                                                                   "recipe": None})


def _echo(params, c):
    return Draft(pulls=[{"id": "x", "connection": "FinMax", "rows": [{"dimension": "All Accounts", "members": ["ac.C30"]}],
                         "columns": [{"dimension": "All Time Periods", "members": [c.closed]}]}],
                 layout={"title": f"Echo {params['n']}", "columns": [{"header": c.closed}],
                         "sections": [{"title": "S", "lines": [{"key": "a", "label": "A",
                                                                 "cells": [{"pull": "x", "row": ["ac.C30"], "col": [c.closed]}]}]}],
                         "checks": [{"label": "a is a", "expr": "a-a"}]},
                 assumptions=["NOK"], result={"period": c.closed})


def test_expand_draft_applies_overrides_and_keeps_the_recipe(monkeypatch):
    monkeypatch.setitem(recipes.RECIPES, "echo", _echo)
    raw = {"version": 1, "question": "q", "gate": GATE,
           "recipe": {"name": "echo", "params": {"n": 1}, "overrides": {"layout": {"title": "Mine"}}}}
    out, extra = recipes.expand(raw, ctx())
    assert out["layout"]["title"] == "Mine" and out["pulls"][0]["id"] == "x"
    assert out["assumptions"] == ["NOK"] and out["recipe"]["name"] == "echo"
    assert extra == {"period": "Oct 26", "warnings": []}
    contract.validate(out)


def test_a_spec_with_a_layout_is_not_expanded():
    raw = {"version": 1, "question": "q", "gate": GATE, "layout": {"title": "t"}}
    assert recipes.expand(raw, ctx()) == (raw, {})


def test_unknown_recipe_names_the_known_ones():
    raw = {"version": 1, "question": "q", "gate": GATE, "recipe": {"name": "nope"}}
    with pytest.raises(ValueError, match="unknown recipe 'nope'; known: .*monthly_trend"):
        recipes.expand(raw, ctx())


def test_pull_dict_converts_a_template_pullspec():
    ps = PullSpec("tr", "FinMax", [Axis("All Accounts", ["ac.C30"])], [Axis("All Time Periods", ["Oct 26"])],
                  {"All Entities": "un.SY"})
    assert recipes.pull_dict(ps) == {"id": "tr", "connection": "FinMax",
                                     "rows": [{"dimension": "All Accounts", "members": ["ac.C30"]}],
                                     "columns": [{"dimension": "All Time Periods", "members": ["Oct 26"]}],
                                     "pov": {"All Entities": "un.SY"}}


@pytest.mark.anyio
async def test_sv_answer_expands_a_recipe_in_the_excel_job_and_returns_its_period(monkeypatch, tmp_path):
    from smartview_mcp import config, tools_core
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    monkeypatch.setitem(recipes.RECIPES, "echo", _echo)
    t = SimpleNamespace(path="C:/x/a.xlsx", created=False)
    monkeypatch.setattr(tools_core, "_call_user", lambda job, what: _job(job))
    monkeypatch.setattr(tools_core.templates, "closed_month", lambda h, *a: "Oct 26")
    monkeypatch.setattr(tools_core.catalog, "load", lambda: None)
    _gate_names_no_recipe(monkeypatch)
    seen = {}

    def run(h, raw, target, allow, **kw):
        seen["raw"] = raw
        return tools_core.pipeline.Run(t, 1, contract.validate(raw), {}, {"warnings": []}, [],
                                       {"ok": True, "findings": []}, [], raw=raw)
    monkeypatch.setattr(tools_core.pipeline, "run", run)
    monkeypatch.setattr(tools_core.runlog, "write_sheet", lambda t, e: None)
    monkeypatch.setattr(tools_core.reply, "finish", lambda *a, **k: None)
    monkeypatch.setattr(tools_core.reply, "result", lambda *a, **k: {"sheet": "Q1 Answer", "warnings": ["w"]})
    out = await tools_core.sv_answer({"version": 1, "question": "q", "gate": GATE,
                                      "recipe": {"name": "echo", "params": {"n": 2}}})
    assert seen["raw"]["layout"]["title"] == "Echo 2" and seen["raw"]["recipe"]["name"] == "echo"
    assert out["period"] == "Oct 26" and out["warnings"] == ["w"] and out["outline"]["recipe"] == "echo"
    assert out["followups"] == []


@pytest.mark.anyio
async def test_sv_answer_turns_a_recipe_error_into_a_tool_error(monkeypatch):
    from mcp.server.mcpserver.exceptions import ToolError

    from smartview_mcp import tools_core
    monkeypatch.setattr(tools_core, "_call_user", lambda job, what: _job(job))
    monkeypatch.setattr(tools_core.templates, "closed_month", lambda h, *a: "Oct 26")
    monkeypatch.setattr(tools_core.catalog, "load", lambda: None)
    _gate_names_no_recipe(monkeypatch)
    with pytest.raises(ToolError, match="unknown recipe 'nope'"):
        await tools_core.sv_answer({"version": 1, "question": "q", "gate": GATE, "recipe": {"name": "nope"}})


async def _job(job):
    return job(None)


# ------------------------------------------------------------------ specs: data-driven top-N repeat
def _fake_pull(rows: dict[str, float], col=("ac.C70",), labels=None):
    """A Pull-like object: rows member -> value in one column."""
    names = list(rows)
    return SimpleNamespace(sheet="Q1 FinMax cb", layout=SimpleNamespace(col_tuples=[tuple(col)], nr=1),
                           parsed=SimpleNamespace(rows=[[m, rows[m]] for m in names]),
                           row_labels={(m,): (labels or {}).get(m, (f"{m} Name (un)",)) for m in names},
                           col_labels={tuple(col): tuple(col)})


def test_a_top_repeat_ranks_members_by_absolute_value_among_the_given_set():
    from smartview_mcp import specs
    p = _fake_pull({"A": -50.0, "B": 10.0, "C": -300.0, "P": -360.0})
    spec = {"sections": [{"title": "Top", "repeat": {
        "over": {"top": 2, "pull": "cb", "by": ["ac.C70"], "among": ["A", "B", "C"], "abs": True},
        "line": {"key": "c{i}", "label": "{label}", "cells": [{"pull": "cb", "row": ["{m}"], "col": ["ac.C70"]}]}},
        "lines": [{"key": "tot", "label": "Total", "cells": [{"pull": "cb", "row": ["P"], "col": ["ac.C70"]}]}]}]}
    out = specs.expand(spec, {"cb": "Q1 FinMax cb"}, {"Q1 FinMax cb": p})
    lines = out["sections"][0]["lines"]
    assert [ln["key"] for ln in lines] == ["c0", "c1", "tot"]
    assert [ln["label"] for ln in lines[:2]] == ["C Name", "A Name"]
    assert lines[0]["cells"][0] == {"pull": "Q1 FinMax cb", "row": ["C"], "col": ["ac.C70"]}


def test_a_top_repeat_without_the_pull_yields_nothing():
    from smartview_mcp import specs
    spec = {"sections": [{"title": "Top", "repeat": {"over": {"top": 2, "pull": "cb", "by": ["x"]},
                                                     "line": {"key": "c{i}", "cells": []}}, "lines": []}]}
    assert specs.expand(spec, {}, {})["sections"][0]["lines"] == []


# ------------------------------------------------------------------ monthly_trend
def test_monthly_trend_is_the_template_on_pull_ids_with_a_recon_check(monkeypatch):
    from smartview_mcp import templates
    monkeypatch.setattr(templates, "member_label", lambda c, n: f"<{n}>")
    raw = {"version": 1, "question": "C30 per month for Telenor Norge", "gate": GATE,
           "recipe": {"name": "monthly_trend", "params": {"entity": "un.SY", "accounts": ["ac.C30"]}}}
    out, extra = recipes.expand(raw, ctx())
    s = contract.validate(out)
    assert [p.id for p in s.pulls] == ["tr", "ytd", "ly"]
    assert s.pulls[0].columns[0].members == ["Jan 26", "Feb 26", "Mar 26", "Apr 26", "May 26", "Jun 26",
                                             "Jul 26", "Aug 26", "Sep 26", "Oct 26"]
    line = out["layout"]["sections"][0]["lines"][0]
    assert line["cells"][0] == {"pull": "tr", "row": ["ac.C30"], "col": ["Jan 26"]} and line["label"] == "<ac.C30>"
    assert out["layout"]["sections"][1]["lines"][0]["label"] == "<ac.C30> — Accumulated Oct 26"
    assert out["layout"]["checks"][0] == {"label": "Sum of months = YTD: <ac.C30>", "expr": "a0-y0", "columns": [10]}
    assert out["layout"]["chart"]["title"] == "<ac.C30> per month"
    assert "key_points" not in out["layout"] and extra["period"] == "Oct 26"


# ------------------------------------------------------------------ cost_base
def test_cost_base_ranks_after_the_pull_and_checks_top_plus_other_equals_total(monkeypatch):
    from smartview_mcp import templates
    monkeypatch.setattr(templates, "cost_base_members", lambda conn, parent: (["un.SY.A", "un.SY.B", "un.SY"],
                                                                             {"un.SY.A", "un.SY.B"}))
    raw = {"version": 1, "question": "Cost base per cost centre", "gate": GATE,
           "recipe": {"name": "cost_base", "params": {"top": 1}}}
    out, extra = recipes.expand(raw, ctx())
    contract.validate(out)
    sec = out["layout"]["sections"][0]
    assert sec["repeat"]["over"] == {"top": 1, "pull": "cb", "by": ["ac.C70"], "among": ["un.SY.A", "un.SY.B"],
                                     "abs": True}
    assert [ln["key"] for ln in sec["lines"]] == ["topsum", "rest", "tot"]
    assert out["layout"]["checks"] == [{"label": "Top + other = total", "expr": "topsum+rest-tot",
                                        "columns": [0, 1, 2, 3, 4]}]
    assert out["layout"]["title"] == "Cost base per cost centre — <un.SY> — YTD Oct 26"
    assert out["pulls"][0]["pov"]["All Time Periods"] == "Oct 26" and extra["period"] == "Oct 26"


def test_the_cost_base_draft_expands_to_the_ranked_lines_after_the_pull(monkeypatch):
    from smartview_mcp import specs, templates
    monkeypatch.setattr(templates, "cost_base_members", lambda conn, parent: (["A", "B", "un.SY"], {"A", "B"}))
    out, _ = recipes.expand({"version": 1, "question": "q", "gate": GATE,
                             "recipe": {"name": "cost_base", "params": {"top": 1, "accounts": ["ac.C70"]}}}, ctx())
    p = _fake_pull({"A": -50.0, "B": -80.0, "un.SY": -130.0})
    s = specs.expand(out["layout"], {"cb": p.sheet}, {p.sheet: p})
    lines = s["sections"][0]["lines"]
    assert [ln["key"] for ln in lines] == ["c0", "topsum", "rest", "tot"] and lines[0]["label"] == "B Name"
    assert lines[1]["cells"][0] == {"sum": ["c0"]}
    v = specs.evaluate(s, {p.sheet: p})
    assert v["topsum"][0] == -80.0 and v["rest"][0] == -50.0


def test_cost_base_defaults_top10_other_share_and_ly(monkeypatch):
    from smartview_mcp import specs, templates
    leaves = [f"L{i:02d}" for i in range(12)]
    monkeypatch.setattr(templates, "cost_base_members", lambda conn, parent: (leaves + ["un.SY"], set(leaves)))
    out, _ = recipes.expand({"version": 1, "question": "q", "gate": GATE,
                             "recipe": {"name": "cost_base", "params": {"accounts": ["ac.C70"]}}}, ctx())
    contract.validate(out)
    lay = out["layout"]
    assert lay["sections"][0]["repeat"]["over"]["top"] == 10
    assert [c["header"] for c in lay["columns"]] == ["<ac.C70>", "<ac.C70> Oct 25", "Share %"]
    assert lay["columns"][2]["kind"] == "prosent1"
    assert {p["id"]: p["pov"]["All Time Periods"] for p in out["pulls"]} == {"cb": "Oct 26", "ly": "Oct 25"}
    cur = _fake_pull({**{m: -10.0 * (i + 1) for i, m in enumerate(leaves)}, "un.SY": -780.0})
    ly = _fake_pull({**{m: -5.0 for m in leaves}, "un.SY": -60.0})
    ly.sheet = "Q1 FinMax ly"
    pulls = {cur.sheet: cur, ly.sheet: ly}
    s = specs.expand(lay, {"cb": cur.sheet, "ly": ly.sheet}, pulls)
    lines = s["sections"][0]["lines"]
    assert [ln["key"] for ln in lines] == [f"c{i}" for i in range(10)] + ["topsum", "rest", "tot"]
    assert lines[11]["label"] == "Other cost centres"
    v = specs.evaluate(s, pulls)
    assert v["c0"] == [-120.0, -5.0, pytest.approx(120 / 780)]
    assert v["rest"] == [-30.0, -10.0, pytest.approx(30 / 780)] and v["tot"][2] == 1.0
    assert lay["checks"][0]["columns"] == [0, 1]


def test_departures_flags_switched_off_defaults():
    assert recipes.departures("cost_base", {"top": 25}) == [
        "cost_base top=25 (the standard ranking is the top 10 + Other)"]
    assert recipes.departures("cost_base", {"top": 10, "parent": "un.SY"}) == []
    assert recipes.departures("monthly_trend", {"compare": []})
    assert recipes.departures("monthly_trend", {"compare": ["ly"]}) == []
    assert recipes.departures("pnl", {"compare": ["budget"]})
    assert recipes.departures("pnl", {"compare": ["ly", "budget"]}) == []
    assert recipes.departures("concept", {"compare": ""})
    assert recipes.departures(None, {"top": 25}) == []


def test_trend_has_ly_series_by_default(monkeypatch):
    from smartview_mcp import templates
    monkeypatch.setattr(templates, "member_label", lambda c, n: f"<{n}>")
    raw = {"version": 1, "question": "C30 per month", "gate": GATE,
           "recipe": {"name": "monthly_trend", "params": {"entity": "un.SY", "accounts": ["ac.C30"]}}}
    out, _ = recipes.expand(raw, ctx())
    contract.validate(out)
    ly = next(p for p in out["pulls"] if p["id"] == "ly")
    assert ly["columns"][0]["members"] == [f"{m} 25" for m in ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul",
                                                               "Aug", "Sep", "Oct")]
    assert ly["pov"]["All Balance Types"] == "Periodic"
    sec = out["layout"]["sections"][-1]
    assert sec["title"] == "Same months last year" and sec["lines"][0]["label"] == "<ac.C30> — last year"
    assert sec["lines"][0]["cells"][0] == {"pull": "ly", "row": ["ac.C30"], "col": ["Jan 25"]}
    assert sec["lines"][0]["cells"][10] == {"rowsum": list(range(10))}
    raw["recipe"]["params"]["compare"] = []
    out, _ = recipes.expand(raw, ctx())
    assert [p["id"] for p in out["pulls"]] == ["tr", "ytd"]


# ------------------------------------------------------------------ actual_vs_plan
def test_actual_vs_plan_uses_the_given_closed_month_and_checks_the_finmax_recon(monkeypatch):
    from smartview_mcp import templates
    monkeypatch.setattr(templates, "closed_month", lambda *a, **k: pytest.fail("must use ctx.closed"))
    monkeypatch.setattr(templates, "latest_forecast", lambda yy: "Forecast Q326")
    monkeypatch.setattr(templates, "map_finmax_entity", lambda e, x=None: ("un.SY.2TEC", None))
    raw = {"version": 1, "question": "Technology actual vs plan", "gate": GATE,
           "recipe": {"name": "actual_vs_plan", "params": {"entity": "2TEC Technology"}}}
    out, extra = recipes.expand(raw, ctx())
    contract.validate(out)
    assert [p["id"] for p in out["pulls"]] == ["act", "plan", "fm"]
    assert out["pulls"][0]["pov"]["Period"] == "Oct" and out["pulls"][0]["pov"]["Years"] == "FY26"
    assert out["layout"]["sections"][0]["lines"][0]["cells"][0]["pull"] == "act"
    labels = [c["label"] for c in out["layout"]["checks"]]
    assert "EBITDA FinPlan = FinMax" in labels and len(labels) >= 2
    assert out["layout"]["checks"][labels.index("EBITDA FinPlan = FinMax")]["severity"] == "Warning"
    assert "key_points" not in out["layout"]
    assert extra == {"period": "Oct 26", "warnings": []}


def test_actual_vs_plan_without_a_finmax_entity_has_no_recon_check_and_says_why(monkeypatch):
    from smartview_mcp import templates
    monkeypatch.setattr(templates, "latest_forecast", lambda yy: None)
    monkeypatch.setattr(templates, "map_finmax_entity", lambda e, x=None: (None, "No FinMax entity matches X."))
    raw = {"version": 1, "question": "X vs plan", "gate": GATE,
           "recipe": {"name": "actual_vs_plan", "params": {"entity": "X", "month": "Dec", "year": "FY26"}}}
    out, extra = recipes.expand(raw, ctx())
    contract.validate(out)
    assert "EBITDA FinPlan = FinMax" not in [c["label"] for c in out["layout"]["checks"]]
    assert extra["warnings"] == ["No FinMax entity matches X."] and extra["period"] == "Dec 26"


# ------------------------------------------------------------------ concept
def test_concept_wraps_concepts_plan_in_english(monkeypatch):
    from smartview_mcp import concepts
    monkeypatch.setattr(concepts, "_label", lambda m: m)
    c = concepts.get("splitt")
    raw = {"version": 1, "question": "What do we earn on Splitt?", "gate": GATE,
           "recipe": {"name": "concept", "params": {"concept": c["id"]}}}
    out, extra = recipes.expand(raw, ctx())
    contract.validate(out)
    hdr = [col["header"] for col in out["layout"]["columns"]]
    assert hdr == ["YTD Oct 26", "YTD Oct 25", "Change", "Change %"]
    assert all("title" not in p for p in out["pulls"])                       # None titles are dropped
    assert out["layout"]["source"].startswith("Source:") and extra == {"period": "Oct 26", "concept": "splitt",
                                                                       "warnings": []}
    titles = [s["title"] for s in out["layout"]["sections"]]
    assert "Allocation keys and checks" in titles and "Basis from FinMax (member · account)" in titles
    assert out["layout"]["checks"] and "YTD Oct 26 vs Oct 25" in out["layout"]["title"]


def test_concept_unknown_id_lists_the_known_ones(monkeypatch):
    from smartview_mcp import concepts
    monkeypatch.setattr(concepts, "find", lambda text: [])
    raw = {"version": 1, "question": "q", "gate": GATE, "recipe": {"name": "concept", "params": {"concept": "nope"}}}
    with pytest.raises(ValueError, match="Unknown concept 'nope'"):
        recipes.expand(raw, ctx())


def test_pnl_is_fin_spec_on_pull_ids_with_the_live_read_expectation(tmp_path):
    from fin_util import CAT0
    from test_fin_spec import build as fs_build                # (res, req, sheets) for a filled fact store
    res, req, _ = fs_build(tmp_path)
    facts = lambda area, period, compare: (res, "live read", [], [], res.scope.entity)  # noqa: E731
    out, extra = recipes.expand({"version": 1, "question": "Wholesale P&L", "gate": GATE,
                                 "recipe": {"name": "pnl", "params": {"area": "Wholesale"}}},
                                ctx(cat=CAT0, pnl_facts=facts))
    contract.validate(out)
    assert [p["id"] for p in out["pulls"]] == [p["id"] for p in req.pulls]
    assert out["layout"]["expect"]["source"] == req.source
    assert set(out["layout"]["expect"]["cells"]) == set(req.expected)
    assert any(c.get("static") for c in out["layout"]["checks"])           # 'Period is closed'
    assert extra["period"] == res.scope.period


def test_pnl_needs_a_facts_reader_and_an_area():
    raw = {"version": 1, "question": "q", "gate": GATE, "recipe": {"name": "pnl", "params": {"area": "Wholesale"}}}
    with pytest.raises(ValueError, match="facts reader"):
        recipes.expand(raw, ctx())
    raw["recipe"]["params"] = {}
    with pytest.raises(ValueError, match="area"):
        recipes.expand(raw, ctx(pnl_facts=lambda *a: None))


@pytest.mark.parametrize("name", ["monthly_trend", "cost_base", "actual_vs_plan"])
def test_every_data_recipe_has_a_check(name, monkeypatch):
    from smartview_mcp import templates
    monkeypatch.setattr(templates, "member_label", lambda c, n: n)
    monkeypatch.setattr(templates, "cost_base_members", lambda c, p: (["a", p], {"a"}))
    monkeypatch.setattr(templates, "latest_forecast", lambda yy: None)
    monkeypatch.setattr(templates, "map_finmax_entity", lambda e, x=None: ("un.SY", None))
    params = {"monthly_trend": {}, "cost_base": {}, "actual_vs_plan": {"entity": "BANORWAY"}}[name]
    out, _ = recipes.expand({"version": 1, "question": "q", "gate": GATE,
                             "recipe": {"name": name, "params": params}}, ctx())
    assert out["layout"]["checks"]
