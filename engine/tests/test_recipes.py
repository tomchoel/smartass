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
                  {"All Entities": "un.SY"}, title_suffix="YTD")
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
    seen = {}

    def run(h, raw, target, allow, **kw):
        seen["raw"] = raw
        return tools_core.pipeline.Run(t, 1, contract.validate(raw), {}, {"warnings": []}, [],
                                       {"ok": True, "findings": []}, [], raw=raw)
    monkeypatch.setattr(tools_core.pipeline, "run", run)
    monkeypatch.setattr(tools_core.runlog, "write_sheet", lambda t, e: None)
    monkeypatch.setattr(tools_core.tools_ws, "_finish", lambda *a, **k: None)
    monkeypatch.setattr(tools_core.tools_ws, "_result", lambda *a, **k: {"sheet": "Q1 Answer", "warnings": ["w"]})
    out = await tools_core.sv_answer({"version": 1, "question": "q", "gate": GATE,
                                      "recipe": {"name": "echo", "params": {"n": 2}}})
    assert seen["raw"]["layout"]["title"] == "Echo 2" and seen["raw"]["recipe"]["name"] == "echo"
    assert out["period"] == "Oct 26" and out["warnings"] == ["w"] and out["outline"]["recipe"] == "echo"


@pytest.mark.anyio
async def test_sv_answer_turns_a_recipe_error_into_a_tool_error(monkeypatch):
    from mcp.server.mcpserver.exceptions import ToolError

    from smartview_mcp import tools_core
    monkeypatch.setattr(tools_core, "_call_user", lambda job, what: _job(job))
    monkeypatch.setattr(tools_core.templates, "closed_month", lambda h, *a: "Oct 26")
    monkeypatch.setattr(tools_core.catalog, "load", lambda: None)
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
    assert [p.id for p in s.pulls] == ["tr", "ytd"]
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
                                        "columns": [0, 1, 2, 3]}]
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
    assert lines[1]["cells"] == [{"sum": ["c0"]}]
    v = specs.evaluate(s, {p.sheet: p})
    assert v["topsum"] == [-80.0] and v["rest"] == [-50.0]


def test_templates_cost_base_uses_the_extracted_members(monkeypatch):
    from smartview_mcp import templates
    monkeypatch.setattr(templates, "cost_base_members", lambda conn, parent: (["A", parent], {"A"}))
    plan = templates.cost_base(None, period="Oct 26")
    assert plan.pulls[0].rows[0].members == ["A", "un.SY"]
