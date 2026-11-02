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
