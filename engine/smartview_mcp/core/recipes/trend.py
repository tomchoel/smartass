"""Recipe monthly_trend (B07): one entity's accounts per month up to the closed month, the YTD column checked
against FinMax Accumulated, and the same months last year (unless compare=[]). Wraps templates.monthly_trend
(pulls + layout) on stub pulls."""
from __future__ import annotations

from ... import templates
from smartass.epm.grid import Axis
from . import Draft, RecipeContext, pull_dict, relabel, stub_pulls

DEFAULTS = {"compare": (lambda v: "ly" in (v or []), "the standard trend shows the same months last year")}
FOLLOWUPS = ["P&L for {entity}, YTD {period}",
             "Cost base per cost centre under {entity}, YTD {period}",
             "Actual vs budget for {entity}, YTD {period}"]


def monthly_trend(params: dict, ctx: RecipeContext) -> Draft:
    through = params.get("through") or ctx.closed
    conn = params.get("connection", "FinMax")
    plan = templates.monthly_trend(None, params.get("entity", "un.SY"), params.get("accounts"), through, conn)
    pulls = list(plan.pulls)
    lay = plan.answer(stub_pulls([p.id for p in pulls]))
    lay.pop("key_points", None)                    # value-based; auto_points drafts them after the pull
    with_ly = params.get("compare") is None or "ly" in params["compare"]
    if with_ly:
        tr = pulls[0]
        months = list(tr.cols[0].members)
        last = [templates.same_period_ly(m) for m in months]
        pulls.append(templates.PullSpec("ly", conn, tr.rows, [Axis("All Time Periods", last)], tr.pov))
        lay["sections"].append({"title": "Same months last year", "lines": [
            {"key": f"ly{i}", "label": f"{a} — last year", "kind": "line",
             "cells": [{"pull": "ly", "row": [a], "col": [m]} for m in last] + [{"rowsum": list(range(len(last)))}]}
            for i, a in enumerate(tr.rows[0].members)]})
    for sec in lay["sections"]:
        relabel(sec["lines"], conn, ctx)
    for ch in lay["checks"]:
        head, _, code = ch["label"].rpartition(": ")
        if code.startswith("ac."):
            ch["label"] = f"{head}: {ctx.label(conn, code)}"
    lay["chart"]["title"] = f"{lay['sections'][0]['lines'][-1]['label']} per month"
    return Draft([pull_dict(p) for p in pulls], lay,
                 assumptions=[f"Actual, NOK, Periodic months Jan–{through}" +
                              (", and the same months last year" if with_ly else "")], result={"period": through})
