"""Recipe monthly_trend (B07): one entity's accounts per month up to the closed month, the YTD column checked
against FinMax Accumulated. Wraps templates.monthly_trend (pulls + layout) on stub pulls."""
from __future__ import annotations

from ... import templates
from . import Draft, RecipeContext, pull_dict, relabel, stub_pulls


def monthly_trend(params: dict, ctx: RecipeContext) -> Draft:
    through = params.get("through") or ctx.closed
    conn = params.get("connection", "FinMax")
    plan = templates.monthly_trend(None, params.get("entity", "un.SY"), params.get("accounts"), through, conn)
    lay = plan.answer(stub_pulls([p.id for p in plan.pulls]))
    lay.pop("key_points", None)                    # value-based; auto_points drafts them after the pull
    for sec in lay["sections"]:
        relabel(sec["lines"], conn, ctx)
    for ch in lay["checks"]:
        head, _, code = ch["label"].rpartition(": ")
        if code.startswith("ac."):
            ch["label"] = f"{head}: {ctx.label(conn, code)}"
    lay["chart"]["title"] = f"{lay['sections'][0]['lines'][-1]['label']} per month"
    return Draft([pull_dict(p) for p in plan.pulls], lay,
                 assumptions=[f"Actual, NOK, Periodic months Jan–{through}"], result={"period": through})
