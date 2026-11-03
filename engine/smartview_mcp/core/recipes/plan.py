"""Recipe actual_vs_plan (B05, B14): FinPlan Actual (Working) vs Target and the latest forecast (Final) for one
month or YTD, reconciled to FinMax EBITDA. Wraps templates.actual_vs_plan on stub pulls."""
from __future__ import annotations

from ... import templates
from ...host import InputError
from . import Draft, RecipeContext, pull_dict, stub_pulls

_KEYS = ("month", "year", "basis", "forecast", "finmax_entity")
DEFAULTS: dict = {}            # Target and the latest forecast are always shown: no comparison to switch off
FOLLOWUPS = ["Monthly trend for {entity} through {period}",
             "Cost base per cost centre under {entity}, YTD {period}"]


def actual_vs_plan(params: dict, ctx: RecipeContext) -> Draft:
    entity = params.get("entity")
    if not entity:
        raise InputError("recipe actual_vs_plan needs params.entity: a FinPlan entity, e.g. '2TEC Technology'.")
    plan = templates.actual_vs_plan(None, entity, closed=ctx.closed,
                                    **{k: params[k] for k in _KEYS if params.get(k) is not None})
    ids = [p.id for p in plan.pulls]
    lay = plan.answer(stub_pulls(ids))
    lay.pop("key_points", None)                    # value-based; auto_points drafts them after the pull
    warnings = []
    if "fm" in ids:
        # FinPlan and FinMax group accounts differently: a non-zero difference is a real finding (a Warning)
        lay["checks"].append({"label": "EBITDA FinPlan = FinMax", "expr": "fm_diff", "columns": [0],
                              "severity": "Warning"})
    else:
        warnings += [n for n in lay.get("notes", []) if n.startswith("No FinMax entity")]
    act = next(p for p in plan.pulls if p.id == "act").pov
    period = f"{act['Period']} {act['Years'][2:]}"
    basis = "YTD" if params.get("basis", "ytd") == "ytd" else "month"
    return Draft([pull_dict(p) for p in plan.pulls], lay,
                 assumptions=[f"FinPlan Actual (Working) vs Target (Final), {basis} {period}"],
                 warnings=warnings, result={"period": period})
