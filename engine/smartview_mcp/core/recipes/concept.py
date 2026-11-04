"""Recipe concept (B09): a business concept from the concept library (knowledge/concepts.json) - booked,
attributed and estimated items, allocation keys and the concept's own checks. Wraps concepts.plan."""
from __future__ import annotations

from ... import concepts, templates
from smartass.xl.host import InputError
from . import Draft, RecipeContext

DEFAULTS = {"compare": (lambda v: bool(v), "the standard comparison is the same period last year")}
FOLLOWUPS = ["Monthly trend through {period}", "Actual vs budget, YTD {period}"]


def concept(params: dict, ctx: RecipeContext) -> Draft:
    if not params.get("concept"):
        raise InputError("recipe concept needs params.concept: a concept id from knowledge\\concepts.json.")
    c = concepts.get(params["concept"])                 # ValueError naming the known concepts for an unknown id
    period = params.get("period") or ctx.closed
    ly = params.get("compare") or templates.same_period_ly(period)
    pulls, spec = concepts.plan(c, period, ly, params.get("entity"))
    pulls = [{k: v for k, v in p.items() if v is not None} for p in pulls]
    return Draft(pulls, spec, assumptions=[f"Actual, NOK, YTD {period} vs YTD {ly}; definition "
                                           f"'{c.get('id')}' from the concept library"],
                 result={"period": period, "concept": c.get("id")})
