"""Recipe pnl (B01, B13, B15): the full P&L of a business area or unit - month and YTD vs last year (FinMax),
budget and the latest forecast (FinPlan), bridges, KPIs, landing - from the warehouse facts. Wraps
fin_spec.build unchanged; the warehouse = source sheet reconciliation travels as layout.expect and becomes
static checks after the pull (pipeline.expect_checks)."""
from __future__ import annotations

from ... import fin_spec, tools_fin
from ... import fin_pnl as pnl_mod                  # the P&L engine (tools_fin imports it the same way)
from ...host import InputError
from . import Draft, RecipeContext


def pnl(params: dict, ctx: RecipeContext) -> Draft:
    if ctx.pnl_facts is None:
        raise InputError("The pnl recipe needs the warehouse (no facts function in this context).")
    if not params.get("area"):
        raise InputError("recipe pnl needs params.area: a business area or unit, as in fin_pnl.")
    compare = tuple(params.get("compare") or pnl_mod.COMPARISONS)
    res, mode, fills, notes, _entity = ctx.pnl_facts(params["area"], params.get("period"), compare)
    req = fin_spec.build(res, ctx.cat, now=ctx.now or None, mode=mode, refresh_notes=notes)
    lay = dict(req.spec)
    stale = tools_fin.stale_note(fills)
    if stale:
        lay["notes"] = [stale] + list(lay.get("notes") or [])
    lay["expect"] = {"source": req.source, "cells": {k: [list(c) for c in v] for k, v in req.expected.items()}}
    summary = pnl_mod.summary(res, mode, fills, notes, cat=ctx.cat, detail=params.get("detail", "compact"))
    pulls = [{k: v for k, v in p.items() if v is not None} for p in req.pulls]
    return Draft(pulls, lay, assumptions=[f"{res.scope.entity.label}, NOK, {res.scope.period}"],
                 warnings=pnl_mod.tool_warnings(res, fills, notes), result=summary)
