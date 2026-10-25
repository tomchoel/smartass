"""Finance tools on the warehouse and the calc engine: fin_pnl (the first slice) and sv_warehouse. Registered on the
shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

import contextlib
from typing import Literal

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from . import atlas, catalog, extract, fin_spec, qlog, refresh, templates
from . import fin_pnl as pnl                      # the tool below is called fin_pnl
from .host import HostError
from .server import WRITE, _call_user, _err, tool
from .timing import Timer, stage
from .warehouse import MemoryWarehouse, Warehouse, WarehouseUnavailable

Comparison = Literal["ly", "budget", "forecast"]


def resolve_area(cat: catalog.Catalog, area: str) -> catalog.Entity:
    """The one business area `area` names; a readable ToolError for none or several. An abolished area (2CON
    Consumer, 2IG Infrastructure Growth: emptied on 1 June 2026, history restated, so its figures are empty) is never
    answered: the error explains the split and lists the successors to offer the user."""
    found = catalog.resolve_entities(cat, area)
    known = ", ".join(e.label for e in cat.entities if not e.abolished)
    if not found:
        raise ToolError(f"'{area}' names no business area. Next step: call fin_pnl with one of: {known}.")
    if len(found) > 1:
        raise ToolError(f"'{area}' names several areas ({', '.join(e.label for e in found)}). Next step: ask the "
                        "user which one, then call fin_pnl once per area.")
    e = found[0]
    if e.abolished:
        succ = [x for x in cat.entities if x.key in e.successors]
        names = ", ".join(f"{x.label} ({x.finmax})" for x in succ)
        raise ToolError(f"{e.label} ({e.finmax}) was abolished on {e.abolished}: {e.note} Its figures are empty "
                        f"after the restatement, so they are not reported. Successors: {names}. Next step: tell the "
                        "user about the split and ask which successor(s) they want, then call fin_pnl once per "
                        "successor.")
    return e


def stale_note(fills) -> str | None:
    """A visible warning when a refresh or top-up job failed: the answer then uses the earlier warehouse data."""
    failed = [r.fill for r in fills if r.fill.result == "error"]
    if not failed:
        return None
    first = (failed[0].detail.get("errors") or ["?"])[0]
    return (f"Warning: {len(failed)} warehouse job(s) failed this time ({failed[0].job}: {first[:140]}); figures "
            "that were not refreshed come from the earlier read shown in the data line.")


def _units(db, entity: catalog.Entity) -> tuple[tuple[str, str], ...]:
    """(member, label) of the area's sub-units from the atlas, for the sub-unit bridge."""
    labels = {}

    def kids(m):
        rows = [r for r in atlas.children(db, "FinMax", m) if r["dimension"] == "All Entities"]
        labels.update({r["name"]: templates._clean(r["label"]) or r["name"] for r in rows})
        return [r["name"] for r in rows]
    names = refresh.sub_units(kids, entity.finmax) if entity.finmax else []
    return tuple((n, labels.get(n, n)) for n in names)


def prepare(h, cat: catalog.Catalog, entity: catalog.Entity, period: str | None, compare: tuple[str, ...],
            store=None, retrieve=None, db=None):
    """Refresh check + top-up + calc engine: (Result, store mode, fills, refresh notes). The warehouse when it can
    be opened, otherwise live pulls of this area into memory (slower, and the answer says so). db: the atlas
    (default: the user's)."""
    own = db is None
    db = atlas.connect() if own else db
    try:
        return _prepare(h, cat, entity, period, compare, store, retrieve, db)
    finally:
        if own:
            db.close()


def _prepare(h, cat, entity, period, compare, store, retrieve, db):
    with stage("closed month"):
        closed = templates.closed_month(h)
    p = period or closed
    templates.split_period(p)                                  # 'Sep 26' or a readable InputError
    retrieve = retrieve or extract.route(host=lambda: h)
    forecast = templates.latest_forecast(templates.split_period(p)[1])
    units = _units(db, entity)
    scope = pnl.Scope(entity, p, closed, forecast, units, compare)
    notes: list[str] = []

    def run(st, fills):
        with stage("top up (live)"):
            fills += refresh.top_up(st, retrieve, cat, entity, p, closed, forecast, [u for u, _ in units])
        with stage("calc engine"):
            return pnl.compute(cat, scope, st), fills
    try:
        st = store if store is not None else Warehouse()
        with stage("warehouse refresh"):
            fills = refresh.refresh(st, retrieve, cat, closed, db, templates.latest_forecast, notes=notes)
        res, fills = run(st, fills)
        return res, "warehouse", fills, notes
    except (WarehouseUnavailable, ImportError) as e:
        mem = MemoryWarehouse()
        with stage("live pulls (no warehouse)"):
            fills = [extract.fill(mem, retrieve, cat, extract.base_grids(cat, closed, templates.latest_forecast,
                                                                         keys={entity.key}), "live", closed)]
        res, fills = run(mem, fills)
        return res, f"live pulls (the warehouse was unavailable: {str(e)[:160]})", fills, notes


@tool(annotations=WRITE)
async def fin_pnl(area: str, period: str | None = None, compare: list[Comparison] | None = None,
                  target: str = "session", allow_archive: bool = False, q: int | None = None,
                  ctx: Context | None = None) -> dict:
    """ONE call for 'P&L for <business area>': the month and YTD against last year (FinMax), budget and the latest
    forecast (FinPlan), materiality flags (catalogue rule), an EBITDA bridge by line and by sub-unit, KPIs (cost per
    FTE first, margins) and the year-end landing. The numbers come from the local warehouse (refreshed automatically
    when a month closes or the atlas shows a restatement; the open month and the latest forecast are read live) and
    are delivered as an analysis in the session's workbook: live Smart View source sheets, every answer cell a
    formula, and checks that the warehouse equals each source sheet. area: e.g. 'Wholesale', 'Business', '2TEC'
    (Consumer and Infrastructure Growth were abolished on 1 June 2026: the tool names their successors);
    period: 'Sep 26' (default: the last closed month); compare: any of 'ly', 'budget', 'forecast' (default all).
    Write the reply from the returned key_points, flags, facts, missing and freshness - never compute a number
    yourself."""
    cat = catalog.load()
    entity = resolve_area(cat, area)
    cmp = tuple(compare or pnl.COMPARISONS)

    def job(h):
        from .tools_ws import PullIn, _build_job
        tm = Timer.start()
        res, mode, fills, notes = prepare(h, cat, entity, period, cmp)
        req = fin_spec.build(res, cat, mode=mode, refresh_notes=notes)
        stale = stale_note(fills)
        if stale:
            req.spec["notes"].insert(0, stale)
        pulls = [PullIn.model_validate(x) for x in req.pulls]
        out = _build_job(h, tm, req.question, pulls, req.spec, target, allow_archive, q, replace=bool(q),
                         before_answer=lambda spec, by_id: fin_spec.with_reconciliation(spec, req.expected, by_id))
        warnings = ([stale] if stale else []) + list(notes) + list(res.problems)
        if warnings:
            out["warnings"] = warnings + list(out.get("warnings") or [])
        return {**out, **pnl.summary(res, mode, fills, notes)}
    try:
        out = await _call_user(job, what="fin_pnl")
    except (HostError, ToolError, ValueError) as e:      # ValueError: InputError (a bad period) and engine errors
        qlog.write({"tool": "fin_pnl", "area": area, "period": period, "error": str(e)})
        raise _err(e)
    qlog.write({"tool": "fin_pnl", "area": area, "period": out["period"], "workbook": out["workbook"],
                "q": out["question"], "checks": out["checks"], "timings": out["timings"]})
    return out


@tool(annotations=WRITE)
async def sv_warehouse(action: Literal["status", "refresh"] = "status") -> dict:
    """The local warehouse behind fin_pnl. status: file, rows, closed month, last fills and their reconciliation
    (cheap, no EPM). refresh: run the refresh check now (a base fill the first time, about a minute; afterwards
    only what changed), read-only against EPM."""
    if action == "status":
        try:
            return Warehouse().status()
        except (WarehouseUnavailable, ImportError) as e:
            raise ToolError(f"warehouse unavailable: {e}. Next step: retry in a minute (another smartass process "
                            "may be writing it); if it repeats, call sv_status.") from None
    cat = catalog.load()

    def job(h):
        Timer.start()
        with stage("closed month"):
            closed = templates.closed_month(h)
        notes: list[str] = []
        with stage("warehouse refresh"), contextlib.closing(atlas.connect()) as db:
            fills = refresh.refresh(Warehouse(), extract.route(host=lambda: h), cat, closed, db,
                                    templates.latest_forecast, notes=notes)
        return {"closed_month": closed, "notes": notes,
                "jobs": [{"job": r.fill.job, "result": r.fill.result, "cells": r.fill.cells,
                          "seconds": r.fill.seconds, **r.fill.detail} for r in fills]
                or ("nothing ran: " + "; ".join(notes) if notes else "nothing to do: the warehouse is up to date")}
    try:
        return await _call_user(job, what="sv_warehouse refresh")
    except (HostError, ToolError, WarehouseUnavailable, ImportError) as e:
        raise _err(e)
