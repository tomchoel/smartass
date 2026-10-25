"""Finance tools on the warehouse and the calc engine: fin_pnl (the first slice) and sv_warehouse. Registered on the
shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

import contextlib
from typing import Literal

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from . import atlas, catalog, extract, fin_spec, qlog, refresh, templates
from . import fin_pnl as pnl                      # the tool below is called fin_pnl
from .analytics import periods
from .host import HostError, InputError
from .server import WRITE, _call_user, _err, tool
from .timing import Timer, stage
from .warehouse import OPEN, MemoryWarehouse, Warehouse, WarehouseUnavailable

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


UNIT_CAP = 40                                     # sub-units in the bridge (the largest 5 shown, the rest grouped)


def _units(db, entity: catalog.Entity) -> tuple[tuple[tuple[str, str], ...], str | None]:
    """((member, label) of the area's sub-units from the atlas, for the sub-unit bridge; a note when the cap left
    some out)."""
    labels = {}

    def kids(m):
        rows = [r for r in atlas.children(db, "FinMax", m) if r["dimension"] == "All Entities"]
        labels.update({r["name"]: templates._clean(r["label"]) or r["name"] for r in rows})
        return [r["name"] for r in rows]
    names = refresh.sub_units(kids, entity.finmax, limit=10_000) if entity.finmax else []
    note = None
    if len(names) > UNIT_CAP:
        note = (f"{entity.label} has {len(names)} sub-units; the first {UNIT_CAP} are read for the sub-unit bridge, "
                f"the other {len(names) - UNIT_CAP} are in 'Other units and eliminations'.")
        names = names[:UNIT_CAP]
    return tuple((n, labels.get(n, n)) for n in names), note


def _aliases(db):
    """alias(cube, account) -> the atlas label of the account, cleaned ('C30 REVENUES'), or None."""
    def alias(cube, account):
        try:
            r = db.execute("SELECT label FROM member WHERE connection=? AND dimension=? AND name=? LIMIT 1",
                           (cube, catalog.ACCOUNT_DIM.get(cube, ""), account)).fetchone()
        except Exception:   # noqa: BLE001 - an atlas without the table: the catalogue labels
            return None
        return (templates._clean(r[0]) or None) if r and r[0] else None
    return alias


def _forecast(st, retrieve, cat, entity, p, closed, newest_first: list[str], fills: list):
    """The forecast to use: the newest with data on FinPlan's grand line for the year (the newest is read by
    top_up). An empty newest member (created, not filled in) falls back to the next one, read live: (name, note)."""
    yy = templates.split_period(p)[1]
    if not newest_first or not entity.finplan:
        return (newest_first[0] if newest_first else None), None
    newest = newest_first[0]
    if pnl.forecast_has_data(st, cat, entity, newest, yy):
        return newest, None
    for older in newest_first[1:]:
        grids = extract.finplan_grids(cat, [(older, "Final", yy)], [entity.finplan])
        fills.append(extract.fill(st, retrieve, cat, grids, "top_up", closed, closed_label=OPEN))
        if pnl.forecast_has_data(st, cat, entity, older, yy):
            return older, f"{newest} is empty; {older} used"
    return newest, f"{newest} is empty and no earlier forecast of FY{yy:02d} has data"


def _open(store):
    """The store to use: the given one, else the DuckDB warehouse (WarehouseUnavailable when DuckDB will not load)."""
    if store is not None:
        return store
    try:
        import duckdb  # noqa: F401 - the warehouse needs it; a broken install answers from live pulls
        return Warehouse()
    except ImportError as e:
        raise WarehouseUnavailable(f"DuckDB will not load: {e}") from None


def check_period(p: str, closed: str) -> None:
    """InputError for a period the warehouse cannot answer: after the open month (the month after the closed one),
    or so old that last year's comparison (from its opening December) is outside the 36-month window."""
    nxt = periods.shift(closed, 1)
    if periods.after(p, nxt):
        raise InputError(f"{p} is not available yet: the last closed month is {closed}; the open month is {nxt}.")
    window = periods.window(closed)
    need = periods.opening(periods.ly(p))
    if periods.index(need) < periods.index(window[0]):
        earliest = next(m for m in window if periods.index(periods.opening(periods.ly(m))) >= periods.index(window[0]))
        raise InputError(f"{p} is older than the warehouse holds: it keeps the 36 months {window[0]}..{closed}, and "
                         f"{p} against last year needs {need}. Ask for {earliest} or later.")


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
    yy = templates.split_period(p)[1]                          # 'Sep 26' or a readable InputError
    check_period(p, closed)
    retrieve = retrieve or extract.route(host=lambda: h)
    candidates = templates.forecasts(yy)
    units, cap_note = _units(db, entity)
    labels = pnl.line_labels(cat, _aliases(db))
    notes: list[str] = []

    def run(st, fills):
        with stage("top up (live)"):
            fills += refresh.top_up(st, retrieve, cat, entity, p, closed, candidates[0] if candidates else None,
                                    [u for u, _ in units])
            forecast, fc_note = _forecast(st, retrieve, cat, entity, p, closed, candidates, fills)
        scope = pnl.Scope(entity, p, closed, forecast, units, compare, labels=labels, forecast_note=fc_note,
                          notes=(cap_note,) if cap_note else ())
        with stage("calc engine"):
            return pnl.compute(cat, scope, st), fills
    try:
        st = _open(store)
        with stage("warehouse refresh"):
            fills = refresh.refresh(st, retrieve, cat, closed, db, templates.latest_forecast, notes=notes)
        res, fills = run(st, fills)
        return res, "warehouse", fills, notes
    except WarehouseUnavailable as e:
        mem = MemoryWarehouse()
        with stage("live pulls (no warehouse)"):
            fills = [extract.fill(mem, retrieve, cat, extract.base_grids(cat, closed, templates.latest_forecast,
                                                                         keys={entity.key}), "live", closed)]
        res, fills = run(mem, fills)
        return res, f"live pulls (the warehouse was unavailable: {str(e)[:160]})", fills, notes


@tool(annotations=WRITE)
async def fin_pnl(area: str, period: str | None = None, compare: list[Comparison] | None = None,
                  target: str = "session", allow_archive: bool = False, q: int | None = None,
                  detail: Literal["compact", "full"] = "compact", ctx: Context | None = None) -> dict:
    """ONE call for 'P&L for <business area>': the month and YTD against last year (FinMax), budget and the latest
    forecast (FinPlan), materiality flags (catalogue rule), an EBITDA bridge by line and by sub-unit, KPIs (cost per
    FTE first, margins) and the year-end landing. The numbers come from the local warehouse (refreshed automatically
    when a month closes or the atlas shows a restatement; the open month and the latest forecast are read live) and
    are delivered as an analysis in the session's workbook: live Smart View source sheets, every answer cell a
    formula, and checks that the warehouse equals each source sheet. area: e.g. 'Wholesale', 'Business', '2TEC'
    (Consumer and Infrastructure Growth were abolished on 1 June 2026: the tool names their successors);
    period: 'Sep 26' (default: the last closed month); compare: any of 'ly', 'budget', 'forecast' (default all) -
    it ONLY decides which comparisons get flags and comments: the workbook and the table ALWAYS show last year,
    budget, the forecast and the landing. An empty newest forecast falls back to the newest one with data (said in
    warnings). A finished year (December of an earlier year) has no landing; an open period's flags are provisional.
    detail: 'compact' (default: the per-line table, flags, landing, missing_count) | 'full' (adds every fact).
    Write the reply from the returned key_points, table, flags, warnings and freshness - never compute a number
    yourself; pass every warning on."""
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
        warnings = pnl.tool_warnings(res, fills, notes)
        out["warnings"] = list(dict.fromkeys(warnings + list(out.get("warnings") or [])))
        return {**out, **pnl.summary(res, mode, fills, notes, cat=cat, detail=detail)}
    try:
        out = await _call_user(job, what="fin_pnl")
    except (HostError, ToolError, ValueError) as e:      # ValueError: InputError (a bad period) and engine errors
        qlog.write({"tool": "fin_pnl", "area": area, "period": period, "error": str(e)})
        raise _err(e)
    qlog.write({"tool": "fin_pnl", "area": area, "period": out["period"], "workbook": out["workbook"],
                "q": out["question"], "checks": out["checks"], "timings": out["timings"]})
    return out


def warehouse_job(h, action: str, area: str | None = None, store=None, retrieve=None, db=None) -> dict:
    """sv_warehouse refresh | rebuild (testable without the server). rebuild without an area: delete the file and
    do a full base fill; with an area: re-read its 36 months and the company total."""
    cat = catalog.load()
    with stage("closed month"):
        closed = templates.closed_month(h)
    st = _open(store)
    what = None
    if action == "rebuild":
        if area:
            what = refresh.request_rebuild(st, resolve_area(cat, area).key)
        else:
            st.reset()
            what = refresh.request_rebuild(st)
    notes: list[str] = []
    retrieve = retrieve or extract.route(host=lambda: h)
    with stage("warehouse refresh"), contextlib.ExitStack() as stack:
        adb = db if db is not None else stack.enter_context(contextlib.closing(atlas.connect()))
        fills = refresh.refresh(st, retrieve, cat, closed, adb, templates.latest_forecast, notes=notes)
    out = {"closed_month": closed, "notes": notes,
           "jobs": [{"job": r.fill.job, "result": r.fill.result, "cells": r.fill.cells,
                     "seconds": r.fill.seconds, **r.fill.detail} for r in fills]
           or ("nothing ran: " + "; ".join(notes) if notes else "nothing to do: the warehouse is up to date")}
    if what:
        out["rebuild"] = what
    return out


@tool(annotations=WRITE)
async def sv_warehouse(action: Literal["status", "refresh", "rebuild"] = "status", area: str | None = None) -> dict:
    """The local warehouse behind fin_pnl. status: file, rows, closed month, last fills and their reconciliation
    (cheap, no EPM). refresh: run the refresh check now (a base fill the first time, about 10 s; afterwards only
    what changed). rebuild: re-read from EPM - with `area` (e.g. 'Wholesale') that area's 36 months and the company
    total; without, the whole file is recreated (a full base fill; also the fix for 'warehouse schema is older').
    Use rebuild when a fin_pnl check 'Warehouse = source sheet' shows DIFF: call sv_warehouse('rebuild', area=<the
    area>), then fin_pnl again. Read-only against EPM."""
    if action == "status":
        try:
            return _open(None).status()
        except WarehouseUnavailable as e:
            raise ToolError(f"warehouse unavailable: {e}. Next step: retry in a minute (another smartass process "
                            "may be writing it); if it repeats, call sv_status.") from None

    def job(h):
        Timer.start()
        return warehouse_job(h, action, area)
    try:
        return await _call_user(job, what=f"sv_warehouse {action}")
    except (HostError, ToolError, WarehouseUnavailable) as e:
        raise _err(e)
