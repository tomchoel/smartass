"""Warehouse refresh triggers (Stage 2), checked at the start of every fin_* call. The check itself makes no EPM call
when nothing changed: the closed month is cached per day (templates.closed_month), the atlas snapshot and the
latest forecast name are local.
- No base fill yet: the base fill.
- A new closed month: fill the new month(s) and re-read the two before (late postings), with the YTD spot column.
- The atlas shows changed members under an area (atlas_history, i.e. a reorg or restatement): re-read that area's
  36 months (and the company total).
- The latest forecast name changed: re-read FinPlan's forecast.
A job that fails leaves its state key as it was, so the next call tries again; the answer uses the stale data and
says so. top_up() reads what a question needs live: the latest forecast, an open month and the sub-units."""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Callable

from . import atlas, atlas_history
from .analytics import periods
from .catalog import Catalog, Entity
from .extract import FillResult, Grid, areas, base_grids, base_plans, fill, finmax_grids, finplan_grids
from .grid import Axis
from .warehouse import OPEN

CLOSED, SNAPSHOT, FORECAST = "closed_month", "entities_snapshot", "forecast"


@dataclass(frozen=True)
class Job:
    kind: str                         # base | new_month | restatement | forecast
    grids: tuple[Grid, ...]
    why: str
    state: dict                       # state keys to set when the job succeeds


def plan_refresh(cat: Catalog, state: dict[str, str], closed: str, snapshot: int | None, forecast: str | None,
                 touched: set[str], forecast_of: Callable[[int], str | None]) -> list[Job]:
    """The jobs to run (pure). state: the warehouse's state keys; snapshot: the newest FinMax entity snapshot id;
    forecast: the latest forecast name for the closed month's year; touched: area keys with changed members."""
    snap = "" if snapshot is None else str(snapshot)
    if not state.get(CLOSED):
        return [Job("base", tuple(base_grids(cat, closed, forecast_of)), "no base fill yet",
                    {CLOSED: closed, SNAPSHOT: snap, FORECAST: forecast or ""})]
    jobs = []
    last = state[CLOSED]
    if last != closed:
        start = last if periods.after(closed, last) else closed
        months = sorted(set(periods.between(start, closed)) | {periods.shift(closed, -2), periods.shift(closed, -1),
                                                               closed}, key=periods.index)
        fm = [e.finmax for e in areas(cat, "FinMax")]
        fp = [e.finplan for e in areas(cat, "FinPlan")]
        years = sorted({periods.split_period(m)[1] for m in months})
        jobs.append(Job("new_month", tuple(finmax_grids(cat, months, fm, spot=closed)
                                           + finplan_grids(cat, [("Actual", "Working", y) for y in years], fp)),
                        f"closed month {last} -> {closed}", {CLOSED: closed}))
    if snap != state.get(SNAPSHOT, "") and touched:
        jobs.append(Job("restatement", tuple(base_grids(cat, closed, forecast_of, keys=touched | {"company"})),
                        f"the atlas changed under {', '.join(sorted(touched))}", {SNAPSHOT: snap}))
    elif snap != state.get(SNAPSHOT, ""):
        jobs.append(Job("restatement", (), "the atlas changed outside the areas", {SNAPSHOT: snap}))
    if (forecast or "") != state.get(FORECAST, ""):
        yy = periods.split_period(closed)[1]
        fp = [e.finplan for e in areas(cat, "FinPlan")]
        plans = [p for p in base_plans(closed, forecast_of) if p[0] == forecast and p[2] == yy]
        jobs.append(Job("forecast", tuple(finplan_grids(cat, plans, fp)),
                        f"latest forecast {state.get(FORECAST) or 'none'} -> {forecast or 'none'}",
                        {FORECAST: forecast or ""}))
    return jobs


def touched_areas(cat: Catalog, db, old: str | None, new: int | None) -> set[str]:
    """Area keys whose members changed between two FinMax entity snapshots (added, removed, moved or relabelled
    members, and the parents they moved from and to, traced up the atlas). Empty when either snapshot is unknown;
    when the stored old id no longer matches a snapshot of this dimension (ValueError), every area counts as touched."""
    if not old or new is None or str(new) == old:
        return set()
    by_member = {e.finmax: e.key for e in areas(cat, "FinMax") if e.key != "company"}
    try:
        d = atlas_history.diff_snapshots(db, int(old), new)
    except ValueError:                 # the stored snapshot id is unknown or of another dimension: trust nothing
        return set(by_member.values())
    names: set[str] = set()
    for kind in ("added", "removed"):
        for m in d.get(kind, []):
            names |= {m["name"], *m.get("parents", [])}
    for m in d.get("moved", []):
        names |= {m["name"], *m.get("from", []), *m.get("to", [])}
    names |= {m["name"] for m in d.get("relabelled", [])}
    out = set()
    for n in names:
        for x in [n, *atlas.ancestors(db, "FinMax", n)]:
            if x in by_member:
                out.add(by_member[x])
    return out


def refresh(store, retrieve, cat: Catalog, closed: str, db, forecast_of: Callable[[int], str | None],
            now: Callable[[], float] = time.time) -> list[FillResult]:
    """Run what plan_refresh asks for; each successful job updates its state keys."""
    state = store.state()
    snapshot = atlas_history.latest_snapshot(db, "FinMax", "All Entities")
    forecast = forecast_of(periods.split_period(closed)[1])
    touched = touched_areas(cat, db, state.get(SNAPSHOT), snapshot)
    out = []
    for job in plan_refresh(cat, state, closed, snapshot, forecast, touched, forecast_of):
        if job.grids:
            res = fill(store, retrieve, cat, list(job.grids), job.kind, closed,
                       atlas_build="" if snapshot is None else str(snapshot), now=now)
            out.append(res)
            if not res.ok:
                continue
        store.set_state(**job.state)
    return out


def sub_units(children: Callable[[str], list[str]], entity: str, limit: int = 40) -> list[str]:
    """The sub-units of an area for its bridge: its children, going down through single-child levels (2WHS has
    one child, 3WHS_W, whose children are the units)."""
    kids = children(entity)
    seen = {entity}
    while len(kids) == 1 and kids[0] not in seen:
        seen.add(kids[0])
        kids = children(kids[0])
    return kids[:limit]


def top_up(store, retrieve, cat: Catalog, entity: Entity, period: str, closed: str, forecast: str | None,
           units: list[str], ebitda: str = "ac.EBITDA", now: Callable[[], float] = time.time) -> list[FillResult]:
    """Read live what one question needs and store it: the area's latest forecast and, when `period` is after the
    closed month, its open months (both with closed_month 'open'); and the sub-units' YTD EBITDA now and last year
    (closed history, stored as it comes)."""
    yy = periods.split_period(period)[1]
    live: list[Grid] = []
    if forecast and entity.finplan:
        live += finplan_grids(cat, [(forecast, "Final", yy)], [entity.finplan])
    if periods.after(period, closed):
        if entity.finmax:
            live += finmax_grids(cat, periods.between(closed, period), [entity.finmax])
        if entity.finplan:
            live += finplan_grids(cat, [("Actual", "Working", yy)], [entity.finplan])
    out = []
    if live:
        out.append(fill(store, retrieve, cat, live, "top_up", closed, closed_label=OPEN, now=now))
    if units:
        g = Grid("FinMax", {"All Scenarios": "Actual", "All Currencies": "NOK", "All Accounts": ebitda,
                            "All Balance Types": "Accumulated"}, [Axis("All Entities", units)],
                 [Axis("All Time Periods", [period, periods.ly(period)])], "FinMax sub-unit EBITDA YTD")
        out.append(fill(store, retrieve, cat, [g], "top_up", closed, now=now))
    return out
