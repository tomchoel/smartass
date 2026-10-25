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

from . import atlas, atlas_history, config
from .drill_browser import ProfileLock
from .analytics import periods
from .catalog import Catalog, Entity
from .extract import FillResult, Grid, areas, base_grids, base_plans, fill, finmax_grids, finplan_grids
from .grid import Axis
from .warehouse import OPEN, WarehouseUnavailable

CLOSED, SNAPSHOT, FORECAST = "closed_month", "entities_snapshot", "forecast"
REBUILD = "rebuild_areas"           # area keys to re-read in full (sv_warehouse rebuild <area>), comma-separated
FIRST_FILL_WAITS = 6                # a first fill running elsewhere: wait up to this x WAREHOUSE_LOCK_WAIT_S
RETRY_FIRST_S, RETRY_LATER_S = 15 * 60, 24 * 3600     # a permanent content error: retry after 15 min, then daily


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
    rebuild = {k for k in (state.get(REBUILD) or "").split(",") if k}
    if rebuild:
        jobs.append(Job("rebuild", tuple(base_grids(cat, closed, forecast_of, keys=rebuild | {"company"})),
                        f"rebuild requested for {', '.join(sorted(rebuild))}", {REBUILD: ""}))
    last = state[CLOSED]
    if last != closed:
        start = last if periods.after(closed, last) else closed
        months = sorted(set(periods.between(start, closed)) | {periods.shift(closed, -2), periods.shift(closed, -1),
                                                               closed}, key=periods.index)
        fm = [e.finmax for e in areas(cat, "FinMax")]
        fp = [e.finplan for e in areas(cat, "FinPlan")]
        years = sorted({periods.split_period(m)[1] for m in months})
        yy = periods.split_period(closed)[1]
        plans = [("Actual", "Working", y) for y in years] + [("Target", "Final", y) for y in (yy, yy + 1)]
        plans += [(fc, "Final", y) for y in (yy, yy + 1) if (fc := forecast_of(y))]    # year roll: FY+1 appears
        jobs.append(Job("new_month", tuple(finmax_grids(cat, months, fm, spot=closed)
                                           + finplan_grids(cat, list(dict.fromkeys(plans)), fp)),
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
    company = next((e for e in cat.entities if e.key == "company"), None)
    if company and company.finmax in names:            # a change directly under the root: the company total only
        out.add("company")
    for n in names:
        for x in [n, *atlas.ancestors(db, "FinMax", n)]:
            if x in by_member:
                out.add(by_member[x])
    return out


def refresh(store, retrieve, cat: Catalog, closed: str, db, forecast_of: Callable[[int], str | None],
            now: Callable[[], float] = time.time, notes: list[str] | None = None) -> list[FillResult]:
    """Run what plan_refresh asks for; each successful job updates its state keys. One refresh at a time per
    warehouse (a non-blocking <db>.refresh.lock): when another process holds it, nothing runs and `notes` (if given)
    says so, so the answer comes from the current data - except before the first fill, when it waits for the other
    process (up to FIRST_FILL_WAITS x WAREHOUSE_LOCK_WAIT_S) and then raises WarehouseUnavailable. A job that
    failed on content (#NoAccess, an unknown member) is not retried for 15 minutes, then once a day; a transport
    error is retried on the next call."""
    notes = notes if notes is not None else []
    path = getattr(store, "path", None)
    lock = ProfileLock(str(path) + ".refresh") if path else None
    if lock is not None and not lock.acquire():
        if store.state().get(CLOSED):
            notes.append("warehouse refresh skipped: another smartass process is refreshing; answering from the "
                         "current data")
            return []
        # nothing filled yet: an answer from an empty warehouse would be all gaps - wait for that first fill
        deadline = time.monotonic() + config.WAREHOUSE_LOCK_WAIT_S * FIRST_FILL_WAITS
        while not lock.acquire():
            if time.monotonic() > deadline:
                raise WarehouseUnavailable("the first warehouse fill is still running in another smartass process")
            time.sleep(0.2)
    try:
        return _refresh(store, retrieve, cat, closed, db, forecast_of, now, notes)
    finally:
        if lock is not None:
            lock.release()


def _refresh(store, retrieve, cat, closed, db, forecast_of, now, notes) -> list[FillResult]:
    state = store.state()                                  # read after the lock: another process may just have run
    snapshot = atlas_history.latest_snapshot(db, "FinMax", "All Entities")
    forecast = forecast_of(periods.split_period(closed)[1])
    touched = touched_areas(cat, db, state.get(SNAPSHOT), snapshot)
    out = []
    for job in plan_refresh(cat, state, closed, snapshot, forecast, touched, forecast_of):
        wait = float(state.get(f"retry_after_{job.kind}") or 0) - now()
        if wait > 0 and job.grids:
            notes.append(f"warehouse refresh '{job.kind}' paused for {max(1, round(wait / 60))} min after: "
                         f"{state.get(f'last_error_{job.kind}', '')}")
            continue
        try:
            if job.grids:
                res = fill(store, retrieve, cat, list(job.grids), job.kind, closed,
                           atlas_build="" if snapshot is None else str(snapshot), now=now,
                           atlas_db=db)
                out.append(res)
                if not res.ok:
                    if res.fill.detail.get("content_only"):
                        fails = int(state.get(f"fails_{job.kind}") or 0) + 1
                        store.set_state(**{f"retry_after_{job.kind}": now() + (RETRY_FIRST_S if fails == 1
                                                                                 else RETRY_LATER_S),
                                           f"fails_{job.kind}": fails,
                                           f"last_error_{job.kind}": (res.fill.detail["errors"] or [""])[0][:300]})
                    continue
            store.set_state(**job.state, **{f"retry_after_{job.kind}": 0, f"fails_{job.kind}": 0,
                                            f"last_error_{job.kind}": ""})
        except WarehouseUnavailable as e:
            notes.append(f"warehouse state not saved: {e}")
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
           units: list[str], ebitda: str | None = None, now: Callable[[], float] = time.time) -> list[FillResult]:
    """Read live what one question needs and store it: the area's latest forecast and, when `period` is after the
    closed month, ONLY its open months (all stored with closed_month 'open'; the warehouse never lets an open row
    replace a closed one); and the sub-units' YTD on the catalogue's grand line now and last year (labelled open when
    the period is open, else closed history)."""
    yy = periods.split_period(period)[1]
    is_open = periods.after(period, closed)
    live: list[Grid] = []
    if forecast and entity.finplan:
        live += finplan_grids(cat, [(forecast, "Final", yy)], [entity.finplan])
    if is_open:
        opened = periods.between(closed, period)
        if entity.finmax:
            live += finmax_grids(cat, opened, [entity.finmax], opening=False)
        if entity.finplan:
            years = sorted({periods.split_period(m)[1] for m in opened})
            live += finplan_grids(cat, [("Actual", "Working", y) for y in years], [entity.finplan], months=opened)
    out = []
    if live:
        out.append(fill(store, retrieve, cat, live, "top_up", closed, closed_label=OPEN, now=now))
    if units:
        account = ebitda or next(x.account for x in cat.pnl["FinMax"] if x.kind == "grand")
        for p, label in ((period, OPEN if is_open else None), (periods.ly(period), None)):   # last year is history
            g = Grid("FinMax", {"All Scenarios": "Actual", "All Currencies": "NOK", "All Accounts": account,
                                "All Balance Types": "Accumulated"}, [Axis("All Entities", units)],
                     [Axis("All Time Periods", [p])], f"FinMax sub-unit YTD {p}")
            out.append(fill(store, retrieve, cat, [g], "top_up", closed, closed_label=label, now=now))
    return out


def request_rebuild(store, area: str | None = None) -> str:
    """sv_warehouse rebuild: with an area key, its 36 months (and the company total) are re-read on the next refresh;
    without, every state key is cleared so the next refresh is a full base fill (the caller resets the file first).
    Returns what was requested."""
    if area:
        keys = {k for k in (store.state().get(REBUILD) or "").split(",") if k} | {area}
        store.set_state(**{REBUILD: ",".join(sorted(keys)), "retry_after_rebuild": 0})
        return f"{area} and the company total: all 36 months"
    store.set_state(**{CLOSED: "", SNAPSHOT: "", FORECAST: "", REBUILD: "", "retry_after_base": 0})
    return "every area: a full base fill"
