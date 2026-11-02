"""The pnl recipe's data side: area/unit resolution, the warehouse or live store, top-up and the calc engine."""
from __future__ import annotations

from mcp.server.mcpserver.exceptions import ToolError

from . import atlas, catalog, extract, refresh, templates
from . import fin_pnl as pnl
from .analytics import periods
from .host import InputError
from .timing import stage
from .warehouse import OPEN, MemoryWarehouse, Warehouse, WarehouseUnavailable


def resolve_area(cat: catalog.Catalog, area: str, units: bool = False) -> catalog.Entity:
    """The one business area `area` names; a readable ToolError for none or several. An abolished area (2CON
    Consumer, 2IG Infrastructure Growth: emptied on 1 June 2026, history restated, so its figures are empty) is never
    answered: the error explains the split and lists the successors to offer the user. units: the caller also
    takes a unit below an area (fin_pnl), said in the error."""
    found = catalog.resolve_entities(cat, area)
    known = ", ".join(e.label for e in cat.entities if not e.abolished)
    if not found:
        if units:
            raise ToolError(f"'{area}' names no business area and no FinMax entity the local atlas knows. Next step: "
                            f"find the unit with sv_find_member('{area}') and call fin_pnl with its member name "
                            f"(e.g. 'un.SY.5OK_W'), or with one of: {known}.")
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


UNIT = "unit:"                                    # key prefix of the synthetic Entity of a unit below an area
LIVE_UNIT = "live pulls (sub-unit: not kept in the warehouse)"
ENTITIES = "All Entities"


def is_unit(entity: catalog.Entity) -> bool:
    """A unit below a catalogue area: always read live, never kept in the warehouse (the refresh triggers only know
    the catalogue areas, so stored unit rows would go stale silently)."""
    return entity.key.startswith(UNIT)


def _entity_names(db, text: str) -> list[str]:
    """FinMax entity names equal to `text`, with or without 'un.SY.', any case (names are ASCII codes)."""
    t = text.strip().lower()
    try:
        rows = db.execute("SELECT DISTINCT name FROM member WHERE connection='FinMax' AND dimension=? AND "
                          "lower(name) IN (?, ?) ORDER BY name", (ENTITIES, t, f"un.sy.{t}")).fetchall()
    except Exception:   # noqa: BLE001 - an atlas without the table: catalogue areas only
        return []
    return [r[0] for r in rows]


def _finplan_of(db, member: str) -> str | None:
    """The FinPlan Entity member of a FinMax unit: its code then a space ('un.SY.4SEL_T' -> '4SEL_T Network Service
    Enablement'); None when FinPlan has none (most units below an area)."""
    code = member.removeprefix("un.SY.")
    r = db.execute("SELECT name FROM member WHERE connection='FinPlan' AND dimension='Entity' AND substr(name, 1, ?) "
                   "= ? ORDER BY name LIMIT 1", (len(code) + 1, f"{code} ")).fetchone()
    return r[0] if r else None


def resolve_scope(cat: catalog.Catalog, area: str, db) -> catalog.Entity:
    """fin_pnl's entity: a catalogue area (resolve_area), or any FinMax entity the atlas knows below a non-abolished
    area (e.g. 'un.SY.5OK_W' or '5OK_W') as a synthetic Entity (key 'unit:<member>', label from the atlas, the
    FinPlan member with the same code if any). A readable ToolError otherwise."""
    names = _entity_names(db, area)
    if len(names) > 1:
        raise ToolError(f"'{area}' matches several FinMax entities ({', '.join(names)}). Next step: ask the user "
                        "which one, then call fin_pnl with its member name.")
    if not names:
        return resolve_area(cat, area, units=True)
    m = names[0]
    by_member = {e.finmax: e for e in cat.entities if e.finmax}
    if m in by_member:                                     # an area's own member: the area
        return resolve_area(cat, m)
    top = next((by_member[a] for a in atlas.ancestors(db, "FinMax", m) if a in by_member), None)
    if top is None:
        raise ToolError(f"{m} is not below a business area in the atlas. Next step: check it with "
                        f"sv_member_info('{m}') and ask the user for the area or unit they mean.")
    if top.abolished:
        raise ToolError(f"{m} is below {top.label} ({top.finmax}), abolished on {top.abolished}: {top.note} Its "
                        "figures are empty after the restatement. Next step: find the unit's new place with "
                        "sv_find_member and ask the user, then call fin_pnl with that member.")
    r = db.execute("SELECT label FROM member WHERE connection='FinMax' AND dimension=? AND name=? LIMIT 1",
                   (ENTITIES, m)).fetchone()
    return catalog.Entity(key=UNIT + m, label=templates._clean(r[0] if r else None) or m, aliases=(), finmax=m,
                          finplan=_finplan_of(db, m), note=f"a unit below {top.label} ({top.finmax})")


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
    if not names:                                          # a leaf unit (or an atlas without its children)
        note = f"{entity.label} has no sub-units in the local atlas: no sub-unit bridge."
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
    if not entity.finplan:                                 # no FinPlan member: no forecast to use
        return None, None
    if not newest_first:
        return None, None
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
    be opened, otherwise live pulls of this area into memory (slower, and the answer says so). A unit below an area
    (is_unit) is always read live into memory and never touches the warehouse (`store` is not used). db: the atlas
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
    scope_notes = [n for n in (f"{entity.label} ({entity.finmax}) is {entity.note}; read live, not kept in the "
                               "warehouse." if is_unit(entity) else None, cap_note) if n]

    def run(st, fills):
        with stage("top up (live)"):
            fills += refresh.top_up(st, retrieve, cat, entity, p, closed,
                                    candidates[0] if candidates and entity.finplan else None, [u for u, _ in units])
            forecast, fc_note = _forecast(st, retrieve, cat, entity, p, closed, candidates, fills)
        scope = pnl.Scope(entity, p, closed, forecast, units, compare, labels=labels, forecast_note=fc_note,
                          notes=tuple(scope_notes))
        with stage("calc engine"):
            return pnl.compute(cat, scope, st), fills

    def live(why):
        mem = MemoryWarehouse()
        with stage(f"live pulls ({why})"):
            fills = [extract.fill(mem, retrieve, cat, extract.base_grids(cat, closed, templates.latest_forecast,
                                                                         entities=[entity]), "live", closed)]
        return run(mem, fills)
    if is_unit(entity):
        res, fills = live("sub-unit")
        return res, LIVE_UNIT, fills, notes
    try:
        st = _open(store)
        with stage("warehouse refresh"):
            fills = refresh.refresh(st, retrieve, cat, closed, db, templates.latest_forecast, notes=notes)
        res, fills = run(st, fills)
        return res, "warehouse", fills, notes
    except WarehouseUnavailable as e:
        res, fills = live("no warehouse")
        return res, f"live pulls (the warehouse was unavailable: {str(e)[:160]})", fills, notes


