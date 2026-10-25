"""Extract jobs (Stage 2): the grids that fill the warehouse, built from the catalogue, read over Smart View's HTTP
protocol with the COM route as fallback (like tools_query._retrieve), stored as warehouse rows and reconciled.

Base fill: every business area (not abolished) plus the company total x the catalogue's lines x 36 months.
- FinMax Actual: flow lines Periodic (YTD is the sum of the months); stock lines (FTE) Accumulated = the month-end
  level, plus the opening December; one Accumulated spot column at the closed month for the months = YTD check.
- FinPlan: Actual/Working (FY-1, FY), Target/Final and the latest Forecast/Final (FY-1, FY, FY+1, where the forecast
  member exists), Periodic months. FinPlan periods are stored as 'Sep 26'."""
from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from typing import Callable

from . import config, ops
from .analytics import periods, reconcile
from .analytics.frame import from_rows
from .catalog import Catalog, Entity
from .grid import Axis, Layout, Parsed
from .host import InputError
from .warehouse import OPEN, Fill, Row, WarehouseUnavailable, WhKey

log = logging.getLogger(__name__)

Retrieve = Callable[[str, dict, list[Axis], list[Axis]], tuple[Layout, Parsed, float, str]]
FINMAX_POV = {"All Scenarios": "Actual", "All Currencies": "NOK"}
CHUNK = 12                                        # months per FinMax grid: about 1 100 cells, well under a second
FINMAX_SCENARIOS = {("Actual", ""): "actual"}


@dataclass(frozen=True)
class Grid:
    conn: str
    pov: dict
    rows: list[Axis]
    cols: list[Axis]
    what: str

    @property
    def cells(self) -> int:
        n = 1
        for a in self.rows + self.cols:
            n *= len(a.members)
        return n


def areas(cat: Catalog, cube: str) -> list[Entity]:
    """The entities a base fill covers in `cube`: every area that is not abolished, the company total first."""
    attr = "finmax" if cube == "FinMax" else "finplan"
    ents = [e for e in cat.entities if not e.abolished and getattr(e, attr)]
    return sorted(ents, key=lambda e: e.key != "company")


def _norm(d: str) -> str:
    return d.removeprefix("All ").replace(" ", "").casefold()


def wh_key(conn: str, members: dict[str, str]) -> WhKey:
    """The warehouse key of one cell, from its full member set {dimension: member} (any dimension spelling)."""
    m = {_norm(d): v for d, v in members.items()}
    if conn == "FinPlan":
        return (conn, m["scenario"], m["version"], m["entity"], m["account"],
                periods.from_finplan(m["period"], m["years"]), m["balancetypes"])
    return (conn, m["scenarios"], "", m["entities"], m["accounts"], m["timeperiods"], m["balancetypes"])


def to_rows(lay: Layout, parsed: Parsed, conn: str, route: str, fetched_at: float, closed_label: str,
            atlas_build: str = "", closed: str | None = None) -> tuple[list[Row], list[str]]:
    """Warehouse rows for every data cell of a retrieved grid (parsed with suppression off), and the problems:
    cells that are not data (#NoAccess, #Invalid) are not stored. With `closed`, an Actual cell of a month after it
    is not closed history (FinPlan carries blank actuals for the rest of the year): it is labelled OPEN."""
    pov = dict(lay.pov)
    rows, problems = [], []
    if parsed.pov_problems:                      # a POV member Smart View did not know: every value would be wrong
        return rows, problems
    unknown = set(parsed.unknown_members)        # 'Dimension: member' - their cells are comments, not data (not 0)
    for r in parsed.rows:
        combo = r[:lay.nr]
        if any(f"{a.dimension}: {x}" in unknown for a, x in zip(lay.rows, combo)):
            continue
        for j, ctuple in enumerate(lay.col_tuples):
            if any(f"{a.dimension}: {x}" in unknown for a, x in zip(lay.cols, ctuple)):
                continue
            v = r[lay.nr + j]
            members = {**pov, **{a.dimension: x for a, x in zip(lay.rows, combo)},
                       **{a.dimension: x for a, x in zip(lay.cols, ctuple)}}
            key = wh_key(conn, members)
            if isinstance(v, str):
                problems.append(f"{' / '.join(combo)} / {' / '.join(ctuple)}: {v}")
                continue
            label = OPEN if closed and key[1] == "Actual" and periods.after(key[5], closed) else closed_label
            rows.append(Row(key, None if v is None else float(v), route, fetched_at, label, atlas_build))
    return rows, problems


def _chunks(xs: list, n: int) -> list[list]:
    return [xs[i:i + n] for i in range(0, len(xs), n)]


def finmax_grids(cat: Catalog, months: list[str], entities: list[str], spot: str | None = None,
                 opening: bool = True) -> list[Grid]:
    """opening=False leaves the opening December out of the stock levels (a live read of open months only)."""
    lines = cat.pnl["FinMax"]
    flows = [x.account for x in lines if x.measure == "flow"]
    stocks = [x.account for x in lines if x.measure == "stock"]
    out = []
    for ch in _chunks(months, CHUNK):
        out.append(Grid("FinMax", {**FINMAX_POV, "All Balance Types": "Periodic"},
                        [Axis("All Entities", entities), Axis("All Accounts", flows)],
                        [Axis("All Time Periods", ch)], f"FinMax Actual {ch[0]}..{ch[-1]} Periodic"))
    if stocks:
        levels = sorted(set(months) | ({periods.opening(months[0])} if opening else set()), key=periods.index)
        out.append(Grid("FinMax", {**FINMAX_POV, "All Balance Types": "Accumulated"},
                        [Axis("All Entities", entities), Axis("All Accounts", stocks)],
                        [Axis("All Time Periods", levels)], f"FinMax Actual levels {levels[0]}..{levels[-1]}"))
    if spot:
        out.append(Grid("FinMax", {**FINMAX_POV, "All Balance Types": "Accumulated"},
                        [Axis("All Entities", entities), Axis("All Accounts", flows)],
                        [Axis("All Time Periods", [spot])], f"FinMax Actual YTD {spot} (months = YTD check)"))
    return out


def finplan_grids(cat: Catalog, plans: list[tuple[str, str, int]], entities: list[str],
                  months: list[str] | None = None) -> list[Grid]:
    """plans: (scenario member, version, two-digit year); one grid of entities x lines x the 12 months each, or only
    the given months ('Oct 26' labels) of that year - a plan whose year has none of them is left out."""
    accts = [x.account for x in cat.pnl["FinPlan"]]
    out = []
    for s, v, yy in plans:
        mo = list(periods.MONTHS) if months is None else [
            periods.finplan(m)[0] for m in months if periods.split_period(m)[1] == yy]
        if mo:
            out.append(Grid("FinPlan", {"Scenario": s, "Version": v, "Years": f"FY{yy:02d}",
                                        "All Balance Types": "Periodic"},
                            [Axis("Entity", entities), Axis("Account", accts)], [Axis("Period", mo)],
                            f"FinPlan {s}/{v} FY{yy:02d}" + ("" if months is None else f" {mo[0]}..{mo[-1]}")))
    return out


def base_plans(closed: str, forecast_of: Callable[[int], str | None]) -> list[tuple[str, str, int]]:
    yy = periods.split_period(closed)[1]
    plans = [("Actual", "Working", y) for y in (yy - 1, yy)] + [("Target", "Final", y) for y in (yy - 1, yy, yy + 1)]
    return plans + [(fc, "Final", y) for y in (yy - 1, yy, yy + 1) if (fc := forecast_of(y))]


def base_grids(cat: Catalog, closed: str, forecast_of: Callable[[int], str | None],
               keys: set[str] | None = None) -> list[Grid]:
    """The base fill (all areas, or the area keys in `keys`)."""
    fm = [e.finmax for e in areas(cat, "FinMax") if keys is None or e.key in keys]
    fp = [e.finplan for e in areas(cat, "FinPlan") if keys is None or e.key in keys]
    out = finmax_grids(cat, periods.window(closed), fm, spot=closed) if fm else []
    return out + (finplan_grids(cat, base_plans(closed, forecast_of), fp) if fp else [])


def route(http=None, host: Callable[[], object] | None = None) -> Retrieve:
    """Smart View's HTTP protocol first (svhttp.shared() unless `http` is given); on any failure that is not an input
    error, the COM route through the hidden Excel when `host` (a function returning the ExcelHost) is given. A grid
    too big for one HTTP request (cells / 200 > SVHTTP_TIMEOUT_S, like tools_query) goes straight to COM."""
    def retrieve(conn, pov, rows, cols):
        from . import svhttp
        n_cells = 1
        for a in [*rows, *cols]:
            n_cells *= max(1, len(a.members))
        too_big = host is not None and n_cells / 200 > config.SVHTTP_TIMEOUT_S
        if not too_big:
            try:
                lay, p, secs = (http or svhttp.shared()).retrieve(conn, pov, rows, cols, False, False)
                return lay, p, secs, "http"
            except InputError:
                raise
            except Exception as e:   # noqa: BLE001 - the HTTP route failed: Smart View in Excel, if we may
                if host is None:
                    raise
                log.warning("warehouse extract: HTTP route failed (%s); using Smart View in Excel", e)
        h = host()
        lay, p, secs = h.on_host(lambda: ops.query(h, conn, pov, rows, cols, False, False),
                                 timeout=max(config.CALL_TIMEOUT_S, n_cells / 200))
        return lay, p, secs, "com"
    return retrieve


@dataclass
class FillResult:
    fill: Fill
    rows: list[Row] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.fill.result != "error"


def check_batch(cat: Catalog, rows: list[Row], closed: str, store=None) -> tuple[int, list[str], list[str]]:
    """(checks run, problems, information) over the rows of one fill: the identity checks per entity and month in
    both cubes, the Periodic months against the Accumulated spot column (the rest of the YTD months come from
    `store`, when the fill read only the newest ones) and the areas of the batch against the company total."""
    data = {r.key: r.value for r in rows}
    n, problems, info = 0, [], []
    fm = from_rows("FinMax", cat.pnl["FinMax"], data, FINMAX_SCENARIOS)
    forecasts = {k[1] for k in data if k[0] == "FinPlan" and k[1].startswith("Forecast")}
    fp = from_rows("FinPlan", cat.pnl["FinPlan"], data,
                   {("Actual", "Working"): "actual_finplan", ("Target", "Final"): "budget",
                    **{(f, "Final"): f for f in forecasts}})
    for frame in (fm, fp):
        k, p = reconcile.identities(frame, cat.checks[frame.cube])
        n, problems = n + k, problems + p
    flows = {x.account for x in cat.pnl["FinMax"] if x.measure == "flow"}
    ytd = periods.ytd(closed)
    periodic = {k: v for k, v in data.items() if k[0] == "FinMax" and k[6] == "Periodic"}
    spot = {k: v for k, v in data.items() if k[0] == "FinMax" and k[6] == "Accumulated" and k[5] == closed
            and k[4] in flows}
    if spot and store is not None:
        try:
            stored = store.get("FinMax", {k[3] for k in spot}, {k[4] for k in spot}, ytd, "Actual", "", "Periodic",
                               closed_only=True)
            periodic = {**stored, **periodic}
        except WarehouseUnavailable:
            pass
    k, p = reconcile.months_vs_ytd(periodic, spot, ytd)
    n, problems = n + k, problems + p
    company = next((e for e in cat.entities if e.key == "company"), None)
    present = {k[3] for k in data if k[0] == "FinMax"}
    others = [e.finmax for e in areas(cat, "FinMax") if e.key != "company" and e.finmax in present]
    if company and company.finmax and others:
        line = next((x.key for x in cat.pnl["FinMax"] if x.kind == "grand"), None)
        text = reconcile.areas_vs_total(fm, others, company.finmax, line, closed, "actual") if line else None
        if text:
            info.append(text)
    return n, problems, info


def fill(store, retrieve: Retrieve, cat: Catalog, grids: list[Grid], job: str, closed: str,
         closed_label: str | None = None, atlas_build: str = "", now: Callable[[], float] = time.time) -> FillResult:
    """Read every grid, store the rows, reconcile them and record the job in `fill`. A grid that fails is named in
    the result (the others are still stored); a failed job leaves the refresh state as it was, so it is retried."""
    t0, started = time.monotonic(), now()
    rows, errors, warnings, routes, transport = [], [], [], set(), False
    for g in grids:
        try:
            lay, parsed, _secs, how = retrieve(g.conn, g.pov, g.rows, g.cols)
        except Exception as e:   # noqa: BLE001 - one failed grid must not lose the others
            transport = True
            errors.append(f"{g.what}: {type(e).__name__}: {str(e).splitlines()[0][:200] if str(e) else ''}")
            continue
        routes.add(how)
        got, bad = to_rows(lay, parsed, g.conn, how, now(), closed_label or closed, atlas_build, closed)
        rows += got
        found = [f"{g.what}: {b}" for b in bad[:5]] + [f"{g.what}: {w}" for w in parsed.pov_problems]
        found += [f"{g.what}: unknown members {', '.join(parsed.unknown_members)}"] if parsed.unknown_members else []
        if got:
            warnings += found                    # some cells were unusable, the grid's other cells were stored
        else:
            errors += found or [f"{g.what}: no cells returned"]            # the whole grid failed
    detail = {"grids": len(grids), "routes": sorted(routes), "errors": errors[:20], "warnings": warnings[:20],
              "checks": 0, "problems": [], "info": [], "refused": 0, "content_only": bool(errors) and not transport}   # content_only: retrying cannot help
    try:
        written = store.put(rows)
    except WarehouseUnavailable as e:
        detail["errors"] = [f"warehouse unavailable, {len(rows)} cells not stored: {e}", *errors][:20]
        detail["content_only"] = False
        f = Fill(job, f"{len(grids)} grids", started, round(time.monotonic() - t0, 2), 0, "error", detail)
        return FillResult(f, [])
    n, problems, info = check_batch(cat, rows, closed, store) if rows else (0, [], [])
    result = "error" if errors else ("warn" if problems or warnings else "ok")
    detail.update(checks=n, problems=problems[:20], info=info, refused=len({r.key for r in rows}) - written)
    f = Fill(job, f"{len(grids)} grids", started, round(time.monotonic() - t0, 2), written, result, detail)
    try:
        store.record_fill(f)
    except WarehouseUnavailable as e:
        log.warning("warehouse extract: the fill record was not stored: %s", e)
    return FillResult(f, rows)
