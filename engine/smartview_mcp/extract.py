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
from .warehouse import Fill, Row, WhKey

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


def finplan_scenarios(forecast: str | None) -> dict[tuple[str, str], str]:
    """(FinPlan scenario member, version) -> catalogue scenario key, for analytics.frame.from_rows."""
    out = {("Actual", "Working"): "actual_finplan", ("Target", "Final"): "budget"}
    if forecast:
        out[(forecast, "Final")] = "forecast"
    return out


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
            atlas_build: str = "") -> tuple[list[Row], list[str]]:
    """Warehouse rows for every data cell of a retrieved grid (parsed with suppression off), and the problems:
    cells that are not data (#NoAccess, #Invalid) are not stored."""
    pov = dict(lay.pov)
    rows, problems = [], []
    for r in parsed.rows:
        combo = r[:lay.nr]
        for j, ctuple in enumerate(lay.col_tuples):
            v = r[lay.nr + j]
            members = {**pov, **{a.dimension: x for a, x in zip(lay.rows, combo)},
                       **{a.dimension: x for a, x in zip(lay.cols, ctuple)}}
            key = wh_key(conn, members)
            if isinstance(v, str):
                problems.append(f"{' / '.join(combo)} / {' / '.join(ctuple)}: {v}")
                continue
            rows.append(Row(key, None if v is None else float(v), route, fetched_at, closed_label, atlas_build))
    return rows, problems


def _chunks(xs: list, n: int) -> list[list]:
    return [xs[i:i + n] for i in range(0, len(xs), n)]


def finmax_grids(cat: Catalog, months: list[str], entities: list[str], spot: str | None = None) -> list[Grid]:
    lines = cat.pnl["FinMax"]
    flows = [x.account for x in lines if x.measure == "flow"]
    stocks = [x.account for x in lines if x.measure == "stock"]
    out = []
    for ch in _chunks(months, CHUNK):
        out.append(Grid("FinMax", {**FINMAX_POV, "All Balance Types": "Periodic"},
                        [Axis("All Entities", entities), Axis("All Accounts", flows)],
                        [Axis("All Time Periods", ch)], f"FinMax Actual {ch[0]}..{ch[-1]} Periodic"))
    if stocks:
        levels = sorted(set(months) | {periods.opening(months[0])}, key=periods.index)
        out.append(Grid("FinMax", {**FINMAX_POV, "All Balance Types": "Accumulated"},
                        [Axis("All Entities", entities), Axis("All Accounts", stocks)],
                        [Axis("All Time Periods", levels)], f"FinMax Actual levels {levels[0]}..{levels[-1]}"))
    if spot:
        out.append(Grid("FinMax", {**FINMAX_POV, "All Balance Types": "Accumulated"},
                        [Axis("All Entities", entities), Axis("All Accounts", flows)],
                        [Axis("All Time Periods", [spot])], f"FinMax Actual YTD {spot} (months = YTD check)"))
    return out


def finplan_grids(cat: Catalog, plans: list[tuple[str, str, int]], entities: list[str]) -> list[Grid]:
    """plans: (scenario member, version, two-digit year); one grid of entities x lines x the 12 months each."""
    accts = [x.account for x in cat.pnl["FinPlan"]]
    return [Grid("FinPlan", {"Scenario": s, "Version": v, "Years": f"FY{yy:02d}", "All Balance Types": "Periodic"},
                 [Axis("Entity", entities), Axis("Account", accts)], [Axis("Period", list(periods.MONTHS))],
                 f"FinPlan {s}/{v} FY{yy:02d}") for s, v, yy in plans]


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
    error, the COM route through the hidden Excel when `host` (a function returning the ExcelHost) is given."""
    def retrieve(conn, pov, rows, cols):
        from . import svhttp
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
                                 timeout=config.CALL_TIMEOUT_S)
        return lay, p, secs, "com"
    return retrieve


@dataclass
class FillResult:
    fill: Fill
    rows: list[Row] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.fill.result != "error"


def check_batch(cat: Catalog, rows: list[Row], closed: str) -> tuple[int, list[str], list[str]]:
    """(checks run, problems, information) over the rows of one fill: the identity checks per entity and month in
    both cubes, the Periodic months against the Accumulated spot column, and the areas against the company total."""
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
    periodic = {k: v for k, v in data.items() if k[0] == "FinMax" and k[6] == "Periodic"}
    spot = {k: v for k, v in data.items() if k[0] == "FinMax" and k[6] == "Accumulated" and k[5] == closed
            and k[4] in {x.account for x in cat.pnl["FinMax"] if x.measure == "flow"}}
    k, p = reconcile.months_vs_ytd(periodic, spot, periods.ytd(closed))
    n, problems = n + k, problems + p
    company = next((e for e in cat.entities if e.key == "company"), None)
    others = [e.finmax for e in areas(cat, "FinMax") if e.key != "company"]
    if company and company.finmax:
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
    rows, errors, routes = [], [], set()
    for g in grids:
        try:
            lay, parsed, _secs, how = retrieve(g.conn, g.pov, g.rows, g.cols)
        except Exception as e:   # noqa: BLE001 - one failed grid must not lose the others
            errors.append(f"{g.what}: {type(e).__name__}: {str(e).splitlines()[0][:200] if str(e) else ''}")
            continue
        routes.add(how)
        got, bad = to_rows(lay, parsed, g.conn, how, now(), closed_label or closed, atlas_build)
        rows += got
        errors += [f"{g.what}: {b}" for b in bad[:5]] + [f"{g.what}: {w}" for w in parsed.pov_problems]
        errors += [f"{g.what}: unknown members {', '.join(parsed.unknown_members)}"] if parsed.unknown_members else []
    store.put(rows)
    n, problems, info = check_batch(cat, rows, closed) if rows else (0, [], [])
    result = "error" if errors else ("warn" if problems else "ok")
    detail = {"grids": len(grids), "routes": sorted(routes), "errors": errors[:20], "checks": n,
              "problems": problems[:20], "info": info}
    f = Fill(job, f"{len(grids)} grids", started, round(time.monotonic() - t0, 2), len(rows), result, detail)
    store.record_fill(f)
    return FillResult(f, rows)
