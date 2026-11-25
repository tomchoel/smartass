"""Provider over the Smart View provider's HTTP route (svhttp.SvHttp, read-only, ~0.2 s per grid once signed in).
FinMax/FinLite: rows Accounts x Entities, columns Time Periods ('Sep 26'), POV scenario, NOK, balance.
FinPlan: rows Account x Entity, columns Period ('Sep'), POV Scenario, Version, Years ('FY26'), balance - one grid
per year, periods relabelled to 'Sep 26'. A request with products adds 'All Products' (the dimension's name in all
three cubes) as a third row dimension. #Missing gives no fact (missing is never zero); #NoAccess / #Invalid
give a warning and no fact.

A request the provider route could not fetch is fetched once over the second engine (epm/rest.py, Planning REST:
svhttp.SvHttp.fetch_rest), the same grids parsed the same way - unless its error is one REST would meet as well
(svhttp.NoFallback: the token, the scope, the atlas, an alias clash) or a missing sign-in. Its FetchResult says route
"rest" and why (fallback: the provider route's error).

A member EPM refuses over REST (rest.MemberRefused: it does not exist there, or the user may not read it) on an axis of
several is left out of every grid of the request, and the request is fetched again without it - at most MAX_LEFT_OUT
members, one per round; the facts are then exactly those of the request without it, the member has none (the path of
a missing cell), and FetchResult.dropped names it, so the reply says its figures are missing (pipeline.prepare). A POV
member, an axis's only member or more than MAX_LEFT_OUT members fail the request with a clear error."""
from __future__ import annotations

import time
from dataclasses import dataclass, field

from ..epm.grid import Axis, parse as parse_grid
from ..epm.rest import ATLAS_BUILD, MemberRefused
from ..epm.svhttp import NoFallback, SvHttpError

from ..model.frame import Fact
from ..model.periods import Month, parse
from ..model.requests import GridRequest

RETRY_WAIT_S = 5.0     # a request lost on the way (svhttp.RequestLost) is retried once after this wait
MAX_LEFT_OUT = 3       # members EPM refuses over REST that one request is fetched without, one per round

MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


@dataclass
class FetchResult:
    request: GridRequest | None
    facts: list[Fact] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    route: str = "http"
    seconds: float = 0.0
    fallback: str = ""          # route "rest": the provider route's error that sent the request to the second engine
    dropped: list[tuple[str, str, str]] = field(default_factory=list)   # (cube, dimension, member) left out over REST


@dataclass
class SliceFetchResult:
    cube: str
    grids: list[tuple[object, object]] = field(default_factory=list)   # (Parsed, caller metadata)
    warnings: list[str] = field(default_factory=list)
    route: str = "http"
    seconds: float = 0.0
    fallback: str = ""
    dropped: list[tuple[str, str, str]] = field(default_factory=list)


def _extra_axes(req: GridRequest) -> tuple[dict[str, str], list[Axis], list[tuple[str, str]]]:
    """Split req.dimensions into POV overrides (1 member) and extra row Axes (>1 members)."""
    pov_over: dict[str, str] = {}
    axes: list[Axis] = []
    single_dims: list[tuple[str, str]] = []
    for dim, ms in req.dimensions:
        if len(ms) == 1:
            pov_over[dim] = ms[0]
            single_dims.append((dim, ms[0]))
        elif len(ms) > 1:
            axes.append(Axis(dim, list(ms)))
    return pov_over, axes, single_dims


def grids(req: GridRequest) -> list[tuple[dict, list[Axis], list[Axis], dict[str, str]]]:
    """(pov, rows, cols, {column member: 'Sep 26' label}) per grid, as both engines fetch them."""
    pov_over, extra_axes, _ = _extra_axes(req)
    if req.cube == "FinPlan":
        by_year: dict[str, list[str]] = {}
        for p in req.periods:
            y, m = parse(p)
            by_year.setdefault(f"FY{y % 100:02d}", []).append(MONTHS[m - 1])
        out = []
        for fy, months in by_year.items():
            pov = {"Scenario": req.scenario, "Version": req.version, "Years": fy, "All Balance Types": req.balance,
                   **pov_over}
            out.append((pov, [Axis("Account", list(req.accounts)), Axis("Entity", list(req.entities)),
                              *_products(req), *extra_axes],
                        [Axis("Period", months)], {mo: f"{mo} {fy[2:]}" for mo in months}))
        return out
    pov = {"All Scenarios": req.scenario, "All Currencies": "NOK", "All Balance Types": req.balance, **pov_over}
    return [(pov, [Axis("All Accounts", list(req.accounts)), Axis("All Entities", list(req.entities)),
                   *_products(req), *extra_axes],
             [Axis("All Time Periods", list(req.periods))], {p: p for p in req.periods})]


PRODUCT_DIMENSION = "All Products"         # FinMax, FinLite and FinPlan (atlas, 04.11.2026)


def _products(req: GridRequest) -> list[Axis]:
    return [Axis(PRODUCT_DIMENSION, list(req.products))] if req.products else []


def facts_from_parsed(req: GridRequest, parsed, labels: dict[str, str]) -> tuple[list[Fact], list[str]]:
    """Facts and warnings from one parsed grid (grid.parse output), the same for both engines.
    #Missing (None) gives no fact; an error cell gives a warning and no fact."""
    facts: list[Fact] = []
    warnings: list[str] = []
    if parsed.unknown_members:
        warnings.append(f"{req.cube}: Smart View did not know {', '.join(parsed.unknown_members)}")
    elif getattr(parsed, "empty", False):
        what = ", ".join(req.accounts[:3]) + (" …" if len(req.accounts) > 3 else "")
        warnings.append(f"{req.cube}: no data for {what} in {req.scenario} ({', '.join(req.entities[:2])}; "
                        f"{req.periods[0]}–{req.periods[-1]})" if req.periods else f"{req.cube}: no data for {what}")
    _, extra_axes, single_dims = _extra_axes(req)
    extra_dim_names = [a.dimension for a in extra_axes]
    nr = (3 if req.products else 2) + len(extra_axes)
    heads = parsed.columns[nr:]
    for row in parsed.rows:
        account, entity = row[0], row[1]
        product = row[2] if req.products else ""
        extra_start = 3 if req.products else 2
        row_dims = tuple(sorted([*single_dims,
                                 *((d, row[extra_start + idx]) for idx, d in enumerate(extra_dim_names))]))
        values = row[nr:]
        for head, v in zip(heads, values):
            if v is None:
                continue
            if isinstance(v, str):
                where = " / ".join(x for x in (account, entity, product, head) if x)
                warnings.append(f"{req.cube} {where}: {v} (left out)")
                continue
            facts.append(Fact(req.cube, req.query, account, entity, labels[head], req.scenario, req.version,
                              float(v), req.balance, product, dimensions=row_dims))
    return facts, warnings


class SvHttpProvider:
    def __init__(self, sv):
        self.sv = sv

    def _run_http(self, cube: str, specs: list[tuple[dict, list[Axis], list[Axis], object]],
                  suppress_missing: bool, suppress_zeros: bool) -> tuple[list[tuple[object, object]], list[str]]:
        out_parsed, out_warnings = [], []
        for pov, rows, cols, meta in specs:
            _, parsed, _ = self._retrieve(out_warnings, cube, pov, rows, cols, suppress_missing, suppress_zeros)
            out_parsed.append((parsed, meta))
        return out_parsed, out_warnings

    def _retrieve(self, warnings_target, *args):
        """sv.retrieve; a lost request (cold sign-in) is tried once more after RETRY_WAIT_S, anything else is raised."""
        try:
            return self.sv.retrieve(*args)
        except SvHttpError as e:
            if "lost the request" not in str(e):
                raise
        time.sleep(RETRY_WAIT_S)
        msg = f"{args[0]}: the first request was lost (cold sign-in) and retried"
        if isinstance(warnings_target, list):
            warnings_target.append(msg)
        else:
            warnings_target.warnings.append(msg)
        return self.sv.retrieve(*args)

    def fetch(self, req: GridRequest) -> FetchResult:
        t0 = time.monotonic()
        specs = grids(req)
        try:
            parsed_pairs, warns = self._run_http(req.cube, specs, False, False)
            out = FetchResult(req, warnings=warns)
            for parsed, labels in parsed_pairs:
                facts, fw = facts_from_parsed(req, parsed, labels)
                out.facts += facts
                out.warnings += fw
        except NoFallback:
            raise
        except SvHttpError as e:
            out = self._over_rest(req, e)
        out.seconds = time.monotonic() - t0
        return out

    def fetch_slice(self, cube: str, specs: list[tuple[dict, list[Axis], list[Axis], object]],
                    *, suppress_missing: bool = True, suppress_zeros: bool = False,
                    label: str = "slice") -> SliceFetchResult:
        """Fetch arbitrary multi-dimensional grids via the same Smart View + REST dual engine."""
        t0 = time.monotonic()
        try:
            parsed_pairs, warns = self._run_http(cube, specs, suppress_missing, suppress_zeros)
            out = SliceFetchResult(cube, grids=parsed_pairs, warnings=warns)
        except NoFallback:
            raise
        except SvHttpError as e:
            out = self._slice_over_rest(cube, specs, label, e, suppress_missing, suppress_zeros)
        out.seconds = time.monotonic() - t0
        return out

    def _over_rest_loop(self, cube: str, label: str, first: SvHttpError, run_once):
        """Shared REST retry loop leaving out MemberRefused members up to MAX_LEFT_OUT."""
        left: list[tuple[str, str]] = []
        while True:
            try:
                return run_once(left)
            except MemberRefused as e:
                if e.droppable and len(left) < MAX_LEFT_OUT:
                    left.append((e.dimension, e.member))
                    continue
                why = e if not e.droppable else SvHttpError(
                    f"Planning REST refused more than {MAX_LEFT_OUT} members of the {cube} grid: "
                    f"{', '.join(m for _, m in left + [(e.dimension, e.member)])} - {ATLAS_BUILD}")
            except SvHttpError as e:
                why = e
            raise SvHttpError(f"{first} - the second engine (Planning REST) could not fetch it either ({label}): "
                              f"{why}") from None

    def _over_rest(self, req: GridRequest, first: SvHttpError) -> FetchResult:
        """The request's grids over the second engine; a member EPM refuses there is left out of all of them and the
        request fetched again (at most MAX_LEFT_OUT)."""
        return self._over_rest_loop(req.cube, f"query {req.query}", first,
                                    lambda left: self._rest_round(req, first, left))

    def _slice_over_rest(self, cube: str, specs, label: str, first: SvHttpError,
                         suppress_missing: bool, suppress_zeros: bool) -> SliceFetchResult:
        def round_once(left):
            out = SliceFetchResult(cube, route="rest", fallback=str(first),
                                   dropped=[(cube, d, m) for d, m in left])
            for pov, rows, cols, meta in specs:
                f = self.sv.fetch_rest(cube, pov, _without(rows, left), _without(cols, left))
                parsed = parse_grid(f.layout, f.values, suppress_missing, suppress_zeros)
                out.grids.append((parsed, meta))
                out.warnings += f.warnings
            return out
        return self._over_rest_loop(cube, label, first, round_once)

    def _rest_round(self, req: GridRequest, first: SvHttpError, left: list[tuple[str, str]]) -> FetchResult:
        out = FetchResult(req, route="rest", fallback=str(first), dropped=[(req.cube, d, m) for d, m in left])
        for pov, rows, cols, labels in grids(req):
            f = self.sv.fetch_rest(req.cube, pov, _without(rows, left), _without(cols, left))
            facts, warns = facts_from_parsed(req, parse_grid(f.layout, f.values, False, False), labels)
            out.facts += facts
            out.warnings += f.warnings + warns
        return out

    def closed_months(self) -> dict[str, Month]:
        """FinMax's 'Current Reporting Month' for every cube (FinLite and FinPlan close with FinMax)."""
        m = parse(self.sv.closed_month("FinMax"))
        return {"FinMax": m, "FinLite": m, "FinPlan": m}


def _without(axes: list[Axis], left: list[tuple[str, str]]) -> list[Axis]:
    """The axes without the members left out - an axis keeps its members when none would remain (one year's grid of
    a FinPlan request may hold only that member there): EPM then refuses it again, as a member that cannot be left
    out."""
    out = []
    for a in axes:
        kept = [m for m in a.members if (a.dimension, m) not in left]
        out.append(Axis(a.dimension, kept or list(a.members)))
    return out
