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
    request: GridRequest
    facts: list[Fact] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    route: str = "http"
    seconds: float = 0.0
    fallback: str = ""          # route "rest": the provider route's error that sent the request to the second engine
    dropped: list[tuple[str, str, str]] = field(default_factory=list)   # (cube, dimension, member) left out over REST


def grids(req: GridRequest) -> list[tuple[dict, list[Axis], list[Axis], dict[str, str]]]:
    """(pov, rows, cols, {column member: 'Sep 26' label}) per grid, as both engines fetch them."""
    if req.cube == "FinPlan":
        by_year: dict[str, list[str]] = {}
        for p in req.periods:
            y, m = parse(p)
            by_year.setdefault(f"FY{y % 100:02d}", []).append(MONTHS[m - 1])
        out = []
        for fy, months in by_year.items():
            pov = {"Scenario": req.scenario, "Version": req.version, "Years": fy, "All Balance Types": req.balance}
            out.append((pov, [Axis("Account", list(req.accounts)), Axis("Entity", list(req.entities)), *_products(req)],
                        [Axis("Period", months)], {mo: f"{mo} {fy[2:]}" for mo in months}))
        return out
    pov = {"All Scenarios": req.scenario, "All Currencies": "NOK", "All Balance Types": req.balance}
    return [(pov, [Axis("All Accounts", list(req.accounts)), Axis("All Entities", list(req.entities)), *_products(req)],
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
    nr = 3 if req.products else 2                   # row dimensions: account, entity (, product)
    heads = parsed.columns[nr:]
    for row in parsed.rows:
        account, entity, values = row[0], row[1], row[nr:]
        product = row[2] if req.products else ""
        for head, v in zip(heads, values):
            if v is None:
                continue
            if isinstance(v, str):
                where = " / ".join(x for x in (account, entity, product, head) if x)
                warnings.append(f"{req.cube} {where}: {v} (left out)")
                continue
            facts.append(Fact(req.cube, req.query, account, entity, labels[head], req.scenario, req.version,
                              float(v), req.balance, product))
    return facts, warnings


class SvHttpProvider:
    def __init__(self, sv):
        self.sv = sv

    def _retrieve(self, out: FetchResult, *args):
        """sv.retrieve; a lost request (cold sign-in) is tried once more after RETRY_WAIT_S, anything else is raised."""
        try:
            return self.sv.retrieve(*args)
        except SvHttpError as e:
            if "lost the request" not in str(e):
                raise
        time.sleep(RETRY_WAIT_S)
        out.warnings.append(f"{args[0]}: the first request was lost (cold sign-in) and retried")
        return self.sv.retrieve(*args)

    def fetch(self, req: GridRequest) -> FetchResult:
        t0 = time.monotonic()
        try:
            out = FetchResult(req)
            for pov, rows, cols, labels in grids(req):
                _, parsed, _ = self._retrieve(out, req.cube, pov, rows, cols, False, False)
                facts, warns = facts_from_parsed(req, parsed, labels)
                out.facts += facts
                out.warnings += warns
        except NoFallback:
            raise
        except SvHttpError as e:
            out = self._over_rest(req, e)
        out.seconds = time.monotonic() - t0
        return out

    def _over_rest(self, req: GridRequest, first: SvHttpError) -> FetchResult:
        """The request's grids over the second engine; a member EPM refuses there is left out of all of them and the
        request fetched again (at most MAX_LEFT_OUT). If it cannot be fetched, the provider route's error with REST's
        after it."""
        left: list[tuple[str, str]] = []              # (dimension, member) left out, in the order EPM named them
        while True:
            try:
                return self._rest_round(req, first, left)
            except MemberRefused as e:
                if e.droppable and len(left) < MAX_LEFT_OUT:
                    left.append((e.dimension, e.member))
                    continue
                why = e if not e.droppable else SvHttpError(
                    f"Planning REST refused more than {MAX_LEFT_OUT} members of the {req.cube} grid: "
                    f"{', '.join(m for _, m in left + [(e.dimension, e.member)])} - {ATLAS_BUILD}")
            except SvHttpError as e:
                why = e
            raise SvHttpError(f"{first} - the second engine (Planning REST) could not fetch it either (query "
                              f"{req.query}): {why}") from None

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
