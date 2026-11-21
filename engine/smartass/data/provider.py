"""Provider over v1's Smart View HTTP route (svhttp.SvHttp, read-only, ~0.2 s per grid once signed in).
FinMax/FinLite: rows Accounts x Entities, columns Time Periods ('Sep 26'), POV scenario, NOK, balance.
FinPlan: rows Account x Entity, columns Period ('Sep'), POV Scenario, Version, Years ('FY26'), balance - one grid
per year, periods relabelled to 'Sep 26'. A request with products adds 'All Products' (the dimension's name in all
three cubes) as a third row dimension. #Missing gives no fact (missing is never zero); #NoAccess / #Invalid
give a warning and no fact. The ad hoc Excel fallback arrives in Plan 1C."""
from __future__ import annotations

import time
from dataclasses import dataclass, field

from ..epm.grid import Axis
from ..epm.svhttp import SvHttpError

from ..model.frame import Fact
from ..model.periods import Month, parse
from ..model.requests import GridRequest

RETRY_WAIT_S = 5.0     # a request lost on the way (svhttp.RequestLost) is retried once after this wait

MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


@dataclass
class FetchResult:
    request: GridRequest
    facts: list[Fact] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    route: str = "http"
    seconds: float = 0.0


def _grids(req: GridRequest) -> list[tuple[dict, list[Axis], list[Axis], dict[str, str]]]:
    """(pov, rows, cols, {column member: 'Sep 26' label}) per grid."""
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
    """Facts and warnings from one parsed grid (grid.parse output), the same for the HTTP and the ad hoc route.
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
        out = FetchResult(req)
        t0 = time.monotonic()
        for pov, rows, cols, labels in _grids(req):
            _, parsed, _ = self._retrieve(out, req.cube, pov, rows, cols, False, False)
            facts, warns = facts_from_parsed(req, parsed, labels)
            out.facts += facts
            out.warnings += warns
        out.seconds = time.monotonic() - t0
        return out

    def closed_months(self) -> dict[str, Month]:
        """FinMax's 'Current Reporting Month' for every cube (FinLite and FinPlan close with FinMax)."""
        m = parse(self.sv.closed_month("FinMax"))
        return {"FinMax": m, "FinLite": m, "FinPlan": m}


class ComboProvider:
    """Fixed results (the ad hoc fallback's) for the requests that failed over HTTP, the primary provider for the rest."""

    def __init__(self, primary, fixed: dict):
        self.primary, self.fixed = primary, fixed

    def closed_months(self):
        return self.primary.closed_months()

    def fetch(self, req: GridRequest) -> FetchResult:
        return self.fixed[req] if req in self.fixed else self.primary.fetch(req)
