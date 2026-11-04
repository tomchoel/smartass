"""Provider over v1's Smart View HTTP route (svhttp.SvHttp, read-only, ~0.2 s per grid once signed in).
FinMax/FinLite: rows Accounts x Entities, columns Time Periods ('Sep 26'), POV scenario, NOK, balance.
FinPlan: rows Account x Entity, columns Period ('Sep'), POV Scenario, Version, Years ('FY26'), balance - one grid
per year, periods relabelled to 'Sep 26'. #Missing gives no fact (missing is never zero); #NoAccess / #Invalid
give a warning and no fact. The ad hoc Excel fallback arrives in Plan 1C."""
from __future__ import annotations

import time
from dataclasses import dataclass, field

from smartview_mcp.grid import Axis

from ..model.frame import Fact
from ..model.periods import Month, parse
from ..model.requests import GridRequest

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
            out.append((pov, [Axis("Account", list(req.accounts)), Axis("Entity", list(req.entities))],
                        [Axis("Period", months)], {mo: f"{mo} {fy[2:]}" for mo in months}))
        return out
    pov = {"All Scenarios": req.scenario, "All Currencies": "NOK", "All Balance Types": req.balance}
    return [(pov, [Axis("All Accounts", list(req.accounts)), Axis("All Entities", list(req.entities))],
             [Axis("All Time Periods", list(req.periods))], {p: p for p in req.periods})]


class SvHttpProvider:
    def __init__(self, sv):
        self.sv = sv

    def fetch(self, req: GridRequest) -> FetchResult:
        out = FetchResult(req)
        t0 = time.monotonic()
        for pov, rows, cols, labels in _grids(req):
            _, parsed, _ = self.sv.retrieve(req.cube, pov, rows, cols, False, False)
            if parsed.unknown_members:
                out.warnings.append(f"{req.cube}: Smart View did not know {', '.join(parsed.unknown_members)}")
            heads = parsed.columns[2:]
            for row in parsed.rows:
                account, entity, values = row[0], row[1], row[2:]
                for head, v in zip(heads, values):
                    if v is None:
                        continue
                    if isinstance(v, str):
                        out.warnings.append(f"{req.cube} {account} / {entity} / {head}: {v} (left out)")
                        continue
                    out.facts.append(Fact(req.cube, req.query, account, entity, labels[head], req.scenario,
                                          req.version, float(v), req.balance))
        out.seconds = time.monotonic() - t0
        return out

    def closed_months(self) -> dict[str, Month]:
        """FinMax's 'Current Reporting Month' for every cube (FinLite and FinPlan close with FinMax)."""
        m = parse(self.sv.closed_month("FinMax"))
        return {"FinMax": m, "FinLite": m, "FinPlan": m}
