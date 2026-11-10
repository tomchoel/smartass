# engine/smartass/model/series.py
"""A query's series (value, ly, budget, forecast) as scenario, version and months. Actual vs plan lives entirely in
FinPlan: budget and forecast exist only there (FinPlan actuals are version Working, plans version Final).
The table lives in the `scenarios` rule (knowledge/semantic/rules)."""
from __future__ import annotations

from .periods import Month, resolve, shift
from .resolver import Resolver
from .spec import Query


PLAN_SERIES = ("budget", "forecast")


def scenario(cube: str, series: str, r: Resolver) -> tuple[str, str]:
    """(scenario, version) of a series in a cube, from the `scenarios` rule (knowledge/semantic/rules)."""
    if series not in ("value", "ly", "budget", "forecast"):
        raise ValueError(f"unknown series {series!r}")
    table = r.scenarios().get(cube, {})
    if series not in table:
        raise ValueError(f"{series} exists only in FinPlan")
    scen, version = table[series]
    return (r.latest_forecast() if scen == "latest_forecast" else scen, version)


def months(q: Query, series: str, closed: Month, base: list[Month] | None = None) -> list[Month]:
    if base is not None:
        ms = list(base)
    else:
        ms = resolve(q.plan_periods if series in PLAN_SERIES and q.plan_periods else q.periods, closed)
    return [shift(m, -12) for m in ms] if series == "ly" else ms
