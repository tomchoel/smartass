# engine/smartass/model/series.py
"""A query's series (value, ly, budget, forecast) as scenario, version and months. Actual vs plan lives entirely in
FinPlan: budget and forecast exist only there (FinPlan actuals are version Working, plans version Final)."""
from __future__ import annotations

from .periods import Month, resolve, shift
from .resolver import Resolver
from .spec import Query


def scenario(cube: str, series: str, r: Resolver) -> tuple[str, str]:
    if cube == "FinPlan":
        table = {"value": ("Actual", "Working"), "ly": ("Actual", "Working"), "budget": ("Target", "Final")}
        if series in table:
            return table[series]
        if series == "forecast":
            return (r.latest_forecast(), "Final")
        raise ValueError(f"unknown series {series!r}")
    if series not in ("value", "ly", "budget", "forecast"):
        raise ValueError(f"unknown series {series!r}")
    if series in ("budget", "forecast"):
        raise ValueError(f"{series} exists only in FinPlan")
    return ("Actual", "")


def months(q: Query, series: str, closed: Month, base: list[Month] | None = None) -> list[Month]:
    ms = list(base) if base is not None else resolve(q.periods, closed)
    return [shift(m, -12) for m in ms] if series == "ly" else ms
