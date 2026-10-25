"""Variance with the house rule 'positive = better'. better = 'higher' (every P&L line: FinMax costs are negative):
current - comparison. better = 'lower' (the positive cost-per-FTE KPIs): comparison - current, so a rise is negative.
better = 'none' (FTE): current - comparison, no judgement. The % is variance / |comparison|; a zero comparison has
no %."""
from __future__ import annotations

from dataclasses import dataclass

from ..catalog import Kpi, Line
from .facts import Fact, combine, missing


@dataclass(frozen=True)
class Variance:
    line: str
    current: Fact
    comparison: Fact
    abs: Fact
    pct: Fact
    favourable: bool | None           # None: unknown, zero, or the line has no better direction (FTE)


def variance(line: "Line | Kpi", current: Fact, comparison: Fact, id: str) -> Variance:
    lower = line.better == "lower"
    a = combine(f"{id}.var", current.unit,
                f"{comparison.id} - {current.id} (comparison - current: lower is better)" if lower
                else f"{current.id} - {comparison.id}",
                (lambda x, y: y - x) if lower else (lambda x, y: x - y), current, comparison)
    if a.value is None:
        pct = missing(f"{id}.pct", "pct", "variance / |comparison|", a.reason or "no variance", a.inputs)
    elif comparison.value == 0:
        pct = missing(f"{id}.pct", "pct", "variance / |comparison|", "the comparison is zero: no %", a.inputs)
    else:
        pct = combine(f"{id}.pct", "pct", f"{a.id} / |{comparison.id}|", lambda v, c: v / abs(c), a, comparison)
    fav = None
    if a.value is not None and a.value != 0 and line.better != "none":
        fav = a.value > 0                                    # the sign already carries the direction
    return Variance(line.key, current, comparison, a, pct, fav)
