"""Variance = current - comparison, and % of |comparison|. FinMax signs: costs are negative, so for a line with
better = 'higher' (every P&L line) a positive variance is better. A zero comparison has no %."""
from __future__ import annotations

from dataclasses import dataclass

from ..catalog import Line
from .facts import Fact, combine, missing


@dataclass(frozen=True)
class Variance:
    line: str
    current: Fact
    comparison: Fact
    abs: Fact
    pct: Fact
    favourable: bool | None           # None: unknown, zero, or the line has no better direction (FTE)


def variance(line: Line, current: Fact, comparison: Fact, id: str) -> Variance:
    a = combine(f"{id}.var", current.unit, f"{current.id} - {comparison.id}", lambda x, y: x - y,
                current, comparison)
    if a.value is None:
        pct = missing(f"{id}.pct", "pct", "variance / |comparison|", a.reason or "no variance", a.inputs)
    elif comparison.value == 0:
        pct = missing(f"{id}.pct", "pct", "variance / |comparison|", "the comparison is zero: no %", a.inputs)
    else:
        pct = combine(f"{id}.pct", "pct", f"{a.id} / |{comparison.id}|", lambda v, c: v / abs(c), a, comparison)
    fav = None
    if a.value is not None and a.value != 0 and line.better != "none":
        fav = (a.value > 0) if line.better == "higher" else (a.value < 0)
    return Variance(line.key, current, comparison, a, pct, fav)
