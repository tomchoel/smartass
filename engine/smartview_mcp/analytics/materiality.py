"""Materiality, the catalogue rule (owner decisions 2026-10-25): a variance is material at pct of |comparison| when
it is also at least pct_floor_nok, and/or (rule either | both) at abs_nok. A zero comparison has no %, so only the
NOK amount counts. Lines in `always` are always commented. Only NOK amounts are judged (FTE is not an amount)."""
from __future__ import annotations

from dataclasses import dataclass

from ..catalog import Materiality
from .variance import Variance


@dataclass(frozen=True)
class Flag:
    line: str
    material: bool
    always: bool
    reason: str
    size: float                       # |variance| (0 when unknown): the ranking key

    @property
    def comment(self) -> bool:
        return self.material or self.always


def _nok(v: float) -> str:
    return f"NOK {v:,.0f}".replace(",", " ")


def assess(v: Variance, m: Materiality) -> Flag:
    always = v.line in m.always
    a, c = v.abs.value, v.comparison.value
    if a is None:
        return Flag(v.line, False, always, f"variance unknown ({v.abs.reason})", 0.0)
    if v.abs.unit != "nok":
        return Flag(v.line, False, always, "not an amount: no NOK rule", abs(a))
    size = abs(a)
    base = abs(c) if c is not None else 0.0
    pct_hit = base > 0 and size >= m.pct * base and size >= m.pct_floor_nok
    abs_hit = size >= m.abs_nok
    material = (pct_hit or abs_hit) if m.rule == "either" else (pct_hit and abs_hit)
    why = []
    if pct_hit:
        why.append(f"{size / base:.0%} of the comparison (counted from {_nok(m.pct_floor_nok)})")
    if abs_hit:
        why.append(f"at least {_nok(m.abs_nok)}")
    if not why:
        why.append("below the thresholds" + ("; the comparison is zero: no %" if base == 0 else ""))
    return Flag(v.line, material, always, "; ".join(why), size)


def flags(variances: list[Variance], m: Materiality) -> list[Flag]:
    """One flag per variance, largest |variance| first (a stable sort keeps catalogue order among equals)."""
    return sorted((assess(v, m) for v in variances), key=lambda f: -f.size)
