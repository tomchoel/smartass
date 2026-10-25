"""Materiality, the catalogue rule (owner decisions 2026-10-25). On the YTD variance: material at pct of |comparison|
when that variance is also at least pct_floor_nok, and/or (rule either | both) at abs_nok. A comparison below
pct_floor_nok counts as zero (a small base makes a meaningless %): there is no % test, so only the NOK amount
counts under rule "either", and under rule "both" the YTD variance can never be material (the month swing still can). A single month whose
variance reaches abs_nok is flagged too (basis 'month'). Lines in `always` are always commented. Only NOK amounts are
judged (FTE is not an amount). On subtotals, flags() judges the leaf lines and the always-lines only."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Mapping

from ..catalog import Line, Materiality
from .variance import Variance

_EPS = 1e-6                                            # float noise at an exact threshold


@dataclass(frozen=True)
class Flag:
    line: str
    material: bool
    always: bool
    reason: str
    size: float                       # the larger of |YTD variance| and |month variance| (0 when unknown): the ranking key
    basis: str = ""                   # what made it material: 'ytd', 'month', 'ytd+month' or '' (not material)

    @property
    def comment(self) -> bool:
        return self.material or self.always


def _nok(v: float) -> str:
    return f"NOK {v:,.0f}".replace(",", " ")


def assess(v: Variance, m: Materiality, month: Variance | None = None) -> Flag:
    """The YTD variance `v` judged by the rule; `month` (the single-month variance of the same line) can make it
    material on its own at abs_nok."""
    always = v.line in m.always
    a, c = v.abs.value, v.comparison.value
    if a is None:
        return Flag(v.line, False, always, f"variance unknown ({v.abs.reason})", 0.0)
    if v.abs.unit != "nok":
        return Flag(v.line, False, always, "not an amount: no NOK rule", abs(a))
    size = abs(a)
    base = abs(c) if c is not None else 0.0
    small_base = base < m.pct_floor_nok or base <= 0
    pct_hit = (not small_base) and size + _EPS >= m.pct * base and size + _EPS >= m.pct_floor_nok
    abs_hit = size + _EPS >= m.abs_nok
    ytd_material = (pct_hit or abs_hit) if m.rule == "either" else (pct_hit and abs_hit)
    swing = month.abs.value if month is not None and month.abs.unit == "nok" else None
    month_hit = swing is not None and abs(swing) + _EPS >= m.abs_nok
    material = ytd_material or month_hit
    basis = "ytd+month" if ytd_material and month_hit else "ytd" if ytd_material else "month" if month_hit else ""
    why = []
    if ytd_material:
        if pct_hit:
            why.append(f"{size / base:.0%} of the comparison (counted from {_nok(m.pct_floor_nok)})")
        if abs_hit:
            why.append(f"at least {_nok(m.abs_nok)}")
    if month_hit:
        why.append(f"a single month moved {_nok(abs(swing))} (at least {_nok(m.abs_nok)})")
    if not why:
        why.append("below the thresholds" + (f"; the comparison is under {_nok(m.pct_floor_nok)}: counts as zero, "
                                             "no %" if small_base else ""))
    rank = max(size, abs(swing)) if swing is not None else size
    return Flag(v.line, material, always, "; ".join(why), rank, basis)


def flags(variances: list[Variance], m: Materiality, months: Mapping[str, Variance] | None = None,
          lines: Iterable[Line] | None = None) -> list[Flag]:
    """One flag per variance, largest |variance| first (a stable sort keeps catalogue order among equals).
    months: the single-month variance per line key. lines: when given, only leaf lines (kind 'line') and the
    always-lines are judged - a subtotal (gm, opex, boi) is not flagged again for what its parts already show."""
    leaf = None if lines is None else {x.key for x in lines if x.kind == "line"}
    months = months or {}
    keep = (x for x in variances if leaf is None or x.line in leaf or x.line in m.always)
    return sorted((assess(x, m, months.get(x.line)) for x in keep), key=lambda f: -f.size)
