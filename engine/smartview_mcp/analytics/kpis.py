"""Catalogue KPIs: numerator / denominator over one cube's lines. basis average_stock divides by the average level
(the month before the window + each month end, n+1 points; for a YTD window that is the opening December); annualise
scales a part-year numerator to twelve months. The cube sign is FinMax's (costs negative); a KPI with sign = -1 is shown positive (cost per FTE)."""
from __future__ import annotations

from ..catalog import Kpi
from .facts import UNITS, Fact, combine, missing
from .frame import Frame
from .measures import _span, amount, average_level


def _num(v: float | None) -> str:
    return "?" if v is None else (f"{v:.0f}" if abs(v) >= 1000 else f"{v:.1f}")


def kpi(frame: Frame, k: Kpi, entity: str, scenario: str, months: list[str], id: str | None = None) -> Fact:
    id = id or f"kpi.{k.key}.{scenario}.{entity}.{_span(months)}"
    if k.unit not in UNITS:
        raise ValueError(f"kpi '{k.key}': unit '{k.unit}' is not one of {UNITS}")
    unit = k.unit
    if not months:
        return missing(id, unit, f"{k.label}: no months", "no months")
    num = amount(frame, entity, k.numerator, scenario, months)
    if k.basis == "average_stock":
        den = average_level(frame, entity, k.denominator, scenario, months)
        den_text = f"avg {k.denominator} {_num(den.value)} ({len(months) + 1} levels)"
    else:
        den = amount(frame, entity, k.denominator, scenario, months)
        den_text = f"{k.denominator} {_num(den.value)}"
    factor = 12 / len(months) if k.annualise else 1.0
    how = f"{k.numerator} {_span(months)} {_num(num.value)} / {den_text}" + (
        f" x 12/{len(months)}" if k.annualise else "")
    shown = " (shown positive: cost per FTE)" if k.sign == -1 else ""
    return combine(id, unit, f"{k.label}: {how}{shown}; {den.definition}", lambda a, b: a / b * factor * k.sign,
                   num, den)
