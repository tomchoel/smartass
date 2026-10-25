"""Catalogue KPIs: numerator / denominator over one cube's lines. basis average_stock divides by the average level
(opening December + each month end, n+1 points); annualise scales a part-year numerator to twelve months."""
from __future__ import annotations

from ..catalog import Kpi
from .facts import Fact, combine
from .frame import Frame
from .measures import amount, average_level


def kpi(frame: Frame, k: Kpi, entity: str, scenario: str, months: list[str], id: str | None = None) -> Fact:
    id = id or f"kpi.{k.key}.{scenario}.{months[-1]}"
    num = amount(frame, entity, k.numerator, scenario, months)
    if k.basis == "average_stock":
        den = average_level(frame, entity, k.denominator, scenario, months)
    else:
        den = amount(frame, entity, k.denominator, scenario, months)
    factor = 12 / len(months) if k.annualise else 1.0
    how = f"{k.numerator} / {k.denominator}" + (f" x 12/{len(months)}" if k.annualise else "")
    unit = "pct" if k.unit == "pct" else ("nok_per_fte" if k.basis == "average_stock" else "nok")
    return combine(id, unit, f"{k.label}: {how}; {den.definition}", lambda a, b: a / b * factor, num, den)
