"""Facts: every number the calc engine produces, with what it means and which warehouse cells it came from, so any
figure in an answer can be traced. A number that cannot be computed is a Fact with value None and a reason - never a
silent gap, never a division by zero."""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Callable, Iterable

UNITS = ("nok", "fte", "pct", "nok_per_fte", "count")


@dataclass(frozen=True)
class Fact:
    id: str
    value: float | None
    unit: str                         # one of UNITS; pct is a ratio (0.1 = 10 %)
    definition: str
    inputs: tuple[tuple, ...] = ()    # warehouse keys (cube, scenario, version, entity, account, period, balance)
    reason: str | None = None         # why value is None
    parts: tuple["Fact", ...] = ()    # the facts this one was combined from (a KPI: numerator, denominator)

    # Cube values follow FinMax (costs are negative); a KPI with sign = -1 is shown positive (cost per FTE).

    def __post_init__(self):
        if self.unit not in UNITS:
            raise ValueError(f"fact '{self.id}': unit '{self.unit}' is not one of {UNITS}")

    def as_dict(self, max_inputs: int = 40) -> dict:
        out = {"id": self.id, "value": self.value, "unit": self.unit, "definition": self.definition,
               "inputs": [list(k) for k in self.inputs[:max_inputs]]}
        if len(self.inputs) > max_inputs:
            out["inputs_more"] = len(self.inputs) - max_inputs
        if self.reason:
            out["reason"] = self.reason
        return out


def missing(id: str, unit: str, definition: str, reason: str, inputs: Iterable[tuple] = ()) -> Fact:
    return Fact(id, None, unit, definition, tuple(inputs), reason)


def _inputs(facts: Iterable[Fact]) -> tuple[tuple, ...]:
    return tuple(dict.fromkeys(k for f in facts for k in f.inputs))


def total(id: str, facts: list[Fact], unit: str, definition: str) -> Fact:
    """The sum of facts (0 for none); missing if any part is."""
    gone = next((f for f in facts if f.value is None), None)
    if gone is not None:
        return Fact(id, None, unit, definition, _inputs(facts), f"{gone.id}: {gone.reason}", tuple(facts))
    return Fact(id, float(sum(f.value for f in facts)), unit, definition, _inputs(facts), None, tuple(facts))


def combine(id: str, unit: str, definition: str, fn: Callable[..., float], *facts: Fact) -> Fact:
    """fn over the facts' values; missing when a part is missing or fn divides by zero."""
    gone = next((f for f in facts if f.value is None), None)
    if gone is not None:
        return Fact(id, None, unit, definition, _inputs(facts), f"{gone.id}: {gone.reason}", tuple(facts))
    try:
        value = float(fn(*(f.value for f in facts)))
    except ZeroDivisionError:
        return Fact(id, None, unit, definition, _inputs(facts), "division by zero", tuple(facts))
    if not math.isfinite(value):
        return Fact(id, None, unit, definition, _inputs(facts), "non-finite value", tuple(facts))
    return Fact(id, value, unit, definition, _inputs(facts), None, tuple(facts))
