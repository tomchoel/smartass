"""Fetched facts as one long table: the same rows become the Excel table Data_Q<n>.
Every cube's periods are normalised to month labels ('Sep 26') by the data layer; FinMax has no version ('')."""
from __future__ import annotations

from dataclasses import dataclass

COLUMNS = ("Cube", "Query", "Account", "Entity", "Period", "Scenario", "Version", "Balance", "Value")
FILTER_COLUMNS = COLUMNS[:-1]


@dataclass(frozen=True)
class Fact:
    cube: str
    query: str
    account: str
    entity: str
    period: str
    scenario: str
    version: str
    value: float
    balance: str = "Periodic"


class Frame:
    def __init__(self, facts=()):
        self.facts: list[Fact] = list(facts)

    def total(self, filters: dict[str, tuple[str, ...]]) -> float | None:
        """Sum of the facts matching every filter (column -> allowed values); None when no fact matches."""
        bad = set(filters) - set(FILTER_COLUMNS)
        if bad:
            raise KeyError(f"unknown filter column(s): {sorted(bad)}")
        for col, allowed in filters.items():
            if isinstance(allowed, str):
                raise TypeError(f"filter {col}: give a tuple of values, not a string")
        hit = [x for x in self.facts
               if all(getattr(x, col.lower()) in allowed for col, allowed in filters.items())]
        return sum(x.value for x in hit) if hit else None

    def rows(self) -> list[tuple]:
        return [(x.cube, x.query, x.account, x.entity, x.period, x.scenario, x.version, x.balance, x.value)
                for x in self.facts]
