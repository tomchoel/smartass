"""Fetched facts as one long table: the same rows become the Excel table Data_Q<n>.
Every cube's periods are normalised to month labels ('Sep 26') by the data layer; FinMax has no version ('').
A fact's product is '' when the query asked for none (the cube's product total); the Data table has a Product
column only when some fact has a product, so answers without products keep their table exactly."""
from __future__ import annotations

from dataclasses import dataclass

COLUMNS = ("Cube", "Query", "Account", "Entity", "Period", "Scenario", "Version", "Balance", "Value")
PRODUCT_COLUMNS = ("Cube", "Query", "Account", "Entity", "Product", "Period", "Scenario", "Version", "Balance",
                   "Value")
FILTER_COLUMNS = PRODUCT_COLUMNS[:-1]


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
    product: str = ""
    dimensions: tuple[tuple[str, str], ...] = ()


def fact_value(fact: Fact, col: str) -> str:
    """Return any filter column's value from a Fact (standard fields or extra dimensions)."""
    low = col.lower()
    if low in ("cube", "query", "account", "entity", "product", "period", "scenario", "version", "balance"):
        return getattr(fact, low)
    return dict(fact.dimensions).get(col, "")


class Frame:
    def __init__(self, facts=()):
        self.facts: list[Fact] = list(facts)
        self.lines: dict = {}                    # alias -> LineSet (ledger lines behind a cube cell)

    def dim_names(self) -> tuple[str, ...]:
        names = {d for x in self.facts for d, _ in x.dimensions}
        return tuple(sorted(names))

    def filter_columns(self) -> tuple[str, ...]:
        extra = self.dim_names()
        return (*FILTER_COLUMNS, *extra) if extra else FILTER_COLUMNS

    def total(self, filters: dict[str, tuple[str, ...]]) -> float | None:
        """Sum of the facts matching every filter (column -> allowed values); None when no fact matches."""
        bad = set(filters) - set(self.filter_columns())
        if bad:
            raise KeyError(f"unknown filter column(s): {sorted(bad)}")
        for col, allowed in filters.items():
            if isinstance(allowed, str):
                raise TypeError(f"filter {col}: give a tuple of values, not a string")
        hit = [x for x in self.facts
               if all(fact_value(x, col) in allowed for col, allowed in filters.items())]
        return sum(x.value for x in hit) if hit else None

    def has_products(self) -> bool:
        return any(x.product for x in self.facts)

    def columns(self) -> tuple[str, ...]:
        base = PRODUCT_COLUMNS[:-1] if self.has_products() else COLUMNS[:-1]
        extra = self.dim_names()
        return (*base, *extra, "Value") if extra else (PRODUCT_COLUMNS if self.has_products() else COLUMNS)

    def rows(self) -> list[tuple]:
        fcols = self.columns()[:-1]
        return [(*(fact_value(x, c) for c in fcols), x.value) for x in self.facts]
