# engine/smartass/model/resolver.py
"""The model's only view of knowledge. Sub-spec 2 implements Resolver over knowledge.db; StaticResolver serves
tests and seeds."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Protocol

Sign = Literal["income", "cost", "neutral"]
Status = Literal["approved", "draft"]
Balance = Literal["Periodic", "Accumulated"]
Aggregation = Literal["sum", "last"]
DEFAULT_MATERIALITY = (5_000_000.0, 0.05)


class UnknownMetric(KeyError):
    pass


class Resolver(Protocol):
    def members(self, metric: str, cube: str) -> tuple[str, ...]: ...
    def sign(self, metric: str) -> Sign: ...
    def status(self, metric: str) -> Status: ...
    def materiality(self, metric: str) -> tuple[float, float]: ...
    def balance(self, metric: str) -> Balance: ...
    def aggregation(self, metric: str) -> Aggregation: ...
    def exists(self, member: str, cube: str) -> bool: ...
    def children(self, member: str, cube: str) -> tuple[str, ...]: ...
    def weight(self, member: str, cube: str) -> float: ...
    def label(self, member: str, cube: str) -> str: ...
    def latest_forecast(self) -> str: ...
    def concept(self, name: str): ...          # a product concept (knowledge.semantic.Concept) or None
    def metric_of(self, member: str, cube: str) -> str | None: ...   # the metric of an account (or its parent)
    def pct_floor(self, metric: str) -> float | None: ...  # |base| below it: no variance %; None = no floor


def pct_floor_of(unit: str, materiality: tuple[float, float], own: float | None) -> float | None:
    """The variance % floor: the metric's own floor when it has one; else a tenth of the absolute materiality for a
    NOK metric; a stock or ratio metric (FTE, count, ratio) without its own floor has none. Never zero or below."""
    f = own if own is not None else (materiality[0] / 10 if unit == "NOK" else None)
    return f if f is not None and f > 0 else None


def metric_of_member(metrics, member: str, cube: str, parent: str | None) -> str | None:
    """The name of the metric whose members in the cube are exactly this member (any case); else the metric of the
    member's parent the same way; else None. metrics: MetricDef-like (name, members dict)."""
    for m in (member, parent):
        if not m:
            continue
        key = m.casefold()
        for d in metrics:
            ms = d.members.get(cube, ())
            if len(ms) == 1 and ms[0].casefold() == key:
                return d.name
    return None


@dataclass
class MetricDef:
    name: str
    members: dict[str, tuple[str, ...]]
    sign: Sign = "neutral"
    status: Status = "approved"
    materiality: tuple[float, float] = DEFAULT_MATERIALITY
    balance: Balance = "Periodic"
    aggregation: Aggregation = "sum"
    unit: str = "NOK"                       # NOK | FTE | count | ratio
    pct_floor: float | None = None          # own variance % floor (else materiality abs / 10 for NOK)

    def __post_init__(self):
        if self.balance == "Accumulated" and self.aggregation == "sum":
            raise ValueError(f"metric {self.name!r}: an Accumulated balance cannot be summed over months "
                             "(Accumulated with aggregation 'sum' double counts); use aggregation 'last'")


@dataclass
class StaticResolver:
    metrics: list[MetricDef]
    tree: dict[str, tuple[str, ...]] = field(default_factory=dict)
    labels: dict[str, str] = field(default_factory=dict)
    extra_members: set[str] = field(default_factory=set)
    forecast: str = "Forecast Q326"
    weights: dict[str, float] = field(default_factory=dict)
    concepts: dict = field(default_factory=dict)          # id -> knowledge.semantic.Concept

    def __post_init__(self):
        self._m = {m.name.lower(): m for m in self.metrics}
        self._known = (set(self.extra_members) | set(self.tree)
                       | {c for cs in self.tree.values() for c in cs}
                       | {a for m in self.metrics for ms in m.members.values() for a in ms})

    def members(self, metric: str, cube: str) -> tuple[str, ...]:
        m = self._m.get(metric.lower())
        if m is None:
            if metric in self._known:
                return (metric,)
            raise UnknownMetric(metric)
        if cube not in m.members:
            raise UnknownMetric(f"{metric} has no {cube} definition")
        return m.members[cube]

    def sign(self, metric: str) -> Sign:
        m = self._m.get(metric.lower())
        return m.sign if m else "neutral"

    def status(self, metric: str) -> Status:
        m = self._m.get(metric.lower())
        return m.status if m else "approved"

    def materiality(self, metric: str) -> tuple[float, float]:
        m = self._m.get(metric.lower())
        return m.materiality if m else DEFAULT_MATERIALITY

    def pct_floor(self, metric: str) -> float | None:
        m = self._m.get(metric.lower())
        return pct_floor_of(m.unit, m.materiality, m.pct_floor) if m else None

    def balance(self, metric: str) -> Balance:
        m = self._m.get(metric.lower())
        return m.balance if m else "Periodic"

    def aggregation(self, metric: str) -> Aggregation:
        m = self._m.get(metric.lower())
        return m.aggregation if m else "sum"

    def exists(self, member: str, cube: str) -> bool:
        return member in self._known

    def children(self, member: str, cube: str) -> tuple[str, ...]:
        return self.tree.get(member, ())

    def weight(self, member: str, cube: str) -> float:
        """Consolidation weight in its parent: 1.0 add, -1.0 subtract, 0.0 ignore ('~')."""
        return self.weights.get(member, 1.0)

    def label(self, member: str, cube: str) -> str:
        return self.labels.get(member, member)

    def latest_forecast(self) -> str:
        return self.forecast

    def metric_of(self, member: str, cube: str) -> str | None:
        parent = next((p for p, kids in self.tree.items() if any(k.casefold() == member.casefold() for k in kids)),
                      None)
        return metric_of_member(self.metrics, member, cube, parent)

    def concept(self, name: str):
        key = name.strip().casefold()
        for c in self.concepts.values():
            if key in {c.id.casefold(), c.title.casefold(), *(n.casefold() for n in c.names)}:
                return c
        return None
