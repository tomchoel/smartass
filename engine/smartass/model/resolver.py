# engine/smartass/model/resolver.py
"""The model's only view of knowledge. Sub-spec 2 implements Resolver over knowledge.db; StaticResolver serves
tests and seeds."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Protocol

Sign = Literal["income", "cost", "neutral"]
Status = Literal["approved", "draft"]
DEFAULT_MATERIALITY = (5_000_000.0, 0.05)


class UnknownMetric(KeyError):
    pass


class Resolver(Protocol):
    def members(self, metric: str, cube: str) -> tuple[str, ...]: ...
    def sign(self, metric: str) -> Sign: ...
    def status(self, metric: str) -> Status: ...
    def materiality(self, metric: str) -> tuple[float, float]: ...
    def exists(self, member: str, cube: str) -> bool: ...
    def children(self, member: str, cube: str) -> tuple[str, ...]: ...
    def label(self, member: str, cube: str) -> str: ...
    def latest_forecast(self) -> str: ...


@dataclass
class MetricDef:
    name: str
    members: dict[str, tuple[str, ...]]
    sign: Sign = "neutral"
    status: Status = "approved"
    materiality: tuple[float, float] = DEFAULT_MATERIALITY


@dataclass
class StaticResolver:
    metrics: list[MetricDef]
    tree: dict[str, tuple[str, ...]] = field(default_factory=dict)
    labels: dict[str, str] = field(default_factory=dict)
    extra_members: set[str] = field(default_factory=set)
    forecast: str = "Forecast Q326"

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

    def exists(self, member: str, cube: str) -> bool:
        return member in self._known

    def children(self, member: str, cube: str) -> tuple[str, ...]:
        return self.tree.get(member, ())

    def label(self, member: str, cube: str) -> str:
        return self.labels.get(member, member)

    def latest_forecast(self) -> str:
        return self.forecast
