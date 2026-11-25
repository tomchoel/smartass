# engine/smartass/model/resolver.py
"""The model's only view of knowledge. Sub-spec 2 implements Resolver over knowledge.db; StaticResolver serves
tests and seeds."""
from __future__ import annotations

import tomllib
from dataclasses import dataclass, field
from typing import Literal, Protocol

Sign = Literal["income", "cost", "neutral"]
Status = Literal["approved", "draft"]
Balance = Literal["Periodic", "Accumulated"]
Aggregation = Literal["sum", "last"]


_RULES: dict = {}


def _rules() -> list[dict]:
    """Every rule file of the semantic folder, parsed again when a file's mtime or size changes (edits are seen
    without a restart; the stat of a handful of files is cheap)."""
    from ..core import config
    d = config.SEMANTIC_DIR / "rules"
    files = sorted(d.glob("*.toml"))
    key = (str(d), tuple((p.name, p.stat().st_mtime_ns, p.stat().st_size) for p in files))
    if key not in _RULES:
        _RULES.clear()
        _RULES[key] = [tomllib.loads(p.read_text(encoding="utf-8")) for p in files]
    return _RULES[key]


def rule_data(kind: str) -> dict:
    """The data of the one approved rule of this kind in the semantic folder (knowledge/semantic/rules/*.toml). The
    model reads the same file the build checks, so a rule is defined once; a draft rule is not in force."""
    from ..core import config
    for d in _rules():
        if d.get("kind") == kind and d.get("status", "draft") == "approved":
            return d["data"]
    raise LookupError(f"no approved rule of kind {kind!r} in {config.SEMANTIC_DIR / 'rules'}")


def default_materiality() -> tuple[float, float]:
    d = rule_data("materiality")
    return (float(d["abs_nok"]), float(d["pct"]))


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
    def weight(self, member: str, cube: str, parent: str | None = None) -> float: ...
    def label(self, member: str, cube: str) -> str: ...
    def latest_forecast(self) -> str: ...
    def scenarios(self) -> dict: ...          # the `scenarios` rule data: cube -> series -> [scenario, version]
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
    materiality: tuple[float, float] = field(default_factory=default_materiality)
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

    def metric(self, name: str) -> MetricDef | None:
        """The metric for a name (any case); None for a raw member."""
        return self._m.get(name.strip().lower())

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
        return m.materiality if m else default_materiality()

    def pct_floor(self, metric: str) -> float | None:
        m = self._m.get(metric.lower())
        return pct_floor_of(m.unit, m.materiality, m.pct_floor) if m else None

    def exists(self, member: str, cube: str) -> bool:
        return member in self._known

    def children(self, member: str, cube: str) -> tuple[str, ...]:
        return self.tree.get(member, ())

    def parents(self, member: str, cube: str, dimension: str | None = None) -> tuple[str, ...]:
        """All parents of member in tree (supports shared multi-parent members)."""
        want = member.casefold()
        return tuple(p for p, kids in self.tree.items() if any(k.casefold() == want for k in kids))

    def resolve_product(self, name: str, cube: str) -> str:
        return name

    def weight(self, member: str, cube: str, parent: str | None = None) -> float:
        """Consolidation weight in its parent: checks (parent, member) first, then member; default 1.0."""
        if parent is not None and (parent, member) in self.weights:
            return self.weights[(parent, member)]
        return self.weights.get(member, 1.0)

    def _member_metric(self, token: str) -> MetricDef | None:
        key = token.strip().casefold()
        for d in self.metrics:
            for ms in d.members.values():
                if len(ms) == 1 and ms[0].casefold() == key:
                    return d
        return None

    def balance(self, metric: str) -> Balance:
        m = self._m.get(metric.lower()) or self._member_metric(metric)
        return m.balance if m else "Periodic"

    def aggregation(self, metric: str) -> Aggregation:
        m = self._m.get(metric.lower()) or self._member_metric(metric)
        return m.aggregation if m else "sum"

    def label(self, member: str, cube: str) -> str:
        return self.labels.get(member, member)

    def latest_forecast(self) -> str:
        return self.forecast

    def scenarios(self) -> dict:
        return rule_data("scenarios")

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
