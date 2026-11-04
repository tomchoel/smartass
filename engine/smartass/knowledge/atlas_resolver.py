"""Resolver over v1's atlas.db (every member of the three cubes) and the interim metrics. Replaced by
KnowledgeResolver (knowledge.db); kept until the pipeline is wired to it. No draft metrics."""
from __future__ import annotations

import sqlite3

from ..model.resolver import DEFAULT_MATERIALITY, MetricDef, UnknownMetric
from .atlas_tree import AtlasTree


class AtlasResolver(AtlasTree):
    def __init__(self, db: sqlite3.Connection, metrics: list[MetricDef], forecast_year: int):
        super().__init__(db, forecast_year)
        self._m = {m.name.lower(): m for m in metrics}

    def _metric(self, name: str) -> MetricDef | None:
        return self._m.get(name.lower())

    def metrics(self) -> list[MetricDef]:
        return list(self._m.values())

    def members(self, metric: str, cube: str) -> tuple[str, ...]:
        m = self._metric(metric)
        if m is None:
            if self.exists(metric, cube):
                return (self._spelling(metric, cube),)
            raise UnknownMetric(metric)
        if cube not in m.members:
            raise UnknownMetric(f"{metric} has no {cube} definition")
        return m.members[cube]

    def sign(self, metric):
        m = self._metric(metric)
        return m.sign if m else "neutral"

    def status(self, metric):
        m = self._metric(metric)
        return m.status if m else "approved"

    def materiality(self, metric):
        m = self._metric(metric)
        return m.materiality if m else DEFAULT_MATERIALITY

    def balance(self, metric):
        m = self._metric(metric)
        return m.balance if m else "Periodic"

    def aggregation(self, metric):
        m = self._metric(metric)
        return m.aggregation if m else "sum"
