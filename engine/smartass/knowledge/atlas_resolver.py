"""Resolver over v1's atlas.db (every member of the three cubes) and the interim metrics. Limits until sub-spec 2:
no consolidation operators in the atlas, so every child weighs 1.0 (a parent that is not the plain sum of its
children stops at the reconciliation guard); no draft metrics."""
from __future__ import annotations

import re
import sqlite3

from ..model.resolver import DEFAULT_MATERIALITY, MetricDef, UnknownMetric


class AtlasResolver:
    def __init__(self, db: sqlite3.Connection, metrics: list[MetricDef], forecast_year: int):
        self.db = db
        self._m = {m.name.lower(): m for m in metrics}
        self.forecast_year = forecast_year

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

    def _row(self, member: str, cube: str):
        return self.db.execute("SELECT name, dimension, label FROM member WHERE connection=? AND name=? "
                               "COLLATE NOCASE ORDER BY depth LIMIT 1", (cube, member)).fetchone()

    def canonical(self, member: str, cube: str) -> str:
        return self._spelling(member, cube)

    def _spelling(self, member: str, cube: str) -> str:
        row = self._row(member, cube)
        return row[0] if row else member

    def exists(self, member: str, cube: str) -> bool:
        return self._row(member, cube) is not None

    def children(self, member: str, cube: str) -> tuple[str, ...]:
        row = self._row(member, cube)
        if row is None:
            return ()
        names = [r[0] for r in self.db.execute(
            "SELECT name FROM member WHERE connection=? AND dimension=? AND parent=? ORDER BY ord",
            (cube, row[1], row[0]))]
        return tuple(dict.fromkeys(names))

    def weight(self, member: str, cube: str) -> float:
        return 1.0

    def label(self, member: str, cube: str) -> str:
        row = self._row(member, cube)
        if not row:
            return member
        return re.sub(r"\s*\([^()]*\)\s*$", "", row[2] or row[0]) or row[0]

    def latest_forecast(self) -> str:
        names = [r[0] for r in self.db.execute("SELECT DISTINCT name FROM member WHERE connection='FinPlan' AND "
                                               "dimension='Scenario' AND name LIKE 'Forecast Q%'")]
        cands = sorted(((int(m.group(1)), n) for n in names
                        if (m := re.fullmatch(r"Forecast Q(\d)(\d\d)", n)) and int(m.group(2)) == self.forecast_year),
                       reverse=True)
        if not cands:
            raise UnknownMetric(f"no Forecast Q<q>{self.forecast_year:02d} in the atlas")
        return cands[0][1]
