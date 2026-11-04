# engine/smartass/model/context.py
"""Build context shared by the block builders: the spec, knowledge, the closed month and the fetched facts."""
from __future__ import annotations

from dataclasses import dataclass

from .expr import Cells, Expr, pick
from .frame import Frame
from .periods import Month, label
from .resolver import Resolver
from .series import months, scenario
from .spec import AnalysisSpec, Query


@dataclass
class Ctx:
    spec: AnalysisSpec
    resolver: Resolver
    closed: Month
    frame: Frame

    def query(self, qid: str) -> Query:
        return next(q for q in self.spec.queries if q.id == qid)

    def pick(self, q: Query, metric: str, series: str = "value", entity: str | tuple[str, ...] | None = None,
             base_months: list[Month] | None = None) -> Expr:
        scen, ver = scenario(q.cube, series, self.resolver)
        ents = (q.entity,) if entity is None else ((entity,) if isinstance(entity, str) else tuple(entity))
        f = {"Cube": (q.cube,), "Query": (q.id,), "Account": self.resolver.members(metric, q.cube),
             "Entity": ents, "Period": tuple(label(m) for m in months(q, series, self.closed, base_months)),
             "Scenario": (scen,)}
        if ver:
            f["Version"] = (ver,)
        return pick(f)

    def value(self, e: Expr) -> float | None:
        """Value of a self-contained tree (no Refs)."""
        return e.evaluate(self.frame, Cells(self.frame, {}))
