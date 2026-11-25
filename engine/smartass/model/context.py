# engine/smartass/model/context.py
"""Build context shared by the block builders: the spec, knowledge, the closed month and the fetched facts."""
from __future__ import annotations

from dataclasses import dataclass

from .expr import Cells, Expr, pick
from .frame import Frame
from .periods import Closed, Month, closed_for, label
from .resolver import Resolver
from .series import months, scenario
from .spec import AnalysisSpec, Query


@dataclass
class Ctx:
    spec: AnalysisSpec
    resolver: Resolver
    closed: Closed
    frame: Frame

    def closed_for(self, cube: str) -> Month:
        return closed_for(self.closed, cube)

    def query(self, qid: str) -> Query:
        for q in self.spec.queries:
            if q.id == qid:
                return q
        raise KeyError(f"unknown query {qid!r}")

    def dedup_products(self, q: Query) -> tuple[str, ...]:
        """Deduplicated product set (parent + child overlap counted once) when summing q.products."""
        if len(q.products) <= 1:
            return tuple(q.products)
        from .products import analyze_products
        return analyze_products(q.products, self.resolver, q.cube).dedup_products

    def pick(self, q: Query, metric: str, series: str = "value", entity: str | tuple[str, ...] | None = None,
             base_months: list[Month] | None = None, product: str | tuple[str, ...] | None = None,
             extra_dims: dict[str, tuple[str, ...]] | None = None) -> Expr:
        """The SUMIFS behind one cell. Product filters on the deduplicated query products unless specific
        product(s) are given; a query without products has no Product criterion."""
        scen, ver = scenario(q.cube, series, self.resolver)
        ents = (q.entity,) if entity is None else ((entity,) if isinstance(entity, str) else tuple(entity))
        ms = months(q, series, self.closed_for(q.cube), base_months)
        if self.resolver.aggregation(metric) == "last":
            ms = ms[-1:]
        f = {"Cube": (q.cube,), "Query": (q.id,), "Account": self.resolver.members(metric, q.cube),
             "Entity": ents}
        prods = self.dedup_products(q) if product is None else ((product,) if isinstance(product, str) else tuple(product))
        if prods:
            f["Product"] = tuple(prods)
        f |= {"Period": tuple(label(m) for m in ms), "Scenario": (scen,)}
        if ver:
            f["Version"] = (ver,)
        f["Balance"] = (self.resolver.balance(metric),)
        for dim, val in q.dimensions.items():
            f[dim] = (val,) if isinstance(val, str) else tuple(val)
        for dim, vals in (extra_dims or {}).items():
            f[dim] = vals
        return pick(f)

    def value(self, e: Expr) -> float | None:
        """Value of a self-contained tree (no Refs)."""
        return e.evaluate(self.frame, Cells(self.frame, {}))
