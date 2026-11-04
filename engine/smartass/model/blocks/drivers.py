# engine/smartass/model/blocks/drivers.py
"""drivers: decompose a metric's variance down the entity tree. At each level the children are ranked by absolute
variance and shown until they explain 80 % of the gross movement (at most 3); each shown child is expanded the same
way until `depth`. With drill, the leaves reached are returned as drill requests (Plan 1B turns them into ledger
lines), so commentary can rest on data."""
from __future__ import annotations

from dataclasses import dataclass

from ..context import Ctx
from ..expr import Sub
from ..periods import label
from ..series import months
from ..spec import DriversBlock
from .common import BlockOut, Line, col_headers, row_cells

@dataclass(frozen=True)
class DrillRequest:
    """A leaf the data layer drills into ledger lines: the cell's accounts and months, and the base to compare."""
    query: str
    cube: str
    entity: str
    metric: str
    accounts: tuple[str, ...]
    periods: tuple[str, ...]
    base_periods: tuple[str, ...]
    base: str


COVER = 0.8
MAX_PER_LEVEL = 3


def build_drivers(b: DriversBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    r = ctx.resolver
    cols = ["value", b.base, "var", "var_pct"]
    out = BlockOut(key, "drivers", b.title or f"Drivers of {b.metric}", col_headers(cols), cols=cols,
                   query=q.id)

    def variance(entity: str) -> float:
        return ctx.value(Sub(ctx.pick(q, b.metric, "value", entity), ctx.pick(q, b.metric, b.base, entity))) or 0.0

    def add(entity: str, level: int, lkey: str) -> None:
        out.lines.append(Line(lkey, r.label(entity, q.cube), row_cells(ctx, q, b.metric, cols, key, lkey, entity),
                              level=level, metric=b.metric, entity=entity))

    def leaf(entity: str) -> None:
        if b.drill:
            closed = ctx.closed_for(q.cube)
            out.drills.append(DrillRequest(
                q.id, q.cube, entity, b.metric, tuple(r.members(b.metric, q.cube)),
                tuple(label(m) for m in months(q, "value", closed)),
                tuple(label(m) for m in months(q, b.base, closed)), b.base))

    def walk(entity: str, level: int, prefix: str) -> None:
        kids = r.children(entity, q.cube)
        if level >= b.depth or not kids:
            leaf(entity)
            return
        scored = sorted(((variance(k), k) for k in kids), key=lambda t: -abs(t[0]))
        gross = sum(abs(v) for v, _ in scored) or 1.0
        covered = 0.0
        for i, (v, k) in enumerate(scored):
            lkey = f"{prefix}{i}"
            add(k, level + 1, lkey)
            walk(k, level + 1, lkey + "_")
            covered += abs(v)
            if covered / gross >= COVER or i + 1 >= MAX_PER_LEVEL:
                break

    add(q.entity, 0, "top")
    walk(q.entity, 0, "d")
    return out
