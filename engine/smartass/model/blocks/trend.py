# engine/smartass/model/blocks/trend.py
"""trend: one line per month of the query's periods, a column per metric x series (optionally cumulative);
matrix: children of the query's entity x months, with the parent total and a children = parent check."""
from __future__ import annotations

from ..context import Ctx
from ..expr import Expr, Ref, Sub, Sum
from ..periods import label, resolve
from ..spec import MatrixBlock, TrendBlock
from .common import HEAD, BlockOut, Line, ck, children_sum, counted


def build_trend(b: TrendBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    ms = resolve(q.periods, ctx.closed_for(q.cube))
    combos = [(m, s) for m in b.metrics for s in b.series]
    out = BlockOut(key, "trend", b.title or "Trend", [f"{m} {HEAD[s]}" for m, s in combos],
                   cols=[s for _, s in combos], query=q.id, chart="line" if b.chart else None)
    stock = [ctx.resolver.aggregation(m) == "last" for m, _ in combos]   # a stock is picked per month, never added
    for i, mon in enumerate(ms):
        cells: list[Expr | None] = []
        for j, (m, s) in enumerate(combos):
            this = ctx.pick(q, m, s, base_months=[mon])
            cells.append(Sum((Ref(ck(key, f"p{i - 1}", j)), this)) if b.cumulative and i and not stock[j] else this)
        out.lines.append(Line(f"p{i}", label(mon), cells))
    if not b.cumulative and len(ms) > 1:
        out.lines.append(Line("total", "Total", [Ref(ck(key, f"p{len(ms) - 1}", j)) if stock[j]
                                                 else Sum(tuple(Ref(ck(key, f"p{i}", j)) for i in range(len(ms))))
                                                 for j in range(len(combos))], kind="grand"))
    return out


def build_matrix(b: MatrixBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    ms = resolve(q.periods, ctx.closed_for(q.cube))
    r = ctx.resolver
    kids = r.children(q.entity, q.cube)
    out = BlockOut(key, "matrix", b.title or b.metric, [label(m) for m in ms], cols=["value"] * len(ms), query=q.id)
    for i, e in enumerate(kids):
        out.lines.append(Line(f"e{i}", r.label(e, q.cube), [ctx.pick(q, b.metric, "value", e, [m]) for m in ms],
                              metric=b.metric, entity=e))
    out.lines.append(Line("total", r.label(q.entity, q.cube),
                          [ctx.pick(q, b.metric, "value", q.entity, [m]) for m in ms], kind="grand", metric=b.metric,
                          entity=q.entity))
    live = counted(ctx, q, kids)
    if live:
        out.lines.append(Line("check", "Check: children = parent",
                              [Sub(children_sum(ctx, q, b.metric, live, base_months=[m]), Ref(ck(key, "total", j)))
                               for j, m in enumerate(ms)], kind="check"))
    return out
