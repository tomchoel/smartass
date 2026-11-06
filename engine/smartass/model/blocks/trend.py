# engine/smartass/model/blocks/trend.py
"""trend: one line per month of the query's periods, a column per metric x series (optionally cumulative);
matrix: children of the query's entity x months, with the parent total and a children = parent check."""
from __future__ import annotations

from ..context import Ctx
from ..expr import Expr, Ref, Sub, Sum
from ..periods import label, period_label, resolve
from ..series import months
from ..spec import MatrixBlock, TrendBlock
from ..units import NOK, short_name, unit_class
from .common import HEAD, SCEN, BlockOut, Line, ck, children_sum, code_of, counted, metric_label


def build_trend(b: TrendBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cl = ctx.closed_for(q.cube)
    combos = [(m, s) for m in b.metrics for s in b.series]
    # a series lives on its own months: value and ly on the query's periods, budget and forecast on plan_periods
    own = {s: set(months(q, "value" if s == "ly" else s, cl)) for s in b.series}
    ms = sorted(set().union(*own.values()))
    names = {m: metric_label(ctx.resolver, m, q.cube) for m in b.metrics}
    keys = [f"{names[m]} {HEAD[s]}" for m, s in combos]                  # the stable key_values keys
    heads = [f"{short_name(names[m])} {SCEN[s]}" for m, s in combos]       # the sheet: 'Revenue AC'
    out = BlockOut(key, "trend", b.title or "Trend", heads,
                   cols=[s for _, s in combos], query=q.id, chart="line" if b.chart else None,
                   col_metrics=[m for m, _ in combos], keys=keys,
                   period=period_label(ms) if ms else "")
    stock = [ctx.resolver.aggregation(m) == "last" for m, _ in combos]   # a stock is picked per month, never added
    # only a NOK flow adds up over months: a count, ARPU, a percentage or a stock has no Total (blank)
    flow = [unit_class(ctx.resolver, m, q.cube) == NOK and not stock[j] for j, (m, _) in enumerate(combos)]
    for i, mon in enumerate(ms):
        cells: list[Expr | None] = []
        for j, (m, s) in enumerate(combos):
            if mon not in own[s]:
                cells.append(None)
                continue
            this = ctx.pick(q, m, s, base_months=[mon])
            cells.append(Sum((Ref(ck(key, f"p{i - 1}", j)), this)) if b.cumulative and i and not stock[j] else this)
        out.lines.append(Line(f"p{i}", label(mon), cells))
    if not b.cumulative and len(ms) > 1 and any(flow):
        def total(j: int) -> Expr | None:
            if not flow[j]:
                return None
            return Sum(tuple(Ref(ck(key, f"p{i}", j)) for i in range(len(ms))))
        out.lines.append(Line("total", "Total", [total(j) for j in range(len(combos))], kind="grand"))
    return out


def build_matrix(b: MatrixBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    ms = resolve(q.periods, ctx.closed_for(q.cube))
    r = ctx.resolver
    kids = r.children(q.entity, q.cube)
    out = BlockOut(key, "matrix", b.title or metric_label(r, b.metric, q.cube), [label(m) for m in ms],
                   cols=["value"] * len(ms), query=q.id)
    def empty(e) -> bool:              # nothing in any month (or only rounding noise): no row (Tom 06.11.2026)
        vals = [ctx.value(ctx.pick(q, b.metric, "value", e, [m])) for m in ms]
        return all(v is None or abs(v) < 0.5 for v in vals)

    shown = [e for e in kids if not empty(e)] or list(kids)
    for i, e in enumerate(shown):
        out.lines.append(Line(f"e{i}", r.label(e, q.cube), [ctx.pick(q, b.metric, "value", e, [m]) for m in ms],
                              metric=b.metric, entity=e, code=code_of(e)))
    out.lines.append(Line("total", r.label(q.entity, q.cube),
                          [ctx.pick(q, b.metric, "value", q.entity, [m]) for m in ms], kind="grand", metric=b.metric,
                          entity=q.entity, code=code_of(q.entity)))
    live = counted(ctx, q, kids)
    if live:
        out.lines.append(Line("check", "Check: children = parent",
                              [Sub(children_sum(ctx, q, b.metric, live, base_months=[m]), Ref(ck(key, "total", j)))
                               for j, m in enumerate(ms)], kind="check"))
    return out
