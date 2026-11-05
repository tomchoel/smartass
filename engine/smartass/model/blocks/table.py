# engine/smartass/model/blocks/table.py
"""table: metric rows, entity rows (children of the query's entity) or product rows (the query's products) by
columns, optionally in column groups (one per period, e.g. month and YTD, separated by a gap column); kpi_tiles:
headline metrics with one comparison."""
from __future__ import annotations

from ..context import Ctx
from ..periods import resolve
from ..spec import KpiBlock, TableBlock
from .common import GAP, BlockOut, Line, entity_lines, metric_label, product_lines, row_cells, scenario_heads


def build_table(b: TableBlock, ctx: Ctx, key: str) -> BlockOut:
    """Without periods: one group over the query's months. With periods: one group per entry; the last group keeps
    the plain keys (stable month to month), earlier groups get the key prefix '<expr> · '."""
    q = ctx.query(b.query)
    closed = ctx.closed_for(q.cube)
    groups = b.periods or [None]
    cols: list[str] = []
    bm: list = []
    heads: list[str] = []
    keys: list[str] = []
    for g, expr in enumerate(groups):
        if g:
            cols.append(GAP)
            bm.append(None)
            heads.append("")
            keys.append("")
        base = resolve(expr, closed) if expr else None
        h, k = scenario_heads(ctx, q, list(b.columns), base)
        cols += b.columns
        bm += [base] * len(b.columns)
        heads += h
        keys += k if g == len(groups) - 1 else [f"{expr} · {x}" for x in k]
    bm_or_none = bm if b.periods else None
    out = BlockOut(key, "table", b.title or "Table", heads, cols=cols, query=q.id, keys=keys)
    if b.rows == "entities":
        out.lines = entity_lines(ctx, q, b.metric, cols, key, top=b.top, other=b.other, sort=b.sort, bm=bm_or_none)
    elif b.rows == "products":
        out.lines = product_lines(ctx, q, b.metric, cols, key, top=b.top, other=b.other, sort=b.sort,
                                  bm=bm_or_none)
    else:
        for i, m in enumerate(b.metrics or q.metrics):
            out.lines.append(Line(f"m{i}", metric_label(ctx.resolver, m, q.cube),
                                  row_cells(ctx, q, m, cols, key, f"m{i}", bm=bm_or_none), metric=m,
                                  entity=q.entity))
    return out


def build_kpi(b: KpiBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = ["value", b.compare, "var", "var_pct"]
    heads, keys = scenario_heads(ctx, q, cols)
    out = BlockOut(key, "kpi_tiles", b.title or "Key figures", heads, cols=cols, query=q.id, keys=keys)
    for i, m in enumerate(b.metrics):
        out.lines.append(Line(f"k{i}", metric_label(ctx.resolver, m, q.cube),
                              row_cells(ctx, q, m, cols, key, f"k{i}"), metric=m,
                              entity=q.entity))
    return out
