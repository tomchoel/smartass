# engine/smartass/model/blocks/table.py
"""table: metric rows, entity rows (children of the query's entity) or product rows (the query's products) by
columns; kpi_tiles: headline metrics with one comparison."""
from __future__ import annotations

from ..context import Ctx
from ..spec import KpiBlock, TableBlock
from .common import BlockOut, Line, col_headers, entity_lines, metric_label, product_lines, row_cells


def build_table(b: TableBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = list(b.columns)
    out = BlockOut(key, "table", b.title or "Table", col_headers(cols), cols=cols, query=q.id)
    if b.rows == "entities":
        out.lines = entity_lines(ctx, q, b.metric, cols, key, top=b.top, other=b.other, sort=b.sort)
    elif b.rows == "products":
        out.lines = product_lines(ctx, q, b.metric, cols, key, top=b.top, other=b.other, sort=b.sort)
    else:
        for i, m in enumerate(b.metrics or q.metrics):
            out.lines.append(Line(f"m{i}", metric_label(ctx.resolver, m, q.cube),
                                  row_cells(ctx, q, m, cols, key, f"m{i}"), metric=m,
                                  entity=q.entity))
    return out


def build_kpi(b: KpiBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = ["value", b.compare, "var", "var_pct"]
    out = BlockOut(key, "kpi_tiles", b.title or "Key figures", col_headers(cols), cols=cols,
                   query=q.id)
    for i, m in enumerate(b.metrics):
        out.lines.append(Line(f"k{i}", metric_label(ctx.resolver, m, q.cube),
                              row_cells(ctx, q, m, cols, key, f"k{i}"), metric=m,
                                  entity=q.entity))
    return out
