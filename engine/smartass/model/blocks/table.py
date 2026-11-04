# engine/smartass/model/blocks/table.py
"""table: metric rows or entity rows (children of the query's entity) by columns; kpi_tiles: headline metrics
with one comparison."""
from __future__ import annotations

from ..context import Ctx
from ..spec import KpiBlock, TableBlock
from .common import HEAD, BlockOut, Line, entity_lines, row_cells


def build_table(b: TableBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = list(b.columns)
    out = BlockOut(key, "table", b.title or "Table", [HEAD[c] for c in cols])
    if b.rows == "entities":
        out.lines = entity_lines(ctx, q, b.metric, cols, key, top=b.top, other=b.other, sort=b.sort)
    else:
        for i, m in enumerate(b.metrics or q.metrics):
            out.lines.append(Line(f"m{i}", m, row_cells(ctx, q, m, cols, key, f"m{i}"), metric=m))
    return out


def build_kpi(b: KpiBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = ["value", b.compare, "var", "var_pct"]
    out = BlockOut(key, "kpi_tiles", b.title or "Key figures", [HEAD[c] for c in cols])
    for i, m in enumerate(b.metrics):
        out.lines.append(Line(f"k{i}", m, row_cells(ctx, q, m, cols, key, f"k{i}"), metric=m))
    return out
