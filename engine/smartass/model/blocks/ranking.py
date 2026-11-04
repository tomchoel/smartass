# engine/smartass/model/blocks/ranking.py
"""ranking: the query entity's children (or, with rows 'products', the query's products) ranked by size (absolute value), top N + Other + total, share %, and the
comparison when asked. Suppliers and projects come from drill lines (Plan 1B), not from the cube."""
from __future__ import annotations

from ..context import Ctx
from ..spec import RankingBlock
from .common import BlockOut, col_headers, entity_lines, metric_label, product_lines


def build_ranking(b: RankingBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = ["value"] + ([b.compare, "var", "var_pct"] if b.compare else []) + ["share"]
    title = b.title or f"Top {b.top} by {metric_label(ctx.resolver, b.metric, q.cube)}"
    out = BlockOut(key, "ranking", title, col_headers(cols), cols=cols,
                   query=q.id)
    rows = product_lines if b.rows == "products" else entity_lines
    out.lines = rows(ctx, q, b.metric, cols, key, top=b.top, other=True, sort="abs")
    return out
