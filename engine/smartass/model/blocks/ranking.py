# engine/smartass/model/blocks/ranking.py
"""ranking: the query entity's children ranked by size (absolute value), top N + Other + total, share %, and the
comparison when asked. Suppliers and projects come from drill lines (Plan 1B), not from the cube."""
from __future__ import annotations

from ..context import Ctx
from ..spec import RankingBlock
from .common import HEAD, BlockOut, entity_lines


def build_ranking(b: RankingBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = ["value"] + ([b.compare, "var", "var_pct"] if b.compare else []) + ["share"]
    out = BlockOut(key, "ranking", b.title or f"Top {b.top} by {b.metric}", [HEAD[c] for c in cols])
    out.lines = entity_lines(ctx, q, b.metric, cols, key, top=b.top, other=True, sort="abs")
    return out
