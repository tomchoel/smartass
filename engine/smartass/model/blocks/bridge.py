# engine/smartass/model/blocks/bridge.py
"""bridge: waterfall from the base total (ly / budget / forecast) to the actual total, one step per metric's
variance, plus 'Other / unexplained' so the bridge always closes."""
from __future__ import annotations

from ..context import Ctx
from ..expr import Ref, Sub, Sum
from ..spec import BridgeBlock
from .common import HEAD, BlockOut, Line, ck


def build_bridge(b: BridgeBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    out = BlockOut(key, "bridge", b.title or f"{b.total} bridge", ["Amount"], chart="waterfall")
    out.lines.append(Line("start", f"{b.total} {HEAD[b.base]}", [ctx.pick(q, b.total, b.base)], kind="sum",
                          metric=b.total))
    for i, s in enumerate(b.steps):
        out.lines.append(Line(f"s{i}", s, [Sub(ctx.pick(q, s, "value"), ctx.pick(q, s, b.base))], metric=s))
    steps = Sum(tuple(Ref(ck(key, f"s{i}", 0)) for i in range(len(b.steps))))
    out.lines.append(Line("other", "Other / unexplained",
                          [Sub(Sub(Ref(ck(key, "end", 0)), Ref(ck(key, "start", 0))), steps)]))
    out.lines.append(Line("end", f"{b.total} {HEAD['value']}", [ctx.pick(q, b.total, "value")], kind="grand",
                          metric=b.total))
    return out
