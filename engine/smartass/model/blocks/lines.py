"""lines: the ledger lines behind one cube cell, grouped by supplier (AP_CUSTOMER_NAME), project (INITIATIVE) or
account (ACCOUNT): top N (biggest absolute amount first; spellings that differ only in case are one group), Other,
lines without a <group>, the lines total, the cube value and the gap (cube - lines), share of the lines total."""
from __future__ import annotations

from ..expr import CubeRef, LinePick, Ref, Sub, Sum, share
from ..lines import AMOUNT, norm
from ..spec import LinesBlock
from .common import BlockOut, Line, ck

COLUMN = {"supplier": "AP_CUSTOMER_NAME", "project": "INITIATIVE", "account": "ACCOUNT"}
NOUN = {"supplier": "supplier", "project": "booked project", "account": "account"}


def build_lines(b: LinesBlock, ctx, key: str) -> BlockOut:
    ls = ctx.frame.lines[key]
    col = COLUMN[b.group_by]
    groups: dict[str, list] = {}
    for r in ls.rows:
        name = str(r.get(col) or "").strip()
        if name:
            g = groups.setdefault(norm(name), [name, 0.0, set()])
            g[1] += float(r.get(AMOUNT) or 0.0)
            g[2].add(name)
    ranked = sorted(groups.values(), key=lambda g: -abs(g[1]))
    shown, rest = ranked[:b.top], ranked[b.top:]
    out = BlockOut(key, "lines", b.title or f"Ledger lines by {b.group_by}", ["Amount", "Share %"])
    out.cols = ["amount", "share"]
    total = Ref(ck(key, "total", 0))
    for i, (name, _, spellings) in enumerate(shown):
        out.lines.append(Line(f"g{i}", name, [LinePick(key, col, tuple(sorted(spellings))),
                                              share(Ref(ck(key, f"g{i}", 0)), total)]))
    if rest:
        named = Sum(tuple(Ref(ck(key, f"g{i}", 0)) for i in range(len(shown))))
        other = Sub(Sub(total, Ref(ck(key, "blank", 0))), named)
        out.lines.append(Line("other", "Other", [other, share(Ref(ck(key, "other", 0)), total)]))
    out.lines.append(Line("blank", f"Lines without a {NOUN[b.group_by]}",
                          [LinePick(key, col, ("",)), share(Ref(ck(key, "blank", 0)), total)]))
    out.lines.append(Line("total", "Ledger lines total", [LinePick(key, None, None), None], kind="sum"))
    out.lines.append(Line("cube", "Cube value", [CubeRef(key), None]))
    out.lines.append(Line("gap", "Gap (cube - lines)", [Sub(Ref(ck(key, "cube", 0)), total), None]))
    return out
