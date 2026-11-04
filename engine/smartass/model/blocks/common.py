# engine/smartass/model/blocks/common.py
"""What every block builder produces (BlockOut: headers + lines of expression trees) and the shared builders for
metric rows and entity rows (children, top N, Other, parent total, a children = parent check)."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from ..context import Ctx
from ..expr import Cells, Expr, Ref, Sub, Sum, share, var_pct
from ..frame import Frame
from ..spec import Query

BASES = ("ly", "budget", "forecast")
HEAD = {"value": "Actual", "ly": "Last year", "budget": "Budget", "forecast": "Forecast", "var": "Variance",
        "var_pct": "Variance %", "share": "Share %"}


@dataclass
class Line:
    key: str
    label: str
    cells: list[Expr | None]
    kind: Literal["line", "sum", "grand", "check"] = "line"
    level: int = 0
    metric: str | None = None


@dataclass
class BlockOut:
    key: str
    type: str
    title: str
    headers: list[str]
    lines: list[Line] = field(default_factory=list)
    chart: str | None = None
    drills: list[dict] = field(default_factory=list)


def ck(block: str, line: str, col: int) -> str:
    return f"{block}.{line}.{col}"


def evaluate(outs: list[BlockOut], frame: Frame) -> Cells:
    exprs = {ck(o.key, ln.key, j): e for o in outs for ln in o.lines for j, e in enumerate(ln.cells)}
    cells = Cells(frame, exprs)
    for k in exprs:
        cells.get(k)
    return cells


def derived(c: str, cols: list[str], key: str, lkey: str, total_key: str | None = None) -> Expr | None:
    vi = cols.index("value")
    base = next(x for x in cols if x in BASES) if c in ("var", "var_pct") else None
    if c == "var":
        return Sub(Ref(ck(key, lkey, vi)), Ref(ck(key, lkey, cols.index(base))))
    if c == "var_pct":
        return var_pct(Ref(ck(key, lkey, vi)), Ref(ck(key, lkey, cols.index(base))))
    if c == "share":
        return share(Ref(ck(key, lkey, vi)), Ref(ck(key, total_key, vi))) if total_key else None
    raise ValueError(f"not a derived column: {c}")


def row_cells(ctx: Ctx, q: Query, metric: str, cols: list[str], key: str, lkey: str,
              entity: str | None = None, total_key: str | None = None) -> list[Expr | None]:
    return [ctx.pick(q, metric, c, entity) if c in ("value",) + BASES else derived(c, cols, key, lkey, total_key)
            for c in cols]


def entity_lines(ctx: Ctx, q: Query, metric: str, cols: list[str], key: str, *, top: int | None = None,
                 other: bool = True, sort: str = "none") -> list[Line]:
    r = ctx.resolver
    kids = list(r.children(q.entity, q.cube))

    def val(e: str) -> float:
        return ctx.value(ctx.pick(q, metric, "value", e)) or 0.0

    if sort != "none" or top:
        order = {"desc": lambda e: -val(e), "asc": val, "abs": lambda e: -abs(val(e)),
                 "none": lambda e: -abs(val(e))}[sort]
        kids.sort(key=order)
    shown, rest = (kids[:top], kids[top:]) if top else (kids, [])
    lines = [Line(f"e{i}", r.label(e, q.cube),
                  row_cells(ctx, q, metric, cols, key, f"e{i}", entity=e, total_key="total"), metric=metric)
             for i, e in enumerate(shown)]
    if rest and other:
        cells: list[Expr | None] = []
        for j, c in enumerate(cols):
            if c in ("value",) + BASES:
                cells.append(Sub(Ref(ck(key, "total", j)),
                                 Sum(tuple(Ref(ck(key, f"e{i}", j)) for i in range(len(shown))))))
            else:
                cells.append(derived(c, cols, key, "other", total_key="total"))
        lines.append(Line("other", "Other", cells, metric=metric))
    lines.append(Line("total", r.label(q.entity, q.cube),
                      row_cells(ctx, q, metric, cols, key, "total", entity=q.entity, total_key="total"),
                      kind="grand", metric=metric))
    if kids:
        vi = cols.index("value")
        check: list[Expr | None] = [None] * len(cols)
        check[vi] = Sub(ctx.pick(q, metric, "value", tuple(kids)), Ref(ck(key, "total", vi)))
        lines.append(Line("check", "Check: children = parent", check, kind="check"))
    return lines
