# engine/smartass/model/blocks/common.py
"""What every block builder produces (BlockOut: headers + lines of expression trees) and the shared builders for
metric rows and entity rows (children, top N, Other, parent total, a children = parent check)."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from ..context import Ctx
from ..expr import Cells, Expr, Neg, Ref, Sub, Sum, share, var_pct
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
    entity: str | None = None


@dataclass
class BlockOut:
    key: str
    type: str
    title: str
    headers: list[str]
    cols: list[str] = field(default_factory=list)   # column codes parallel to headers: value, ly, budget,
    # forecast, var, var_pct, share, amount (bridge)
    query: str | None = None
    lines: list[Line] = field(default_factory=list)
    chart: str | None = None
    drills: list = field(default_factory=list)   # DrillRequest (blocks/drivers.py)


def ck(block: str, line: str, col: int) -> str:
    return f"{block}.{line}.{col}"


def evaluate(outs: list[BlockOut], frame: Frame) -> Cells:
    exprs = {ck(o.key, ln.key, j): e for o in outs for ln in o.lines for j, e in enumerate(ln.cells)}
    cells = Cells(frame, exprs)
    for k in exprs:
        cells.get(k)
    return cells


def base_of(cols: list[str], j: int) -> int:
    """Index of the base column (ly/budget/forecast) nearest to the left of column j."""
    return next(i for i in range(j - 1, -1, -1) if cols[i] in BASES)


def col_headers(cols: list[str]) -> list[str]:
    """Column headers; a variance names its base when the table has several base columns."""
    several = sum(c in BASES for c in cols) > 1
    return [f"{HEAD[c]} vs {HEAD[cols[base_of(cols, j)]]}" if several and c in ("var", "var_pct") else HEAD[c]
            for j, c in enumerate(cols)]


def derived(c: str, cols: list[str], key: str, lkey: str, total_key: str | None = None,
            j: int | None = None) -> Expr | None:
    """The derived cell of column j (default: the first column named c)."""
    j = cols.index(c) if j is None else j
    vi = cols.index("value")
    if c in ("var", "var_pct"):
        bi = base_of(cols, j)
        cur, base = Ref(ck(key, lkey, vi)), Ref(ck(key, lkey, bi))
        return Sub(cur, base) if c == "var" else var_pct(cur, base)
    if c == "share":
        return share(Ref(ck(key, lkey, vi)), Ref(ck(key, total_key, vi))) if total_key else None
    raise ValueError(f"not a derived column: {c}")


def row_cells(ctx: Ctx, q: Query, metric: str, cols: list[str], key: str, lkey: str,
              entity: str | None = None, total_key: str | None = None) -> list[Expr | None]:
    return [ctx.pick(q, metric, c, entity) if c in ("value",) + BASES
            else derived(c, cols, key, lkey, total_key, j) for j, c in enumerate(cols)]


def counted(ctx: Ctx, q: Query, kids) -> list[str]:
    """Children that take part in the consolidation (weight 0, '~', is ignored)."""
    return [k for k in kids if ctx.resolver.weight(k, q.cube) != 0]


def children_sum(ctx: Ctx, q: Query, metric: str, kids, series: str = "value",
                 base_months: list | None = None) -> Expr:
    """What the children add up to in the parent: weight +1 children minus weight -1 children; weight 0 skipped."""
    w = {k: ctx.resolver.weight(k, q.cube) for k in kids}
    plus = tuple(k for k in kids if w[k] > 0)
    minus = tuple(k for k in kids if w[k] < 0)
    if not plus and not minus:
        raise ValueError(f"{q.entity}: no child takes part in the consolidation")
    neg = ctx.pick(q, metric, series, minus, base_months) if minus else None
    if not plus:
        return Neg(neg)
    pos = ctx.pick(q, metric, series, plus, base_months)
    return Sub(pos, neg) if neg is not None else pos


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
                  row_cells(ctx, q, metric, cols, key, f"e{i}", entity=e, total_key="total"), metric=metric,
                  entity=e)
             for i, e in enumerate(shown)]
    if rest and other:
        cells: list[Expr | None] = []
        for j, c in enumerate(cols):
            if c in ("value",) + BASES:
                cells.append(Sub(Ref(ck(key, "total", j)),
                                 Sum(tuple(Ref(ck(key, f"e{i}", j)) for i in range(len(shown))))))
            else:
                cells.append(derived(c, cols, key, "other", total_key="total", j=j))
        lines.append(Line("other", "Other", cells, metric=metric))
    lines.append(Line("total", r.label(q.entity, q.cube),
                      row_cells(ctx, q, metric, cols, key, "total", entity=q.entity, total_key="total"),
                      kind="grand", metric=metric, entity=q.entity))
    if counted(ctx, q, kids):
        vi = cols.index("value")
        check: list[Expr | None] = [None] * len(cols)
        check[vi] = Sub(children_sum(ctx, q, metric, counted(ctx, q, kids)), Ref(ck(key, "total", vi)))
        lines.append(Line("check", "Check: children = parent", check, kind="check"))
    return lines
