# engine/smartass/model/blocks/common.py
"""What every block builder produces (BlockOut: headers + lines of expression trees) and the shared builders for
metric rows and entity rows (children, top N, Other, parent total, a children = parent check)."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from ..context import Ctx
from ..expr import Cells, Expr, Material, Neg, Param, Ref, Sub, Sum, share, var_pct
from ..frame import Frame
from ..periods import Month, period_label
from ..resolver import metric_of_member
from ..series import months
from ..spec import Query

BASES = ("ly", "budget", "forecast")
GAP = "gap"             # the spacer column between column groups (header "", cells None)
HEAD = {"value": "Actual", "ly": "Last year", "budget": "Budget", "forecast": "Forecast", "var": "Variance",
        "var_pct": "Variance %", "share": "Share %"}      # the stable key_values keys (display headers differ)
SCEN = {"value": "AC", "ly": "PY", "budget": "BU", "forecast": "FC"}


@dataclass
class Line:
    key: str
    label: str
    cells: list[Expr | None]
    kind: Literal["line", "sum", "grand", "check", "heading"] = "line"   # heading: a concept sub-heading
    level: int = 0
    metric: str | None = None
    entity: str | None = None
    code: str = ""           # the member code shown beside the label (C30, EBITDA_BOI, 2BUS); '' for none
    added: bool = False       # a chain line complete_chain added (key_values fills the cap with these last)
    display: str = ""         # the label on the sheet when it differs from the model label (key_values keep label)


CODE_PREFIXES = ("ac.", "un.SY.", "pd.SY.")


def code_of(member: str | None) -> str:
    """A member without its dimension prefix: 'ac.C30' -> 'C30', 'un.SY.2BUS' -> '2BUS'; others unchanged."""
    if not member:
        return ""
    for p in CODE_PREFIXES:
        if member.startswith(p):
            return member[len(p):]
    if " " in member:       # a FinPlan alias: '2WHS Wholesale' -> '2WHS'; a name ('EBITDA before Other Items') -> ''
        head = member.split(" ", 1)[0]
        return head if any(ch.isdigit() for ch in head) else ""
    return member


def metric_code(r, metric: str, cube: str) -> str:
    """The code of a metric's first member in the cube ('' when it has none there). FinPlan metrics have none: their
    members are names (Revenue, COGS, Gross_Profit, OPEX, EBITDA) that only repeat the label."""
    if cube == "FinPlan":
        return ""
    try:
        ms = r.members(metric, cube)
    except KeyError:
        return ""
    return code_of(ms[0]) if ms else ""


@dataclass
class BlockOut:
    key: str
    type: str
    title: str
    headers: list[str]                              # display headers ('AC YTD Oct 26', 'ΔPY')
    cols: list[str] = field(default_factory=list)   # column codes parallel to headers: value, ly, budget,
    # forecast, var, var_pct, share, amount (bridge), gap (between column groups)
    query: str | None = None
    col_metrics: list[str] = field(default_factory=list)   # trend: the metric behind each column (for chart units)
    lines: list[Line] = field(default_factory=list)
    chart: str | None = None
    drills: list = field(default_factory=list)   # DrillRequest (blocks/drivers.py)
    keys: list[str] = field(default_factory=list)  # stable key_values keys parallel to headers (default: headers)
    chart_of: str | None = None    # a 'chart' out: the key of the block it draws (b<idx>); it has no lines itself
    steps: list[str] = field(default_factory=list)   # concept: the line keys its waterfall draws (items, total)
    period: str = ""               # the period label of the block's query ('Jan 25 – Oct 26'): the subtitle

    def __post_init__(self):
        if not self.keys:
            self.keys = list(self.headers)


MATERIAL = "material"   # the marker column: 1 when the line is material, else 0
MARK_CODES = {MATERIAL}  # marker columns (0/1 numbers; the layout gives them the marker format)
MAT_ABS, MAT_PCT, MAT_SMALL = "mat_abs.", "mat_pct.", "mat_small."


def add_material(o: BlockOut, r) -> None:
    """Append the Material column to a block with a var column: one Material cell per line over every var of the
    last group, each with the first var_pct against its own base (as findings.material judges it: 1 when the line
    is material against any base); check lines and lines without a metric get None. Without a var in the last
    group nothing is added."""
    last = group_start(o.cols, len(o.cols))
    pairs = []
    for vi in (j for j in range(last, len(o.cols)) if o.cols[j] == "var"):
        bi = o.cols[base_of(o.cols, vi)]
        pairs.append((vi, next((j for j in range(last, len(o.cols))
                                if o.cols[j] == "var_pct" and o.cols[base_of(o.cols, j)] == bi), None)))
    if not pairs:
        return
    for ln in o.lines:
        cell: Expr | None = None
        refs = [(Ref(ck(o.key, ln.key, vi)),
                 Ref(ck(o.key, ln.key, pi)) if pi is not None and ln.cells[pi] is not None else None)
                for vi, pi in pairs if ln.cells[vi] is not None]
        if ln.metric and ln.kind != "check" and refs:
            a, pct = r.materiality(ln.metric)
            (v, p), *more = refs
            cell = Material(v, p, Param(MAT_ABS + ln.metric, a), Param(MAT_PCT + ln.metric, pct),
                            Param(MAT_SMALL + ln.metric, a / 10), tuple(more))
        ln.cells.append(cell)
    o.cols.append(MATERIAL)
    o.headers.append("Material")
    o.keys.append("")


def metric_label(r, token: str, cube: str) -> str:
    """The display name of a metric token: the semantic metric name for a metric id, alias or name; for a raw member
    the name of the metric whose only member in this cube it is, else the atlas label ('C48 GROSS MARGIN'); else
    the token as given."""
    find = getattr(r, "metric", None)
    m = find(token) if find is not None else None
    if m is not None:
        return m.name
    defs = getattr(r, "metrics", None)
    defs = defs() if callable(defs) else (defs or [])
    return metric_of_member(defs, token, cube, None) or r.label(token, cube)


def ck(block: str, line: str, col: int) -> str:
    return f"{block}.{line}.{col}"


def evaluate(outs: list[BlockOut], frame: Frame) -> Cells:
    exprs = {ck(o.key, ln.key, j): e for o in outs for ln in o.lines for j, e in enumerate(ln.cells)}
    cells = Cells(frame, exprs)
    for k in exprs:
        cells.get(k)
    return cells


def group_start(cols: list[str], j: int) -> int:
    """Index of the first column of the column group that holds column j (groups are separated by gap)."""
    return next((i + 1 for i in range(j - 1, -1, -1) if cols[i] == GAP), 0)


def base_of(cols: list[str], j: int) -> int:
    """Index of the base column (ly/budget/forecast) nearest to the left of column j, within its group."""
    return next(i for i in range(j - 1, group_start(cols, j) - 1, -1) if cols[i] in BASES)


def value_of(cols: list[str], j: int) -> int:
    """Index of the value column of column j's group."""
    return cols.index("value", group_start(cols, j))


def col_keys(cols: list[str]) -> list[str]:
    """The stable key_values keys of one column group; a variance names its base when the group has several base
    columns. Codes without a key (gap, helpers) get ''."""
    several = sum(c in BASES for c in cols) > 1
    return [f"{HEAD[c]} vs {HEAD[cols[base_of(cols, j)]]}" if several and c in ("var", "var_pct")
            else HEAD.get(c, "") for j, c in enumerate(cols)]


def col_headers(cols: list[str], *, period: str, ly_period: str, fc: str) -> list[str]:
    """Display headers of one column group: AC/PY/BU/FC with the period, ΔPY/ΔBU/ΔFC (and %) after the
    base to the left, Share %, Amount. Codes without a header (gap, helpers) get ''."""
    def one(j: int, c: str) -> str:
        if c == "value":
            return f"AC {period}"
        if c == "ly":
            return f"PY {ly_period}"
        if c == "budget":
            return f"BU {period}"
        if c == "forecast":
            return f"FC {fc}".strip()
        if c in ("var", "var_pct"):
            return f"Δ{SCEN[cols[base_of(cols, j)]]}" + (" %" if c == "var_pct" else "")
        return {"share": "Share %", "amount": "Amount"}.get(c, "")
    return [one(j, c) for j, c in enumerate(cols)]


def forecast_name(ctx: Ctx, q: Query, cols: list[str]) -> str:
    """The short forecast name for a FinPlan query with a forecast column ('Forecast Q326' -> 'Q326'), else ''."""
    if q.cube != "FinPlan" or "forecast" not in cols:
        return ""
    name = ctx.resolver.latest_forecast()
    return name[len("Forecast "):] if name.startswith("Forecast ") else name


def scenario_heads(ctx: Ctx, q: Query, cols: list[str],
                   base_months: list[Month] | None = None) -> tuple[list[str], list[str]]:
    """(display headers, stable keys) of one column group over the query's months (or base_months)."""
    closed = ctx.closed_for(q.cube)
    heads = col_headers(cols, period=period_label(months(q, "value", closed, base_months)),
                        ly_period=period_label(months(q, "ly", closed, base_months)),
                        fc=forecast_name(ctx, q, cols))
    return heads, col_keys(cols)


def derived(c: str, cols: list[str], key: str, lkey: str, total_key: str | None = None,
            j: int | None = None) -> Expr | None:
    """The derived cell of column j (default: the first column named c); its base and value are in j's group."""
    j = cols.index(c) if j is None else j
    vi = value_of(cols, j)
    if c in ("var", "var_pct"):
        bi = base_of(cols, j)
        cur, base = Ref(ck(key, lkey, vi)), Ref(ck(key, lkey, bi))
        return Sub(cur, base) if c == "var" else var_pct(cur, base)
    if c == "share":
        return share(Ref(ck(key, lkey, vi)), Ref(ck(key, total_key, vi))) if total_key else None
    raise ValueError(f"not a derived column: {c}")


def row_cells(ctx: Ctx, q: Query, metric: str, cols: list[str], key: str, lkey: str,
              entity: str | None = None, total_key: str | None = None,
              product: str | None = None, bm: list | None = None) -> list[Expr | None]:
    """One line's cells; bm (parallel to cols) gives each column's months (column groups), None = the query's."""
    return [ctx.pick(q, metric, c, entity, bm[j] if bm else None, product=product) if c in ("value",) + BASES
            else None if c == GAP else derived(c, cols, key, lkey, total_key, j) for j, c in enumerate(cols)]


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
                 other: bool = True, sort: str = "none", bm: list | None = None) -> list[Line]:
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
                  row_cells(ctx, q, metric, cols, key, f"e{i}", entity=e, total_key="total", bm=bm),
                  metric=metric, entity=e, code=code_of(e))
             for i, e in enumerate(shown)]
    if rest and other:
        cells: list[Expr | None] = []
        for j, c in enumerate(cols):
            if c in ("value",) + BASES:
                cells.append(Sub(Ref(ck(key, "total", j)),
                                 Sum(tuple(Ref(ck(key, f"e{i}", j)) for i in range(len(shown))))))
            elif c == GAP:
                cells.append(None)
            else:
                cells.append(derived(c, cols, key, "other", total_key="total", j=j))
        lines.append(Line("other", "Other", cells, metric=metric))
    lines.append(Line("total", r.label(q.entity, q.cube),
                      row_cells(ctx, q, metric, cols, key, "total", entity=q.entity, total_key="total", bm=bm),
                      kind="grand", metric=metric, entity=q.entity, code=code_of(q.entity)))
    if counted(ctx, q, kids):
        vi = cols.index("value")
        check: list[Expr | None] = [None] * len(cols)
        check[vi] = Sub(children_sum(ctx, q, metric, counted(ctx, q, kids), base_months=bm[vi] if bm else None),
                        Ref(ck(key, "total", vi)))
        lines.append(Line("check", "Check: children = parent", check, kind="check"))
    return lines


PRODUCTS_TOTAL = "Total products"


def product_lines(ctx: Ctx, q: Query, metric: str, cols: list[str], key: str, *, top: int | None = None,
                  other: bool = True, sort: str = "none", bm: list | None = None) -> list[Line]:
    """One line per product of the query (top N, sorted like entity_lines), Other, and the total over the query's
    products (no parent member: the total is the sum of the products asked for, so no check line)."""
    r = ctx.resolver
    prods = list(q.products)

    def val(p: str) -> float:
        return ctx.value(ctx.pick(q, metric, "value", product=p)) or 0.0

    if sort != "none" or top:
        order = {"desc": lambda p: -val(p), "asc": val, "abs": lambda p: -abs(val(p)),
                 "none": lambda p: -abs(val(p))}[sort]
        prods.sort(key=order)
    shown, rest = (prods[:top], prods[top:]) if top else (prods, [])
    lines = [Line(f"p{i}", r.label(p, q.cube),
                  row_cells(ctx, q, metric, cols, key, f"p{i}", total_key="total", product=p, bm=bm),
                  metric=metric, entity=q.entity, code=code_of(p))
             for i, p in enumerate(shown)]
    if rest and other:
        cells: list[Expr | None] = []
        for j, c in enumerate(cols):
            if c in ("value",) + BASES:
                cells.append(Sub(Ref(ck(key, "total", j)),
                                 Sum(tuple(Ref(ck(key, f"p{i}", j)) for i in range(len(shown))))))
            elif c == GAP:
                cells.append(None)
            else:
                cells.append(derived(c, cols, key, "other", total_key="total", j=j))
        lines.append(Line("other", "Other", cells, metric=metric))
    lines.append(Line("total", PRODUCTS_TOTAL, row_cells(ctx, q, metric, cols, key, "total", total_key="total",
                                                         bm=bm),
                      kind="grand", metric=metric, entity=q.entity))
    return lines
