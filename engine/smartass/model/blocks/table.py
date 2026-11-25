# engine/smartass/model/blocks/table.py
"""table: metric rows, entity rows (children of the query's entity) or product rows (the query's products) by
columns, optionally in column groups (one per period, e.g. month and YTD, separated by a gap column); kpi_tiles:
headline metrics with one comparison."""
from __future__ import annotations

from ..chain import CHAIN, KINDS, chained
from ..context import Ctx
from ..periods import resolve
from ..spec import KpiBlock, TableBlock
from .common import (GAP, BlockOut, Line, account_lines, add_material, code_of, counted, entity_lines, metric_code,
                     metric_label, negligible, product_lines, row_cells, scenario_heads)


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
    by = {"entities": "by unit", "products": "by product", "accounts": "by account",
          "entities_x_metrics": "by unit and metric"}.get(b.rows)
    if b.rows == "entities_x_metrics":
        default = f"{ctx.resolver.label(q.entity, q.cube)} — {by}"
    else:
        default = f"{metric_label(ctx.resolver, b.metric, q.cube)} {by}" if by and b.metric else "Table"
    out = BlockOut(key, "table", b.title or default, heads, cols=cols, query=q.id, keys=keys)
    if b.rows == "entities":
        out.lines = entity_lines(ctx, q, b.metric, cols, key, top=b.top, other=b.other, sort=b.sort, bm=bm_or_none)
    elif b.rows == "products":
        out.lines = product_lines(ctx, q, b.metric, cols, key, top=b.top, other=b.other, sort=b.sort,
                                  bm=bm_or_none)
    elif b.rows == "accounts":
        out.lines = account_lines(ctx, q, b.metric, cols, key, top=b.top, other=b.other, sort=b.sort,
                                  bm=bm_or_none)
    elif b.rows == "entities_x_metrics":
        out.lines = entities_x_metrics_lines(b, ctx, q, cols, key, bm_or_none)
    else:
        out.lines = metric_lines(b, ctx, q, cols, key, bm_or_none)
    if b.material:
        add_material(out, ctx.resolver)
    return out


def entities_x_metrics_lines(b: TableBlock, ctx: Ctx, q, cols: list[str], key: str, bm: list | None) -> list[Line]:
    """Compare each child entity (plus parent total) across multiple metrics side-by-side under clean headings."""
    r = ctx.resolver
    ms = list(b.metrics or q.metrics)
    kids = counted(ctx, q, list(r.children(q.entity, q.cube)))
    if b.sort != "none" or b.top:
        primary = ms[0]
        kids = [e for e in kids if any(not negligible(ctx, q, m, cols, entity=e) for m in ms)]
        def val(e: str) -> float:
            return ctx.value(ctx.pick(q, primary, "value", e)) or 0.0
        order = {"desc": lambda e: -val(e), "asc": val, "abs": lambda e: -abs(val(e)),
                 "none": lambda e: -abs(val(e))}[b.sort]
        kids.sort(key=order)
        if b.top:
            kids = kids[:b.top]
    units = [*kids, q.entity] if kids else [q.entity]
    lines: list[Line] = []
    pad = [None] * len(cols)
    for ei, ent in enumerate(units):
        is_total = ent == q.entity and bool(kids)
        lines.append(Line(f"h{ei}", r.label(ent, q.cube), list(pad), kind="heading", entity=ent, code=code_of(ent)))
        for mi, m in enumerate(ms):
            lkey = f"em{ei}_{mi}"
            lines.append(Line(lkey, metric_label(r, m, q.cube),
                              row_cells(ctx, q, m, cols, key, lkey, entity=ent, bm=bm),
                              kind="sum" if is_total else "line", level=1, metric=m, entity=ent,
                              code=metric_code(r, m, q.cube)))
    return lines


def metric_lines(b: TableBlock, ctx: Ctx, q, cols: list[str], key: str, bm: list | None) -> list[Line]:
    """One line per metric, keyed m{i} by its index in the table's metrics (stable for the goldens). A chained
    table shows the chain metrics in chain order (Gross margin and EBITDA before other items as sums, EBITDA as the
    grand total), then the other metrics as given."""
    r = ctx.resolver
    ms = list(b.metrics or q.metrics)
    labels = [metric_label(r, m, q.cube) for m in ms]
    order = list(range(len(ms)))
    chain = chained(b, labels)
    if chain:
        pos = {c: i for i, c in enumerate(CHAIN)}
        order.sort(key=lambda i: (0, pos[labels[i]]) if labels[i] in pos else (1, i))
    return [Line(f"m{i}", labels[i], row_cells(ctx, q, ms[i], cols, key, f"m{i}", bm=bm),
                 kind=KINDS.get(labels[i], "line") if chain else "line", metric=ms[i], entity=q.entity,
                 code=metric_code(r, ms[i], q.cube), added=ms[i] in b.chain_added)
            for i in order]


def build_kpi(b: KpiBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = ["value", b.compare, "var", "var_pct"]
    heads, keys = scenario_heads(ctx, q, cols)
    out = BlockOut(key, "kpi_tiles", b.title or "Key figures", heads, cols=cols, query=q.id, keys=keys)
    for i, m in enumerate(b.metrics):
        out.lines.append(Line(f"k{i}", metric_label(ctx.resolver, m, q.cube),
                              row_cells(ctx, q, m, cols, key, f"k{i}"), metric=m,
                              entity=q.entity, code=metric_code(ctx.resolver, m, q.cube)))
    return out
