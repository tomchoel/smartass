# engine/smartass/model/blocks/drivers.py
"""drivers: decompose a metric's variance down the entity tree. At each level the children are ranked by absolute
variance and shown until they explain 80 % of the gross movement (at most 3); each shown child is expanded the same
way until `depth`. With drill, the leaves reached are returned as drill requests (Plan 1B turns them into ledger
lines), so commentary can rest on data."""
from __future__ import annotations

from dataclasses import dataclass

from ..context import Ctx
from ..expr import Neg, Ref, Sub, Sum
from ..periods import label
from ..series import months
from ..spec import DriversBlock
from .common import BlockOut, Line, ck, code_of, counted, derived, metric_code, metric_label, row_cells, scenario_heads

@dataclass(frozen=True)
class DrillRequest:
    """A leaf the data layer drills into ledger lines: the cell's accounts and months, and the base to compare."""
    query: str
    cube: str
    entity: str
    metric: str
    accounts: tuple[str, ...]
    periods: tuple[str, ...]
    base_periods: tuple[str, ...]
    base: str


COVER = 0.8
MAX_PER_LEVEL = 3


def build_drivers(b: DriversBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    r = ctx.resolver
    cols = ["value", b.base, "var", "var_pct"]
    heads, keys = scenario_heads(ctx, q, cols)
    by_label = {"entities": "", "accounts": " by account", "products": " by product"}.get(b.by, "")
    out = BlockOut(key, "drivers", b.title or f"Drivers of {metric_label(r, b.metric, q.cube)}{by_label}", heads,
                   cols=cols, query=q.id, keys=keys)

    def _drill_req(entity: str, accounts: tuple[str, ...]) -> None:
        if not b.drill:
            return
        closed = ctx.closed_for(q.cube)
        last = r.aggregation(b.metric) == "last"
        def cut(ms: list) -> tuple[str, ...]:
            return tuple(label(m) for m in (ms[-1:] if last else ms))
        out.drills.append(DrillRequest(
            q.id, q.cube, entity, b.metric, accounts,
            cut(months(q, "value", closed)), cut(months(q, b.base, closed)), b.base))

    def variance_of(metric_tok: str, entity: str = q.entity, product: str | None = None) -> float:
        return ctx.value(Sub(ctx.pick(q, metric_tok, "value", entity, product=product),
                             ctx.pick(q, metric_tok, b.base, entity, product=product))) or 0.0

    def walk_entities(entity: str, level: int, prefix: str) -> None:
        kids = tuple(counted(ctx, q, r.children(entity, q.cube), entity))
        if level >= b.depth or not kids:
            _drill_req(entity, tuple(r.members(b.metric, q.cube)))
            return
        if len(kids) == 1:
            walk_entities(kids[0], level, prefix)
            return
        scored = sorted(((variance_of(b.metric, entity=k), k) for k in kids), key=lambda t: -abs(t[0]))
        gross = sum(abs(v) for v, _ in scored) or 1.0
        covered, skipped = 0.0, []
        for i, (v, k) in enumerate(scored):
            if covered / gross >= b.cover or i >= b.max_per_level:
                skipped = scored[i:]
                break
            lkey = f"{prefix}{i}"
            out.lines.append(Line(lkey, r.label(k, q.cube), row_cells(ctx, q, b.metric, cols, key, lkey, k),
                                  level=level + 1, metric=b.metric, entity=k, code=code_of(k)))
            walk_entities(k, level + 1, lkey + "_")
            covered += abs(v)
        if b.include_other and skipped:
            others = tuple(k for _, k in skipped)
            okey = f"{prefix}other"
            out.lines.append(Line(okey, f"Other ({len(others)})",
                                  row_cells(ctx, q, b.metric, cols, key, okey, entity=others),
                                  level=level + 1, metric=b.metric))

    def walk_accounts(acc: str, level: int, prefix: str) -> None:
        raw_kids = [k for k in r.children(acc, q.cube) if r.weight(k, q.cube, acc) != 0]
        if level >= b.depth or not raw_kids:
            _drill_req(q.entity, (acc,))
            return
        if len(raw_kids) == 1:
            walk_accounts(raw_kids[0], level, prefix)
            return
        scored = sorted(((variance_of(k), k) for k in raw_kids), key=lambda t: -abs(t[0]))
        gross = sum(abs(v) for v, _ in scored) or 1.0
        covered, skipped, shown_kids = 0.0, [], []
        for i, (v, k) in enumerate(scored):
            if covered / gross >= b.cover or i >= b.max_per_level:
                skipped = scored[i:]
                break
            lkey = f"{prefix}{i}"
            shown_kids.append((lkey, k))
            out.lines.append(Line(lkey, r.label(k, q.cube), row_cells(ctx, q, k, cols, key, lkey),
                                  level=level + 1, metric=k, entity=q.entity, code=code_of(k)))
            walk_accounts(k, level + 1, lkey + "_")
            covered += abs(v)
        if b.include_other and skipped:
            okey = f"{prefix}other"
            cells: list = []
            for ci, c in enumerate(cols):
                if c in ("value", b.base):
                    terms = tuple(
                        Neg(Ref(ck(key, lk, ci))) if r.weight(kid, q.cube, acc) < 0 else Ref(ck(key, lk, ci))
                        for lk, kid in shown_kids
                    )
                    cells.append(Sub(ctx.pick(q, acc, c), Sum(terms)))
                else:
                    cells.append(derived(c, cols, key, okey, j=ci))
            out.lines.append(Line(okey, f"Other accounts ({len(skipped)})", cells, level=level + 1))

    def walk_products() -> None:
        prods = list(ctx.dedup_products(q))
        scored = sorted(((variance_of(b.metric, product=p), p) for p in prods), key=lambda t: -abs(t[0]))
        gross = sum(abs(v) for v, _ in scored) or 1.0
        covered = 0.0
        for i, (v, p) in enumerate(scored):
            if covered / gross >= b.cover or i >= b.max_per_level:
                break
            lkey = f"d{i}"
            out.lines.append(Line(lkey, r.label(p, q.cube), row_cells(ctx, q, b.metric, cols, key, lkey, product=p),
                                  level=1, metric=b.metric, entity=q.entity, code=code_of(p)))
            covered += abs(v)

    if b.by == "accounts":
        roots = tuple(r.members(b.metric, q.cube))
        top_code = code_of(roots[0]) if len(roots) == 1 else metric_code(r, b.metric, q.cube)
        out.lines.append(Line("top", metric_label(r, b.metric, q.cube), row_cells(ctx, q, b.metric, cols, key, "top"),
                              level=0, metric=b.metric, entity=q.entity, code=top_code))
        if len(roots) == 1:
            walk_accounts(roots[0], 0, "d")
        else:
            scored = sorted(((variance_of(acc), acc) for acc in roots), key=lambda t: -abs(t[0]))
            for i, (_, acc) in enumerate(scored[:b.max_per_level]):
                lkey = f"d{i}"
                out.lines.append(Line(lkey, r.label(acc, q.cube), row_cells(ctx, q, acc, cols, key, lkey),
                                      level=1, metric=acc, entity=q.entity, code=code_of(acc)))
                walk_accounts(acc, 1, lkey + "_")
    elif b.by == "products":
        out.lines.append(Line("top", metric_label(r, b.metric, q.cube), row_cells(ctx, q, b.metric, cols, key, "top"),
                              level=0, metric=b.metric, entity=q.entity, code=code_of(q.entity)))
        walk_products()
    else:
        out.lines.append(Line("top", r.label(q.entity, q.cube), row_cells(ctx, q, b.metric, cols, key, "top"),
                              level=0, metric=b.metric, entity=q.entity, code=code_of(q.entity)))
        walk_entities(q.entity, 0, "d")
    return out


def leaf_keys(out: BlockOut) -> set[str]:
    """The lines at the bottom of the shown tree (no deeper line right under them): the drivers a chart and the key
    points name. The top line and the expanded middle levels are sums of these and would count them twice."""
    lines = [ln for ln in out.lines if ln.kind != "check"]
    keys = {ln.key for i, ln in enumerate(lines)
            if ln.level > 0 and (i + 1 == len(lines) or lines[i + 1].level <= ln.level)}
    return keys or {ln.key for ln in lines}
