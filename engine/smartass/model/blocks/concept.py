"""concept: a product concept (knowledge/semantic/concepts/<id>.toml) on its query - one section with a heading line
per concept block and one line per item, the total and the concept's checks.

  booked      Pick(account, products) at the query entity (or account x entities, no product)
  attributed  base x key share; the base is the item's own amount, the key share 100 % unless it has one
  estimated   base x key share; the base is the shared amount, the key share num / den (or another item's key)
  sum         SUM of other items
  missing     no value; its note becomes a caveat

The base and the key share sit in labelled helper columns (Base, Key share, and the same for last year), so the line
is =<base>*<key share> and no constant ever sits inside a formula (tn-formelstil); a 100 % key is a typed input.
A check line that does not come to zero is a finding, never a stop: a concept is an estimate.

The concept's amounts are fetched in their own grids, tagged in the Data table's Query column '<query>:<id>:p' (account
x entity x products) and '<query>:<id>:e' (account x entities, no product), so they never mix with the query's own
metric grid (the cube's product total)."""
from __future__ import annotations

from dataclasses import dataclass

from ...knowledge.semantic import expr_terms
from ..context import Ctx
from ..expr import Const, Expr, Mul, Neg, Ratio, Ref, Sub, Sum, pick
from ..periods import label
from ..series import months, scenario
from ..spec import ConceptBlock, Query
from .common import HEAD, BlockOut, Line, ck, derived, scenario_heads

BALANCE = "Periodic"
HELPERS = {"value": ("base", "key"), "ly": ("base_ly", "key_ly")}
HELPER_HEAD = {"base": "Base", "key": "Key share", "base_ly": "Base last year", "key_ly": "Key share last year"}


def tags(q: Query, concept) -> tuple[str, str]:
    """The Query-column tags of the concept's product grid and entity grid."""
    return f"{q.id}:{concept.id}:p", f"{q.id}:{concept.id}:e"


@dataclass(frozen=True)
class ConceptGrid:
    tag: str
    accounts: tuple[str, ...]
    entities: tuple[str, ...]
    products: tuple[str, ...]


def grids(q: Query, concept) -> list[ConceptGrid]:
    """What the concept reads: one grid for the amounts on products (at the query entity), one for the amounts on
    entities (no product dimension). Each only when the concept has such amounts."""
    ptag, etag = tags(q, concept)
    prod = [p for p in concept.picks() if p.products]
    ents = [p for p in concept.picks() if p.entities]
    out = []
    if prod:
        out.append(ConceptGrid(ptag, tuple(dict.fromkeys(p.account for p in prod)), (q.entity,),
                               tuple(dict.fromkeys(x for p in prod for x in p.products))))
    if ents:
        out.append(ConceptGrid(etag, tuple(dict.fromkeys(p.account for p in ents)),
                               tuple(dict.fromkeys(x for p in ents for x in p.entities)), ()))
    return out


def series_of(q: Query) -> list[str]:
    return ["value"] + (["ly"] if "ly" in q.compare else [])


def _cols(q: Query, concept) -> list[str]:
    cols = ["value"] + (["ly", "var", "var_pct"] if "ly" in q.compare else [])
    if any(it.kind in ("attributed", "estimated") for it in concept.items()):
        for s in series_of(q):
            cols += list(HELPERS[s])
    return cols


def build_concept(b: ConceptBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    c = ctx.resolver.concept(b.concept)
    cols = _cols(q, c)
    heads, _ = scenario_heads(ctx, q, cols)
    out = BlockOut(key, "concept", b.title or c.title, [h or HELPER_HEAD[x] for x, h in zip(cols, heads)],
                   cols=cols, query=q.id, keys=[HEAD.get(x) or HELPER_HEAD[x] for x in cols])
    ptag, etag = tags(q, c)
    closed = ctx.closed_for(q.cube)

    def amount(p, series: str) -> Expr:
        scen, ver = scenario(q.cube, series, ctx.resolver)
        f = {"Cube": (q.cube,), "Query": (ptag if p.products else etag,), "Account": (p.account,),
             "Entity": (q.entity,) if p.products else tuple(p.entities)}
        if p.products:
            f["Product"] = tuple(p.products)
        f |= {"Period": tuple(label(m) for m in months(q, series, closed)), "Scenario": (scen,)}
        if ver:
            f["Version"] = (ver,)
        f["Balance"] = (BALANCE,)
        return pick(f)

    def col(name: str) -> int:
        return cols.index(name)

    def metric_of(p) -> str | None:
        """The metric of the item's account (in the query cube, else FinMax), so the line gets its materiality."""
        if p is None or not p.account:
            return None
        return ctx.resolver.metric_of(p.account, q.cube) or ctx.resolver.metric_of(p.account, "FinMax")

    def ref(item_key: str, j: int) -> Ref:
        return Ref(ck(key, "total" if item_key == "total" else f"i_{item_key}", j))

    def terms(expr: str, j: int, src: Expr | None = None) -> Expr:
        """An expression over item keys (and $src, the check's source amount) in column j."""
        parts = [(sign, src if name == "$src" else ref(name, j)) for sign, name in expr_terms(expr)]
        e = parts[0][1] if parts[0][0] > 0 else Neg(parts[0][1])
        for sign, t in parts[1:]:
            e = Sum((e, t)) if sign > 0 else Sub(e, t)
        return e

    def derive(cells: list, lkey: str) -> list:
        for j, x in enumerate(cols):
            if x in ("var", "var_pct"):
                cells[j] = derived(x, cols, key, lkey, j=j)
        return cells

    for gi, blk in enumerate(c.block):
        out.lines.append(Line(f"g{gi}", blk.title, [None] * len(cols)))
        for it in blk.item:
            lkey = f"i_{it.key}"
            cells: list[Expr | None] = [None] * len(cols)
            lab = it.label
            if it.kind == "booked":
                for s in series_of(q):
                    cells[col(s)] = amount(it.pick(), s)
            elif it.kind == "sum":
                for s in series_of(q):
                    cells[col(s)] = Sum(tuple(ref(k, col(s)) for k in it.of))
            elif it.kind in ("attributed", "estimated"):
                for s in series_of(q):
                    bcol, kcol = (col(h) for h in HELPERS[s])
                    cells[bcol] = amount(it.pick(), s)
                    ks = it.key_share
                    if ks is None:
                        cells[kcol] = Const(1.0)
                    elif ks.same_as:
                        cells[kcol] = Ref(ck(key, f"i_{ks.same_as}", kcol))
                    else:
                        num = terms(ks.num.expr, col(s)) if ks.num.expr else amount(ks.num, s)
                        cells[kcol] = Ratio(num, amount(ks.den, s))
                    cells[col(s)] = Mul(Ref(ck(key, lkey, bcol)), Ref(ck(key, lkey, kcol)))
            else:                                                   # missing
                lab = f"{it.label} (not in {q.cube})"
            if it.kind != "missing":
                derive(cells, lkey)
            kind = "sum" if it.kind == "sum" else "line"
            out.lines.append(Line(lkey, lab, cells, kind=kind, level=1, entity=q.entity,
                                  metric=metric_of(it.pick())))
    total: list[Expr | None] = [None] * len(cols)
    for s in series_of(q):
        total[col(s)] = Sum(tuple(ref(k, col(s)) for k in c.total.of))
    out.lines.append(Line("total", c.total.label, derive(total, "total"), kind="grand", entity=q.entity))
    for i, ch in enumerate(c.check):
        cells = [None] * len(cols)
        for s in series_of(q):
            cells[col(s)] = terms(ch.expr, col(s), amount(ch.source, s) if ch.source else None)
        out.lines.append(Line(f"c{i}", f"Check: {ch.label}", cells, kind="check"))
    return out


def concept_findings(outs: list[BlockOut], cells) -> list:
    """A concept check that does not come to zero: a finding (warning), never a stop - a concept is an estimate."""
    from ..validate import TOLERANCE, Finding      # validate imports the blocks: import it when used
    hits: list[Finding] = []
    for o in outs:
        if o.type != "concept":
            continue
        for ln in o.lines:
            if ln.kind != "check":
                continue
            off = [(o.headers[j], v) for j, e in enumerate(ln.cells)
                   if e is not None and (v := cells.get(ck(o.key, ln.key, j))) is not None and abs(v) > TOLERANCE]
            if off:
                by = ", ".join(f"{v:,.0f} NOK ({h})" for h, v in off)
                hits.append(Finding("warning", f"{o.title}: {ln.label} is off by {by}; the concept is an estimate - "
                                               "check its definition before relying on the total", o.key))
    return hits


def concept_caveats(spec, resolver) -> list[dict]:
    """Info caveats for every concept block: the concept's notes, its missing items and how each estimated item
    is estimated."""
    out: list[dict] = []
    seen: set[str] = set()
    for b in spec.blocks:
        if not isinstance(b, ConceptBlock):
            continue
        c = resolver.concept(b.concept)
        if c is None or c.id in seen:
            continue
        seen.add(c.id)
        texts = list(c.notes)
        for it in c.items():
            if it.kind == "missing":
                texts.append(f"{it.label}: {it.note}" if it.note else f"{it.label}: not in the cube")
            elif it.kind == "estimated":
                ks = it.key_share
                how = ks.text or (f"the key of {ks.same_as}" if ks.same_as else "num / den")
                texts.append(f"{it.label}: estimated = base ({it.base.label}) × key ({how}).")
        out += [{"id": f"concept_{c.id}_{n}", "text": t, "severity": "info"} for n, t in enumerate(texts)]
    return out


def concept_requests(spec, r) -> list[tuple[Query, ConceptGrid, str]]:
    """(query, grid, series) for every concept block, without duplicates (the same concept twice on one query reads
    the same grids)."""
    out: list[tuple[Query, ConceptGrid, str]] = []
    seen: set[tuple] = set()
    by_id = {q.id: q for q in spec.queries}
    for b in spec.blocks:
        if not isinstance(b, ConceptBlock):
            continue
        q = by_id[b.query]
        c = r.concept(b.concept)
        if c is None:
            continue
        for g in grids(q, c):
            for s in series_of(q):
                if (g.tag, s) not in seen:
                    seen.add((g.tag, s))
                    out.append((q, g, s))
    return out
