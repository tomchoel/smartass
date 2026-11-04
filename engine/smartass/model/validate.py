# engine/smartass/model/validate.py
"""Guards. Hard on data (stop, with the fix): unknown entity, unknown or draft metric, an open month without
'preliminary', a children = parent check that does not reconcile. Soft on method (findings for the critic): no
last-year comparison, preliminary months. A total whose sign contradicts the metric is Critical, not a stop
(FinPlan's sign convention is not verified yet, so FinPlan is not checked)."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from .blocks.common import BlockOut, base_of, ck
from .context import Ctx
from .errors import GuardError
from .expr import Cells
from .periods import Closed, closed_for, label, resolve
from .resolver import Resolver, UnknownMetric
from .series import months
from .spec import AnalysisSpec, BridgeBlock, ConceptBlock, DriversBlock, KpiBlock, RankingBlock, TableBlock

TOLERANCE = 1.0  # NOK


@dataclass
class Finding:
    severity: Literal["critical", "warning", "info"]
    text: str
    where: str = ""


def _plan_bases(b) -> set[str]:
    """The plan series (budget / forecast) a block compares the value against."""
    if isinstance(b, TableBlock):
        return {b.columns[base_of(b.columns, j)] for j, c in enumerate(b.columns) if c in ("var", "var_pct")}             & {"budget", "forecast"}
    if isinstance(b, (KpiBlock, BridgeBlock, DriversBlock)):
        return {b.compare if isinstance(b, KpiBlock) else b.base} & {"budget", "forecast"}
    if isinstance(b, RankingBlock):
        return {b.compare} & {"budget", "forecast"}
    return set()


def _span(ms) -> str:
    return label(ms[0]) if len(ms) == 1 else f"{label(ms[0])}..{label(ms[-1])}"


def _like_for_like(spec: AnalysisSpec, q, cl) -> None:
    if not q.plan_periods:
        return
    asked = {s for b in spec.blocks if getattr(b, "query", None) == q.id for s in _plan_bases(b)}
    actual = months(q, "value", cl)
    for s in sorted(asked):
        plan = months(q, s, cl)
        if plan != actual:
            raise GuardError(f"query {q.id}: Actual covers {_span(actual)} but {s} covers {_span(plan)}",
                             "compare like-for-like months, or put the full-year plan in its own query")


def guard_spec(spec: AnalysisSpec, r: Resolver, closed: Closed) -> list[Finding]:
    soft: list[Finding] = []
    for q in spec.queries:
        if not r.exists(q.entity, q.cube):
            raise GuardError(f"unknown entity {q.entity!r} in {q.cube}", "find the member with fin_lookup")
        gone = [p for p in q.products if not r.exists(p, q.cube)]
        if gone:
            raise GuardError(f"unknown product(s) {', '.join(map(repr, gone))} in {q.cube}",
                             "find the product members with fin_lookup (FinMax/FinLite 'pd.SY.<code>')")
        for m in q.metrics:
            try:
                r.members(m, q.cube)
            except UnknownMetric:
                raise GuardError(f"unknown metric or member {m!r} in {q.cube}", "find it with fin_lookup") from None
            if r.status(m) != "approved":
                raise GuardError(f"metric {m!r} is a draft",
                                 "approve it (smartass knowledge approve) or use an approved metric")
        cl = closed_for(closed, q.cube)
        _like_for_like(spec, q, cl)
        late = [label(x) for x in resolve(q.periods, cl) if x > cl]
        if late and not q.preliminary:
            raise GuardError(f"query {q.id}: {', '.join(late)} is after the last closed month {label(cl)}",
                             f"end the period at {label(cl)} or set preliminary: true")
        if late:
            soft.append(Finding("warning", f"query {q.id}: {', '.join(late)} is preliminary (month not closed)",
                                q.id))
        if spec.plan.depth != "quick" and q.cube != "FinPlan" and "ly" not in q.compare:
            soft.append(Finding("warning", f"query {q.id} has no last-year comparison", q.id))
    by_id = {q.id: q for q in spec.queries}
    for b in spec.blocks:
        if isinstance(b, ConceptBlock):
            _guard_concept(b, by_id[b.query], r)
    return soft


def _guard_concept(b: ConceptBlock, q, r: Resolver) -> None:
    c = r.concept(b.concept)
    if c is None:
        raise GuardError(f"unknown concept {b.concept!r}", "find it with fin_lookup(kind='concept')")
    if c.status != "approved":
        raise GuardError(f"concept {c.id!r} is a draft", "approve it (smartass knowledge approve) first")
    gone = [x for xs in c.members().values() for x in xs if not r.exists(x, q.cube)]
    if gone:
        raise GuardError(f"concept {c.id}: {', '.join(gone[:10])} not in {q.cube}",
                         "fix the concept file (knowledge/semantic/concepts) or refresh the atlas")


def guard_result(outs: list[BlockOut], cells: Cells) -> None:
    for o in outs:
        if o.type == "concept":                # a concept's checks are findings (concept_findings), never a stop
            continue
        for ln in o.lines:
            if ln.kind != "check":
                continue
            for j, e in enumerate(ln.cells):
                v = cells.get(ck(o.key, ln.key, j)) if e is not None else None
                if v is not None and abs(v) > TOLERANCE:
                    raise GuardError(f"{o.title}: {ln.label} is off by {v:,.0f} NOK",
                                     "the children do not add up to the parent; check the entity tree with "
                                     "fin_lookup before using these numbers")


def sign_findings(ctx: Ctx) -> list[Finding]:
    out: list[Finding] = []
    for q in ctx.spec.queries:
        if q.cube == "FinPlan":
            continue
        for m in q.metrics:
            s = ctx.resolver.sign(m)
            if s == "neutral":
                continue
            v = ctx.value(ctx.pick(q, m, "value"))
            if v is None:
                continue
            if (s == "cost" and v > 0) or (s == "income" and v < 0):
                want = "negative" if s == "cost" else "positive"
                out.append(Finding("critical", f"{m} in query {q.id} is {v:,.0f}; a {s} metric should be {want} "
                                               f"in {q.cube}", q.id))
    return out
