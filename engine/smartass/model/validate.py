# engine/smartass/model/validate.py
"""Guards. Hard on data (stop, with the fix): unknown entity, unknown or draft metric, an open month without
'preliminary', a children = parent check that does not reconcile. Soft on method (findings for the critic): no
last-year comparison, preliminary months. A total whose sign contradicts the metric is Critical, not a stop
(FinPlan's sign convention is not verified yet, so FinPlan is not checked)."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from .blocks.common import BlockOut, ck
from .context import Ctx
from .expr import Cells
from .periods import Month, label, resolve
from .resolver import Resolver, UnknownMetric
from .spec import AnalysisSpec

TOLERANCE = 1.0  # NOK


class GuardError(Exception):
    def __init__(self, reason: str, fix: str):
        super().__init__(f"{reason} - {fix}")
        self.reason, self.fix = reason, fix


@dataclass
class Finding:
    severity: Literal["critical", "warning", "info"]
    text: str
    where: str = ""


def guard_spec(spec: AnalysisSpec, r: Resolver, closed: Month) -> list[Finding]:
    soft: list[Finding] = []
    for q in spec.queries:
        if not r.exists(q.entity, q.cube):
            raise GuardError(f"unknown entity {q.entity!r} in {q.cube}", "find the member with fin_lookup")
        for m in q.metrics:
            try:
                r.members(m, q.cube)
            except UnknownMetric:
                raise GuardError(f"unknown metric or member {m!r} in {q.cube}", "find it with fin_lookup") from None
            if r.status(m) != "approved":
                raise GuardError(f"metric {m!r} is a draft",
                                 "approve it (smartass knowledge approve) or use an approved metric")
        late = [label(x) for x in resolve(q.periods, closed) if x > closed]
        if late and not q.preliminary:
            raise GuardError(f"query {q.id}: {', '.join(late)} is after the last closed month {label(closed)}",
                             f"end the period at {label(closed)} or set preliminary: true")
        if late:
            soft.append(Finding("warning", f"query {q.id}: {', '.join(late)} is preliminary (month not closed)",
                                q.id))
        if spec.plan.depth != "quick" and q.cube != "FinPlan" and "ly" not in q.compare:
            soft.append(Finding("warning", f"query {q.id} has no last-year comparison", q.id))
    return soft


def guard_result(outs: list[BlockOut], cells: Cells) -> None:
    for o in outs:
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
