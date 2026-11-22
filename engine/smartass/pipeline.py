# engine/smartass/pipeline.py
"""fin_analyze in three steps. prepare(): validate, guard, fetch every grid request in parallel (HTTP, no Excel),
compute the model. The Excel job (exceljob.py: excel_job, redraw_job) draws it in the user's Excel. reply(): the
result dict the agent quotes from (key_values only - never numbers it computed itself)."""
from __future__ import annotations

import contextlib
import logging
import re
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field

from .model.analyze import Result, compute, plan_requests
from .model.blocks.concept import concept_caveats
from .model.chain import complete_chain, restore_chain
from .model.frame import Frame
from .model.periods import parse
from .model.requests import GridRequest
from .model.spec import AnalysisSpec, LinesBlock
from .model.validate import GuardError
from .epm.svhttp import SvHttpError
from .render.excel import sheet_status

MAX_PARALLEL = 6
log = logging.getLogger("smartview.pipeline")


@dataclass
class Prepared:
    spec: AnalysisSpec
    raw: dict
    closed: dict
    requests: list[GridRequest]
    frame: Frame
    result: Result
    warnings: list[str] = field(default_factory=list)
    routes: dict[str, int] = field(default_factory=dict)
    fetch_seconds: float = 0.0
    aggregation: dict[str, str] = field(default_factory=dict)
    failed: list[GridRequest] = field(default_factory=list)          # grids neither engine could fetch
    failed_errors: list[str] = field(default_factory=list)
    fetched: dict = field(default_factory=dict)       # GridRequest -> FetchResult the HTTP route fetched (reused by the fallback round)
    caveats: list[dict] = field(default_factory=list)  # {id, text, severity} of the facts the answer touched


def prepare(raw: dict, resolver, provider, closed: dict | None = None, lines: dict | None = None,
            complete: bool = False) -> Prepared:
    """complete (new questions and follow-up patches): the spec's chained tables get the whole P&L chain first
    (model/chain.py; a query a patch resent without the chain gets it back); the completed spec is the one returned
    as raw, so it is stored and refresh, patch and redraw see the chain."""
    if complete:
        raw = restore_chain(raw)
    spec = AnalysisSpec.model_validate(raw)
    if complete:
        spec = complete_chain(spec, resolver)
        raw = spec.model_dump()
    lines = lines or {}
    for i, b in enumerate(spec.blocks):
        if not isinstance(b, LinesBlock):
            continue
        ls = lines.get(f"b{i}")
        if ls is None:
            raise GuardError(f"block {i} (lines) has no ledger lines", "run it through fin_analyze or fin_drill, "
                             "which drill the cell first")
        if not ls.rows:
            why = f" ({'; '.join(ls.warnings)})" if ls.warnings else ""
            raise GuardError(f"no ledger lines behind {' / '.join(b.cell.values())}{why}",
                             "check the cell (account, entity, period) or the cube value; FinLite holds the lines")
    if closed is None:
        closed = provider.closed_months()
    canonical = getattr(resolver, "canonical", None)
    if canonical is not None:                   # atlas spelling: requests and SUMIFS then use one name
        for q in spec.queries:
            if resolver.exists(q.entity, q.cube):
                q.entity = canonical(q.entity, q.cube)
            q.products = [canonical(x, q.cube) if resolver.exists(x, q.cube) else x for x in q.products]
    reqs = plan_requests(spec, resolver, closed)
    def fetch_one(req):
        try:
            return provider.fetch(req), None
        except SvHttpError as e:        # only the HTTP route's failure; a wrong member (InputError) is wrong on every route
            return None, str(e)

    with ThreadPoolExecutor(max_workers=MAX_PARALLEL) as ex:
        fetched = list(ex.map(fetch_one, reqs))
    failed = [r for r, (g, _) in zip(reqs, fetched) if g is None]
    failed_errors = [err for g, err in fetched if g is None]
    got = [g for g, _ in fetched if g is not None]
    frame = Frame([f for g in got for f in g.facts])
    frame.lines = dict(lines)
    warnings = [w for g in got for w in g.warnings]
    if failed:
        warnings.append(f"{len(failed)} grid(s) could not be fetched over HTTP: {'; '.join(failed_errors[:3])}")
    routes: dict[str, int] = {}
    for g in got:
        routes[g.route] = routes.get(g.route, 0) + 1
    second = [g for g in got if g.route == "rest"]
    if second:                          # one warning, however many grids: the reply's routes count them
        warnings.append(f"{len(second)} grid(s) were fetched over Planning REST, the second engine: the Smart View "
                        f"provider route could not fetch them ({second[0].fallback}). Over REST a cell you have no "
                        "access to cannot be told apart from an empty one (Oracle exports only cells you may read and "
                        "write): such a cell is missing, not #NoAccess")
    left_warnings, left_caveats = left_out(spec, got, resolver)
    warnings += left_warnings
    result = compute(spec, resolver, closed, frame)
    aggregation = {m: resolver.aggregation(m) for q in spec.queries for m in q.metrics}
    return Prepared(spec, raw, closed, reqs, frame, result, warnings, routes, sum(g.seconds for g in got),
                    aggregation, failed, failed_errors, {g.request: g for g in got},
                    caveats(spec, reqs, resolver) + concept_caveats(spec, resolver) + left_caveats)


def left_out(spec: AnalysisSpec, got: list, resolver) -> tuple[list[str], list[dict]]:
    """One warning and one warn caveat (the Answer's Caveats, the Sources sheet's Known limits) per member the second
    engine left out (FetchResult.dropped: EPM refused it over Planning REST), however many requests asked it. It has
    no fact - the path of a missing cell - so a total over several members shows the others' sum: the text says it is
    INCOMPLETE and names the spec's metrics that pick the member (resolver.members)."""
    texts: list[str] = []
    cavs: list[dict] = []
    for cube, dim, member in dict.fromkeys(d for g in got for d in getattr(g, "dropped", ())):
        metrics = []
        for q in spec.queries:
            for m in q.metrics:
                with contextlib.suppress(Exception):          # a metric the resolver cannot place names nothing
                    if q.cube == cube and member in resolver.members(m, cube) and m not in metrics:
                        metrics.append(m)
        which = f" ({', '.join(metrics)})" if metrics else ""
        text = (f"{cube}: EPM refused {dim} member {member} over Planning REST (it does not exist there, or you have "
                f"no access to it), so it was left out: its figures are missing from this answer and every total that "
                f"includes it{which} is INCOMPLETE - check the member or rebuild the atlas (`smartass atlas build`)")
        texts.append(text)
        cavs.append({"id": f"left-out:{cube}:{member}", "text": text, "severity": "warn"})
    return texts, cavs


_YEAR = re.compile(r"(?:Yr|FY) ?(\d{2})")


def _cell_months(period: str) -> list:
    """A lines cell's period as months: 'Sep 26' -> [Sep 26], 'Yr 26' -> the year; anything else -> none."""
    with contextlib.suppress(ValueError):
        return [parse(period)]
    hit = _YEAR.fullmatch(period.strip())
    return [(2000 + int(hit.group(1)), m) for m in range(1, 13)] if hit else []


def caveats(spec: AnalysisSpec, reqs: list[GridRequest], resolver) -> list[dict]:
    """The resolver's caveats for what the answer touched: the entities and accounts of every grid request and lines
    cell (per cube; FinLite shares the FinMax members, so it counts for both), the metrics of the queries and the
    months fetched. A resolver without caveats (a test double) gives none."""
    fn = getattr(resolver, "caveats", None)
    if fn is None:
        return []
    touched: dict[str, set[str]] = {}
    months: set = set()

    def touch(cube: str, names) -> None:
        for c in ({cube, "FinMax"} if cube == "FinLite" else {cube}):
            touched.setdefault(c, set()).update(names)
    for r in reqs:
        touch(r.cube, (*r.entities, *r.accounts, *r.products))
        for per in r.periods:
            with contextlib.suppress(ValueError):
                months.add(parse(per))
    for b in spec.blocks:
        if isinstance(b, LinesBlock):
            touch("FinLite", [v for k, v in b.cell.items() if k in ("All Accounts", "All Entities")])
            months.update(_cell_months(b.cell.get("All Time Periods", "")))
    metric = getattr(resolver, "metric", None)
    ids = {m.id for q in spec.queries for name in q.metrics if metric and (m := metric(name)) is not None}
    out = fn(touched, ids, months)
    cards = getattr(resolver, "card_caveats", None)
    if cards is not None:                   # a level read as a flow is the commonest wrong answer: say it
        have = {c["id"] for c in out}
        out = out + [c for c in cards(touched) if c["id"] not in have]
    return out


def run_prepare(raw: dict, provider, lines: dict | None = None, complete: bool = False) -> Prepared:
    """prepare() for a worker thread: everything that touches sqlite (knowledge.db and the atlas) is opened, used and
    closed here, in the calling thread. knowledge.db is built first when it is missing (said in the warnings). Closed
    months are fetched once. Nothing after prepare needs the resolver: the labels are already in the BlockOut lines
    and the caveats in Prepared."""
    from .core import config
    from .knowledge.refresh import note_closed
    from .knowledge.resolver import open_resolver
    closed = provider.closed_months()
    with contextlib.suppress(Exception):         # fin_status's 'knowledge refresh due'; never stops an answer
        note_closed(config.HOME, closed["FinMax"])
    notes: list[str] = []
    with contextlib.closing(open_resolver(closed["FinPlan"][0] % 100, notes)) as resolver:
        pre = prepare(raw, resolver, provider, closed=closed, lines=lines, complete=complete)
    pre.warnings[:0] = notes
    return pre


def reply(pre: Prepared, q: int, workbook: str, sheets: list[str], mismatches: list[str], lint: list[str],
          layout: list[str] | None = None, charts: list[str] | None = None) -> dict:
    """layout: the layout check's findings (render/check.py), each with its fix; charts: the titles drawn."""
    res = pre.result
    checks = {"sheet_equals_model": sheet_status(mismatches)}
    out = {
        "q": q,
        "workbook": workbook,
        "sheets": sheets,
        "key_values": res.key_values,
        "material": res.material,
        "findings": [{"severity": f.severity, "text": f.text, "where": f.where} for f in res.findings],
        "checks": checks,
        "layout": list(layout or []),
        "charts": list(charts or []),
        "mismatches": mismatches[:20],
        "warnings": pre.warnings + lint,
        "caveats": list(pre.caveats),
        "routes": pre.routes,
        "closed": {c: f"{m[1]:02d}/{m[0]}" for c, m in pre.closed.items()},
        "fetch_seconds": round(pre.fetch_seconds, 2),
        "saved_user_workbook": False,
    }
    if pre.spec.plan.depth in ("standard", "deep"):         # the critic pass (sub-spec 2 §5) reads this
        from .model.critic import brief
        out["critic_brief"] = brief(pre.spec, res, list(pre.caveats), checks, stock=_stock(pre))
    return out


def _stock(pre: Prepared) -> set[str]:
    """Stock metrics (FTE): shown with one decimal (tall1), never summed over months."""
    return {m for m, agg in pre.aggregation.items() if agg == "last"}
