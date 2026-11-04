# engine/smartass/pipeline.py
"""fin_analyze in three steps. prepare(): validate, guard, fetch every grid request in parallel (HTTP, no Excel),
compute the model. excel_job(): runs in the user's Excel (v1 host job) - target, Q<n>, Data + Answer sheets,
styling, readback, spec saved in _svmeta, sheet order, save our own file. reply(): the result dict the agent
quotes from (key_values only - never numbers it computed itself)."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field

from .model.analyze import Result, compute, plan_requests
from .model.frame import Frame
from .model.requests import GridRequest
from .model.spec import AnalysisSpec, LinesBlock
from .model.validate import GuardError

MAX_PARALLEL = 6


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


def prepare(raw: dict, resolver, provider) -> Prepared:
    spec = AnalysisSpec.model_validate(raw)
    if any(isinstance(b, LinesBlock) for b in spec.blocks):
        raise GuardError("the 'lines' block needs fin_drill, which arrives in Plan 1C",
                         "leave the lines block out for now")
    closed = provider.closed_months()
    reqs = plan_requests(spec, resolver, closed)
    with ThreadPoolExecutor(max_workers=MAX_PARALLEL) as ex:
        got = list(ex.map(provider.fetch, reqs))
    frame = Frame([f for g in got for f in g.facts])
    warnings = [w for g in got for w in g.warnings]
    routes: dict[str, int] = {}
    for g in got:
        routes[g.route] = routes.get(g.route, 0) + 1
    result = compute(spec, resolver, closed, frame)
    aggregation = {m: resolver.aggregation(m) for q in spec.queries for m in q.metrics}
    return Prepared(spec, raw, closed, reqs, frame, result, warnings, routes, sum(g.seconds for g in got),
                    aggregation)


def reply(pre: Prepared, q: int, workbook: str, sheets: list[str], mismatches: list[str], lint: list[str]) -> dict:
    res = pre.result
    return {
        "q": q,
        "workbook": workbook,
        "sheets": sheets,
        "key_values": res.key_values,
        "material": res.material,
        "findings": [{"severity": f.severity, "text": f.text, "where": f.where} for f in res.findings],
        "checks": {"sheet_equals_model": "OK" if not mismatches else f"DIFF in {len(mismatches)} cells"},
        "mismatches": mismatches[:20],
        "warnings": pre.warnings + lint,
        "routes": pre.routes,
        "closed": {c: f"{m[1]:02d}/{m[0]}" for c, m in pre.closed.items()},
        "fetch_seconds": round(pre.fetch_seconds, 2),
        "saved_user_workbook": False,
    }
