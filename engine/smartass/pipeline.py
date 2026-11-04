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


def prepare(raw: dict, resolver, provider, closed: dict | None = None) -> Prepared:
    spec = AnalysisSpec.model_validate(raw)
    if any(isinstance(b, LinesBlock) for b in spec.blocks):
        raise GuardError("the 'lines' block needs fin_drill, which arrives in Plan 1C",
                         "leave the lines block out for now")
    if closed is None:
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


def run_prepare(raw: dict, provider) -> Prepared:
    """prepare() for a worker thread: everything that touches sqlite (the atlas connection) is opened, used and closed
    here, in the calling thread. Closed months are fetched once. Nothing after prepare needs the resolver: the
    labels are already in the BlockOut lines."""
    import contextlib

    from smartview_mcp import atlas

    from .knowledge.atlas_resolver import AtlasResolver
    from .knowledge.interim import catalog_metrics
    closed = provider.closed_months()
    with contextlib.closing(atlas.connect()) as db:
        resolver = AtlasResolver(db, catalog_metrics(), forecast_year=closed["FinPlan"][0] % 100)
        return prepare(raw, resolver, provider, closed=closed)


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
# ---------------------------------------------------------------- the Excel job (runs on a v1 user-Excel job thread)
from .render.excel import read_block, write_answer, write_data  # noqa: E402
from .render.layout import layout  # noqa: E402
from .render.readback import compare  # noqa: E402

DATA_TAB = 0x808080      # grey tab for data sheets (BGR)


def _target(h, target: str, title: str):
    from smartview_mcp.excel import target as xltarget
    return xltarget.resolve_target(h, target, allow_archive=False, title=title)


def _next_q(t) -> int:
    from smartview_mcp.excel import sheets
    return sheets.next_question(t)


def _add_sheet(t, name: str, q: int, color=None):
    from smartview_mcp.excel import sheets
    return sheets.add_sheet(t, name, after_q=q, color=color)


def _calculate(t) -> None:
    from smartview_mcp.excel import sheets
    from smartview_mcp.excel.jobs import com
    com(lambda: t.app.Calculate())
    sheets.wait_calculated(t.app)


def _save_spec(t, q: int, raw: dict) -> None:
    from smartview_mcp.excel import sheets
    sheets.save_spec(t, q, {"v2": raw}, ids={}, owned=None)


def _arrange(t) -> None:
    from smartview_mcp import sheetorder
    sheetorder.arrange_sheets(t.wb, mode="full")


def _save(t) -> None:
    from smartview_mcp.excel.jobs import com
    if getattr(t, "created", False):           # only the engine's own analysis file, never the user's workbook
        com(lambda: t.wb.Save())


def _rollback(t) -> None:
    from smartview_mcp.excel import sheets
    sheets.rollback(t)


def excel_job(h, pre: Prepared, target: str = "session") -> dict:
    t = _target(h, target, pre.spec.question[:60])
    try:
        q = _next_q(t)
        ws_answer = _add_sheet(t, f"Q{q} Answer", q)
        ws_data = _add_sheet(t, f"Q{q} Data", q, DATA_TAB)
        write_data(ws_data, pre.frame, q)
        cubes = sorted({x.cube for x in pre.spec.queries})
        closed = ", ".join(f"{c} {m[1]:02d}/{m[0]}" for c, m in sorted(pre.closed.items()) if c in cubes)
        plan = layout(q, pre.spec, pre.result, stock=_stock(pre), subtitle=f"NOK · {' + '.join(cubes)} · closed {closed}")
        lint = write_answer(ws_answer, plan, pre.warnings)
        _calculate(t)
        mismatches = compare(plan.addr, pre.result.cells, read_block(ws_answer, plan))
        _save_spec(t, q, pre.raw)
        _arrange(t)
        _save(t)
    except Exception:
        _rollback(t)
        raise
    return reply(pre, q, str(t.path), [f"Q{q} Answer", f"Q{q} Data"], mismatches, list(lint or []))


def _stock(pre: Prepared) -> set[str]:
    """Stock metrics (FTE): shown with one decimal (tall1), never summed over months."""
    return {m for m, agg in pre.aggregation.items() if agg == "last"}
