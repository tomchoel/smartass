# engine/smartass/server.py
"""smartass2: the v2 MCP server (built beside v1 until v2 passes the test set). Four tools:
fin_analyze (new question or follow-up patch), fin_lookup, fin_status, fin_refresh. Error text (guard), Excel jobs (_call_user, progress notes) and the lifespan
(warm-up, closing the hidden Excel) are v1's via the legacy seam; answer files and session are v2's own."""
from __future__ import annotations

import asyncio
import contextlib
import datetime as dt

from mcp.server.mcpserver import Context, MCPServer

from smartview_mcp import atlas
from smartview_mcp import server as v1

from . import legacy
from .knowledge.atlas_resolver import AtlasResolver
from .knowledge.interim import catalog_metrics
from .followup import find_question, merge_patch, v2_questions
from .model.spec import AnalysisSpec
from .model.validate import GuardError
from .pipeline import RestoreError, excel_job, read_stored, redraw_job, run_prepare

TOOL_NAMES = ("fin_analyze", "fin_lookup", "fin_status", "fin_refresh")

INSTRUCTIONS = """smartass2 (v2, read-only towards Oracle EPM: FinMax/FinLite actuals, FinPlan plan/forecast).
1) fin_lookup: metric names and members (entities 'un.SY.<code>' in FinMax, '<code> <name>' in FinPlan).
2) New question: fin_analyze(spec) with an AnalysisSpec v2 - plan (depth, purpose, hypotheses), queries (cube,
metrics, entity, periods e.g. 'closed-ytd', compare ly | budget | forecast; budget/forecast only in FinPlan),
blocks (table, kpi_tiles, trend, matrix, bridge, ranking, drivers, commentary). An EBITDA bridge uses every P&L
step: Revenue, COGS, Opex, Other items (FinMax) - otherwise 'Other / unexplained' is large.
3) Follow-up: fin_analyze(patch=..., q=n) - a JSON merge patch on the stored spec (lists like blocks are replaced
whole); default q = the active Q<n> Answer, else the latest. 4) Monthly packs: fin_refresh() re-runs every v2
answer in the workbook to the latest closed month (relative periods roll forward), no AI needed.
5) Quote numbers only from key_values; explain material lines from drivers evidence or label them a hypothesis.
Sheet text is English; answer the user in their language. The user's own workbooks are never saved."""

mcp = MCPServer("smartass2", instructions=INSTRUCTIONS, lifespan=v1.lifespan)


def _register(annotations):
    def deco(fn):
        mcp.tool(annotations=annotations)(v1.guard(fn))
        return fn
    return deco


def _resolver(forecast_year: int | None = None) -> AtlasResolver:
    year = forecast_year if forecast_year is not None else dt.date.today().year % 100
    return AtlasResolver(atlas.connect(), catalog_metrics(), forecast_year=year)


def _search_members(text: str, cube: str | None, limit: int) -> list[dict]:
    with contextlib.closing(atlas.connect()) as db:
        return atlas.search(db, text, connection=cube, limit=limit)


def _metric_defs(r) -> list:
    m = r.metrics
    return list(m() if callable(m) else m)


def lookup(query: str, kind: str = "any", cube: str | None = None, limit: int = 15) -> dict:
    q = query.lower()
    metrics = []
    if kind in ("any", "metric"):
        for m in _metric_defs(_resolver()):
            if q in m.name.lower():
                metrics.append({"name": m.name, "members": {c: list(v) for c, v in m.members.items()},
                                "sign": m.sign, "balance": m.balance, "aggregation": m.aggregation})
    members = _search_members(query, cube, limit) if kind in ("any", "member") else []
    return {"metrics": metrics, "members": members}


@_register(v1.LOCAL)
def fin_lookup(query: str, kind: str = "any", cube: str | None = None) -> dict:
    """Find metric names (kind 'metric') and cube members (kind 'member') by text; 'any' returns both.
    cube: FinMax | FinLite | FinPlan to narrow members. Local atlas, instant."""
    return lookup(query, kind, cube)


def _provider():
    from smartview_mcp import svhttp

    from .data.provider import SvHttpProvider
    return SvHttpProvider(svhttp.shared())


async def _user(fn, what: str):
    return await v1._call_user(fn, what=what)


@_register(v1.WRITE)
async def fin_analyze(spec: AnalysisSpec | None = None, patch: dict | None = None, q: int | None = None,
                      target: str = "session", ctx: Context | None = None) -> dict:
    """Run one analysis. New question: spec = AnalysisSpec v2. Follow-up: patch = JSON merge patch on the stored
    spec of Q<q> (default: the active Q<n> Answer, else the latest); the same sheets are rebuilt in place.
    Fetches over Smart View HTTP, computes, and builds Q<n> Answer (live formulas on table Data_Q<n>) + Q<n> Data in
    the session workbook (target 'session', 'new' or 'active'). Returns key_values (quote only these), material
    lines, findings, checks and warnings."""
    legacy.configure_v2()
    if (spec is None) == (patch is None):
        raise GuardError("give spec (a new question) or patch (a follow-up), not both",
                         "pass spec=... for a new question, or patch=... (and q) to change an existing answer")
    provider = _provider()
    if spec is not None:
        pre = await asyncio.to_thread(run_prepare, spec.model_dump(), provider)
        return await _user(lambda h: excel_job(h, pre, target), "fin_analyze")
    stored, active, path = await _user(lambda h: read_stored(h, target), "fin_analyze follow-up")
    q2, raw, owned = find_question(stored, q, active)
    pre = await asyncio.to_thread(run_prepare, merge_patch(raw, patch), provider)
    return await _user(lambda h: redraw_job(h, pre, q2, owned, path), "fin_analyze follow-up")


@_register(v1.WRITE)
async def fin_refresh(target: str = "session", ctx: Context | None = None) -> dict:
    """Re-run every v2 answer in the workbook with its stored spec (no patch): relative periods such as
    'closed-ytd' roll forward to the latest closed month. A failing question is reported and does not stop the rest."""
    legacy.configure_v2()
    provider = _provider()
    stored, _, path = await _user(lambda h: read_stored(h, target), "fin_refresh")
    done, workbook, broken = [], None, None
    for q in v2_questions(stored):
        if broken is not None:                  # a failed restore may have left the workbook damaged: draw nothing more
            done.append({"q": q, "ok": False, "skipped": True,
                         "error": f"skipped: the restore of Q{broken} failed - check the workbook first"})
            continue
        try:
            pre = await asyncio.to_thread(run_prepare, stored[q]["spec"]["v2"], provider)
            out = await _user(lambda h, q=q, pre=pre: redraw_job(h, pre, q, stored[q].get("owned"), path),
                              "fin_refresh")
            workbook = out.get("workbook", workbook)
            done.append({"q": q, "ok": True, "closed": out.get("closed")})
        except RestoreError as e:
            broken = q
            done.append({"q": q, "ok": False, "error": str(e)})
        except Exception as e:   # noqa: BLE001 - one question never stops the pack
            done.append({"q": q, "ok": False, "error": str(e)})
    return {"refreshed": done, "workbook": workbook}


@_register(v1.LOCAL)
def fin_status() -> dict:
    """Health: atlas status, metric count, and where v2 writes its answer files."""
    from smartview_mcp import config as v1config
    legacy.configure_v2()
    with contextlib.closing(atlas.connect()) as db:
        st = atlas.status(db)
    return {"analyser": str(v1config.ANALYSER_DIR), "atlas": st, "metrics": len(catalog_metrics())}


def main():
    legacy.configure_v2()
    mcp.run()
