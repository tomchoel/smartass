# engine/smartass/server.py
"""smartass2: the v2 MCP server (built beside v1 until v2 passes the test set). Three tools in Plan 1B:
fin_analyze, fin_lookup, fin_status. Error text (guard), Excel jobs (_call_user, progress notes) and the lifespan
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
from .model.spec import AnalysisSpec

TOOL_NAMES = ("fin_analyze", "fin_lookup", "fin_status")

INSTRUCTIONS = """smartass2 (v2, read-only towards Oracle EPM: FinMax/FinLite actuals, FinPlan plan/forecast).
Loop: 1) fin_lookup to find metric names and members (entities 'un.SY.<code>' in FinMax, '<code> <name>' in FinPlan).
2) Write an AnalysisSpec (version 2): plan (depth quick|standard|deep, purpose, hypotheses), queries (cube, metrics,
entity, periods like 'closed-ytd', compare ly | budget | forecast - budget/forecast only in FinPlan), blocks (table,
kpi_tiles, trend, matrix, bridge, ranking, drivers, commentary). 3) fin_analyze(spec): builds Q<n> Answer (live
formulas) + Q<n> Data in the session workbook; a refused spec says why and how to fix it. 4) Quote numbers only from
key_values; explain material lines from drivers evidence or label them as a hypothesis. Sheet text is English;
answer the user in their language. The user's own workbooks are never saved."""

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


@_register(v1.WRITE)
async def fin_analyze(spec: AnalysisSpec, target: str = "session", ctx: Context | None = None) -> dict:
    """Run one analysis: validate the spec, fetch the data over Smart View HTTP, compute, and build Q<n> Answer
    (live formulas on table Data_Q<n>) + Q<n> Data in the session workbook (target 'session', 'new' or 'active').
    Returns key_values (quote only these), material lines, findings, checks and warnings."""
    from smartview_mcp import svhttp

    from .data.provider import SvHttpProvider
    from .pipeline import excel_job, prepare
    legacy.configure_v2()
    provider = SvHttpProvider(svhttp.shared())
    closed = await asyncio.to_thread(provider.closed_months)
    resolver = _resolver(closed["FinPlan"][0] % 100)
    pre = await asyncio.to_thread(prepare, spec.model_dump(), resolver, provider)
    return await v1._call_user(lambda h: excel_job(h, pre, target), what="fin_analyze")


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
