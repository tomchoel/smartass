"""Read-only MCP server for Oracle Smart View (FinMax / FinLite / FinPlan on Telenor's EPM Cloud).

Tools never write to EPM: submit, calculate and business-rule functions are not exposed at all.
"""
from __future__ import annotations

import datetime as dt
import logging
import re
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Literal

from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import BaseModel, Field

from . import config, logsetup, userexcel
from .host import ExcelHost, HostError

logsetup.configure(config.SERVER_LOG)
logsetup.add_secret(config.HTTP_SECRET)
log = logging.getLogger("smartview")

Connection = Literal["FinMax", "FinLite", "FinPlan"]
READ = ToolAnnotations(read_only_hint=True, open_world_hint=True)          # reads EPM, writes nothing
LOCAL = ToolAnnotations(read_only_hint=True, open_world_hint=False)        # local atlas / status only
# Writes to the user's workbook or to files (never to EPM): not read-only, so clients do not auto-approve them.
WRITE = ToolAnnotations(read_only_hint=False, destructive_hint=False, idempotent_hint=False, open_world_hint=True)
ADMIN = ToolAnnotations(read_only_hint=False, destructive_hint=False, open_world_hint=False)

_host: ExcelHost | None = None
_STARTED = dt.datetime.now().isoformat(timespec="seconds")   # code on disk at this time is what runs


def host() -> ExcelHost:
    global _host
    if _host is None:
        _host = ExcelHost(visible=False)
    return _host


@asynccontextmanager
async def lifespan(_server):
    try:
        yield {}
    finally:
        try:
            if _host is not None:
                _host.shutdown()
        finally:
            from . import tools_drill
            if tools_drill._browser is not None:
                tools_drill._browser.shutdown()


INSTRUCTIONS = """Read-only access to Telenor Norge's Oracle EPM Cloud cubes (Smart View); answers are built as
analyses in the user's Excel. Same rules for every agent (Claude Code, Copilot, ...):
- Cubes: FinMax = actuals incl. FTE (ac.STA20Y); FinLite = lighter actuals; FinPlan = budget/forecast. Actual vs
  budget/forecast is done entirely in FinPlan (Actual = version Working; Target and the latest 'Forecast Q<q><yy>' =
  version Final) with a reconciliation line to FinMax EBITDA. FinPlan KPIs (volumes, ARPU, %) are read Periodic per
  month - Accumulated sums the months for many of them; actual volumes are not in FinPlan for BANORWAY.
- A 'Smart View rejected these members' warning means those lines were left out (the rest was retrieved): say so.
- Members: entities 'un.SY.<code>' (un.SY.2WHS Wholesale, 2BUS Business, 2SSM, 2EI, 2TEC ...), accounts 'ac.<code>'
  (ac.C30 revenue, ac.C40S COGS, ac.C70 opex, ac.EBITDA), periods 'Sep 26', balance 'Periodic' | 'Accumulated' (=
  YTD), scenario 'Actual', currency 'NOK'. Resolve names with sv_find_member / sv_members (local atlas, instant).
  P&L chain: C30 + C40S = C48; C50S + C60 = C70; C48 + C70 = EBITDA_BOI; + C73 = EBITDA.
- Use the last CLOSED month (the next one is open, with reversals). History is restated to today's organisation
  (the 1 June 2026 reorg emptied 2CON Consumer into 2SSM + 2EI).
- Product / profitability / business-area questions: call sv_scope(question) FIRST (business map + concept library).
  If a concept matches (e.g. 'splitt', or a product family 'fam_*'), answer with sv_concept(concept) - it marks every line booked / estimated (with
  key) / attributed / not in FinMax. No concept: research, PROPOSE the blocks to the user, then save the concept.
- Analyses: ONE call - sv_ws_build(question, pulls, spec) (or sv_analyse for a template). Default target 'session' =
  this session's own workbook (the analyser folder\\<project>\\..., Q1, Q2 ...); 'active' only when the user asks
  for their open workbook. Follow-ups: sv_ws_answer(q, spec_patch=..., target=<workbook>). Add the same period last
  year by default. Write key_points only for interpretation - factual points are drafted from the numbers.
- Format: full NOK 1 000 000, negatives (1 000 000) in tables and with a minus sign in text; costs negative;
  variance = current - comparison (positive = better). Line labels in English as FinMax shows them.
- Quick single numbers: sv_query. The Produktregnskapet archive folders are never written to.
- Invoice / journal-line detail behind any FinLite number: sv_drill(question, cell) (hidden headless browser,
  reconciled to the cube), then sv_ws_answer with {drill: 'd', where: {...}} cells. Never Smart View's own
  drill-through execute: it opens the user's browser.
- Any workbook the user has OPEN (no Smart View needed): xl_* tools - xl_workbooks / xl_overview / xl_read to look,
  xl_write / xl_rows / xl_add_sheet / xl_format / xl_format_sheet to change (each snapshotted; xl_revert undoes),
  xl_trace / xl_check to explain and audit formulas. Never saved by the tools: the user saves. Closed files: the
  separate Excel MCP. Formulas follow skill tn-formelstil (SUM over ranges, XLOOKUP, no guessed ranges)."""

mcp = MCPServer("smartview", instructions=INSTRUCTIONS, lifespan=lifespan)


class AxisIn(BaseModel):
    dimension: str = Field(description="Dimension name, e.g. 'All Accounts', 'All Entities', 'All Time Periods'")
    members: list[str] = Field(description="Member names in order, e.g. ['ac.C30','ac.EBITDA'] or ['Aug 26','Sep 26']")


def _err(e: Exception) -> ToolError:
    return ToolError(str(e))


NEXT_STEP = {
    "HostError": "Excel/Smart View did not respond: call sv_restart, then retry.",
    "DrillError": "The hidden drill browser failed: retry once.",
    "CellTooWide": "Narrow the cell (e.g. child entities) and drill again.",
    "TimeoutError": "Retry with a smaller request (fewer members or periods).",
}
GENERIC_NEXT = ("If it repeats, call sv_status (in a terminal: `smartass doctor`) and send the file from "
                "`smartass doctor --bundle` to the maintainer.")


def with_next_step(msg: str) -> str:
    return msg if "next step" in msg.lower() else f"{msg.rstrip(' .')}. Next step: {GENERIC_NEXT}"


def explain(tool_name: str, e: Exception) -> str:
    base = f"{tool_name} failed - {type(e).__name__}: {logsetup.redact(e)}"
    # Most specific advice along the class hierarchy (HostTimeout -> HostError, CellTooWide before DrillError).
    hint = next((NEXT_STEP[c.__name__] for c in type(e).__mro__ if c.__name__ in NEXT_STEP), None)
    return f"{base.rstrip(' .')}. Next step: {hint} {GENERIC_NEXT}" if hint else with_next_step(base)


def guard(fn):
    """ANY exception reaches the model as readable text with a next step (the SDK otherwise reports only
    'Error executing tool X' for non-ToolError exceptions)."""
    import functools
    import inspect

    import anyio.to_thread

    is_async = inspect.iscoroutinefunction(fn)

    @functools.wraps(fn)          # keeps __wrapped__, so the SDK builds the same input schema
    async def guarded(*args, **kwargs):
        try:
            if is_async:
                return await fn(*args, **kwargs)
            # The SDK runs sync tools in a worker thread; keep that.
            return await anyio.to_thread.run_sync(functools.partial(fn, *args, **kwargs))
        except ToolError as e:
            raise ToolError(with_next_step(logsetup.redact(e))) from None
        except Exception as e:   # noqa: BLE001
            log.exception("tool %s failed", fn.__name__)
            raise ToolError(explain(fn.__name__, e)) from None
    return guarded


def tool(annotations: ToolAnnotations = READ):
    def deco(fn):
        return mcp.tool(annotations=annotations)(guard(fn))
    return deco


async def _call(fn, timeout: float | None = None, what: str = "call"):
    """host().call with every failure surfaced as a readable ToolError."""
    try:
        return await host().call(fn, timeout=timeout, what=what)
    except (HostError, ToolError) as e:
        raise ToolError(str(e)) from None


def _safe_name(s: str) -> str:
    return re.sub(r"[^\w\-. ]+", "_", s).strip()[:80] or "query"


def _doc(extra: str):
    """Append shared reference text to a tool's docstring (applied before @tool reads it)."""
    def deco(fn):
        fn.__doc__ = (fn.__doc__ or "").rstrip() + "\n    " + extra
        return fn
    return deco


def _is_our_analysis_file(path: str) -> bool:
    """A file the MCP itself created in the analyser folder (config.ANALYSER_DIR; never a user's or an archive workbook)."""
    try:
        parent, root = Path(path).resolve().parent, config.ANALYSER_DIR.resolve()
        return (parent == root or parent.parent == root) and not userexcel.is_archive(path)   # analyser\[<project>\]
    except (OSError, ValueError):
        return False


async def _call_user(fn, what: str):
    """host().call for jobs that drive the user's Excel (never killed on timeout)."""
    try:
        return await host().call(fn, what=what, user_excel=True)
    except HostError as e:
        raise ToolError(str(e)) from None


# Tool modules register their tools on `mcp` when imported (module imports only - see the rule above).
# Order = order the model sees the tools in (status, query, ws, drill, xl - as before the split).
from . import tools_status  # noqa: E402,F401,I001
from . import tools_query  # noqa: E402,F401
from . import tools_ws  # noqa: E402,F401
from . import tools_drill  # noqa: E402,F401
from . import tools_xl  # noqa: E402,F401


def main():
    mcp.run()


if __name__ == "__main__":
    main()
