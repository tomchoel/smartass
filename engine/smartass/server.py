# engine/smartass/server.py
"""smartass: the MCP server (stdio; started by run_server.py through child_server.py). Tools: fin_analyze (new
question or follow-up patch), fin_lookup, fin_status, fin_refresh, fin_drill, and xl_snapshot / xl_revert / xl_tidy
for edits to open workbooks made with Excel MCP. Error text (guard), Excel jobs (_call_user, progress notes) and the
lifespan (warm-up, closing the hidden Excel) come from core.serverkit."""
from __future__ import annotations

import asyncio
import contextlib
import datetime as dt

from mcp.server.mcpserver import Context, MCPServer

from . import __version__
from .knowledge import atlas
from .core import serverkit as kit
from .xl.host import InputError

from .data.adhoc import adhoc_fetch
from .data.provider import ComboProvider
from .followup import find_question, merge_patch, v2_questions
from .model.lines import LineSet
from .model.spec import AnalysisSpec, LinesBlock
from .model.validate import GuardError
from .pipeline import RestoreError, excel_job, read_stored, redraw_job, run_prepare

TOOL_NAMES = ("fin_analyze", "fin_lookup", "fin_status", "fin_refresh", "fin_drill", "xl_snapshot", "xl_revert",
              "xl_tidy")

INSTRUCTIONS = """smartass (read-only towards Oracle EPM: FinMax/FinLite actuals, FinPlan plan/forecast).
1) fin_lookup: metric names and members (entities 'un.SY.<code>' in FinMax, '<code> <name>' in FinPlan).
2) New question: fin_analyze(spec) with an AnalysisSpec v2 - plan (depth, purpose, hypotheses), queries (cube,
metrics, entity, periods e.g. 'closed-ytd', compare ly | budget | forecast; budget/forecast only in FinPlan),
blocks (table, kpi_tiles, trend, matrix, bridge, ranking, drivers, commentary). An EBITDA bridge uses every P&L
step: Revenue, COGS, Opex, Other items (FinMax) - otherwise 'Other / unexplained' is large.
3) Follow-up: fin_analyze(patch=..., q=n) - a JSON merge patch on the stored spec (lists like blocks are replaced
whole); default q = the active Q<n> Answer, else the latest. 4) Monthly packs: fin_refresh() re-runs every v2
answer in the workbook to the latest closed month (relative periods roll forward), no AI needed.
5) Quote numbers only from key_values; explain material lines from drivers evidence or label them a hypothesis.
6) Editing a workbook the user has open with Excel MCP: xl_snapshot first, then edit, xl_revert to undo; finish
with xl_tidy (zero BLOCK findings). 7) After the server reloads, Claude Code may keep an old tool schema: if a
parameter seems missing, ask the user to run /mcp and reconnect smartass. Repeated errors: fin_status(restart=True).
Ledger lines: a lines block (cell with All Accounts, All Entities, All Time Periods; Periodic; FinLite) or
fin_drill(cell) ranks suppliers, projects or accounts and ties them to the cube.
Sheet text is English; answer the user in their language. The user's own workbooks are never saved."""

mcp = MCPServer("smartass", instructions=INSTRUCTIONS, lifespan=kit.lifespan)


def _register(annotations):
    def deco(fn):
        mcp.tool(annotations=annotations)(kit.guard(fn))
        return fn
    return deco


def _resolver(forecast_year: int | None = None):
    """A KnowledgeResolver (knowledge.db, built when missing, + the atlas); the caller closes it."""
    from .knowledge.resolver import open_resolver
    year = forecast_year if forecast_year is not None else dt.date.today().year % 100
    return open_resolver(year)


def lookup(query: str, kind: str = "any", cube: str | None = None, limit: int = 15) -> dict:
    """{metrics: [id, name, aliases, description, differs, members, caveats ...], members: [... with caveats]}."""
    with contextlib.closing(_resolver()) as r:
        return r.search(query, kind, cube, limit)


@_register(kit.LOCAL)
def fin_lookup(query: str, kind: str = "any", cube: str | None = None) -> dict:
    """Find metrics (kind 'metric': by name or alias, e.g. 'driftskostnader'; with description, members and how the
    cubes differ) and cube members (kind 'member'); 'any' returns both. Every hit carries its caveats (reorgs,
    abolished units, account traps): pass them on. cube: FinMax | FinLite | FinPlan to narrow. Local, instant."""
    return lookup(query, kind, cube)


def _provider():
    from .epm import svhttp

    from .data.provider import SvHttpProvider
    return SvHttpProvider(svhttp.shared())


def _clean(fn):
    """Refusals raised inside a host job are input errors, not 'Unexpected ... run fin_status(restart=True)'."""
    def job(h):
        try:
            return fn(h)
        except GuardError as e:
            raise InputError(f"{e.reason} - {e.fix}") from None
    return job


async def _user(fn, what: str):
    return await kit._call_user(_clean(fn), what=what)


async def _drill(cell: dict):
    """The FinLite ledger lines behind one cube cell (v1 drill: Smart View link, hidden browser, reconciliation)."""
    from .epm import drillstep
    return await drillstep.drill(cell, connection="FinLite")


async def _line_sets(spec: AnalysisSpec) -> tuple[dict, list[str]]:
    """Drill every lines block: {'b<i>': LineSet} and the drills' warnings. Project grouping marks each line's
    booked project (INITIATIVE) first. A drill with 0 lines still gives a LineSet (prepare refuses it, with the reason)."""
    from .model.lines import add_initiative
    blocks = [(i, b) for i, b in enumerate(spec.blocks) if isinstance(b, LinesBlock)]
    results = await asyncio.gather(*[_drill(b.cell) for _, b in blocks])      # independent cells: drilled concurrently
    out: dict[str, LineSet] = {}
    notes: list[str] = []
    for (i, b), res in zip(blocks, results):
        rows, cols = res.rows, list(res.columns)
        if b.group_by == "project":
            cols = add_initiative(rows, cols)
        out[f"b{i}"] = LineSet(f"b{i}", rows, cols, res.cell_value, reconciliation=res.reconciliation,
                                warnings=list(res.warnings), cell=b.cell)
        notes += res.warnings
    return out, notes


def _has_lines(raw: dict) -> bool:
    return any(isinstance(b, dict) and b.get("type") == "lines" for b in raw.get("blocks") or [])


async def _prepare(raw: dict, provider, spec: AnalysisSpec | None = None, target: str = "session"):
    """(Prepared, drill warnings). A stored or patched raw spec with lines blocks is re-drilled, so lines are fresh.
    Grids the HTTP route could not fetch are pulled as ad hoc grids in the user's Excel (one job), then the model is
    prepared again with those results fixed; this fallback runs at most once."""
    notes: list[str] = []
    lines: dict = {}
    if spec is not None or _has_lines(raw):
        lines, notes = await _line_sets(spec or AnalysisSpec.model_validate(raw))
    pre = await asyncio.to_thread(run_prepare, raw, provider, lines=lines)
    failed = getattr(pre, "failed", None)
    if failed:
        fixed = await _user(lambda h: adhoc_fetch(h, failed, target), "ad hoc fallback")
        pre = await asyncio.to_thread(run_prepare, raw, ComboProvider(provider, {**pre.fetched, **fixed}), lines=lines)
    return pre, notes


TEXT_FIX_PREFIX = "Text broken in EPM"


def _broken_texts(warning: str) -> list[str]:
    """'Text broken in EPM (letters lost at load): a; b - add corrections to ...' -> ['a', 'b']."""
    body = warning.split("): ", 1)[-1].split(" - add corrections", 1)[0]
    return [t.strip() for t in body.split(";") if t.strip()]


def _with_notes(out: dict, notes: list[str]) -> dict:
    """Drill notes join the warnings; the broken-text ones become reply['text_fixes'] (candidates for text_fixes.json)."""
    fixes = [t for n in notes if n.startswith(TEXT_FIX_PREFIX) for t in _broken_texts(n)]
    notes = [n for n in notes if not n.startswith(TEXT_FIX_PREFIX)]
    if notes:
        out["warnings"] = list(out.get("warnings", [])) + notes
    if fixes:
        out["text_fixes"] = list(dict.fromkeys(out.get("text_fixes", []) + fixes))
    return out


async def _new_question(spec: AnalysisSpec, target: str) -> dict:
    pre, notes = await _prepare(spec.model_dump(), _provider(), spec, target)
    return _with_notes(await _user(lambda h: excel_job(h, pre, target), "fin_analyze"), notes)


@_register(kit.WRITE)
async def fin_analyze(spec: AnalysisSpec | None = None, patch: dict | None = None, q: int | None = None,
                      target: str = "session", ctx: Context | None = None) -> dict:
    """Run one analysis. New question: spec = AnalysisSpec v2. Follow-up: patch = JSON merge patch on the stored
    spec of Q<q> (default: the active Q<n> Answer, else the latest); the same sheets are rebuilt in place.
    Fetches over Smart View HTTP, computes, and builds Q<n> Answer (live formulas on table Data_Q<n>) + Q<n> Data in
    the session workbook (target 'session', 'new' or 'active'). Returns key_values (quote only these), material
    lines, findings, checks and warnings."""
    if (spec is None) == (patch is None):
        raise GuardError("give spec (a new question) or patch (a follow-up), not both",
                         "pass spec=... for a new question, or patch=... (and q) to change an existing answer")
    if spec is not None:
        return await _new_question(spec, target)
    provider = _provider()
    stored, active, path = await _user(lambda h: read_stored(h, target), "fin_analyze follow-up")
    q2, raw, owned = find_question(stored, q, active)
    pre, notes = await _prepare(merge_patch(raw, patch), provider, target=path)
    return _with_notes(await _user(lambda h: redraw_job(h, pre, q2, owned, path), "fin_analyze follow-up"), notes)


@_register(kit.WRITE)
async def fin_drill(cell: dict, group_by: str = "supplier", top: int = 10, target: str = "session",
                    ctx: Context | None = None) -> dict:
    """Ledger lines behind one cube cell (FinLite, Periodic): ranked by supplier, project or account, with Other,
    lines without the group, the lines total, the cube value and the gap. cell: {"All Accounts": "ac.C60",
    "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"} (a parent period such as 'Yr 26' works).
    Same sheets and checks as a lines block in fin_analyze."""
    spec = AnalysisSpec.model_validate({
        "version": 2, "question": f"Ledger lines behind {' / '.join(cell.values())}", "plan": {"depth": "quick"},
        "queries": [], "blocks": [{"type": "lines", "cell": cell, "group_by": group_by, "top": top}]})
    return await _new_question(spec, target)


@_register(kit.WRITE)
async def fin_refresh(target: str = "session", ctx: Context | None = None) -> dict:
    """Re-run every v2 answer in the workbook with its stored spec (no patch): relative periods such as
    'closed-ytd' roll forward to the latest closed month. A failing question is reported and does not stop the rest."""
    provider = _provider()
    stored, _, path = await _user(lambda h: read_stored(h, target), "fin_refresh")
    done, workbook, broken = [], None, None
    for q in v2_questions(stored):
        if broken is not None:                  # a failed restore may have left the workbook damaged: draw nothing more
            done.append({"q": q, "ok": False, "skipped": True,
                         "error": f"skipped: the restore of Q{broken} failed - check the workbook first"})
            continue
        try:
            pre, notes = await _prepare(stored[q]["spec"]["v2"], provider, target=path)
            out = _with_notes(await _user(lambda h, q=q, pre=pre: redraw_job(h, pre, q, stored[q].get("owned"), path),
                                          "fin_refresh"), notes)
            workbook = out.get("workbook", workbook)
            done.append({"q": q, "ok": True, "closed": out.get("closed")})
        except RestoreError as e:
            broken = q
            done.append({"q": q, "ok": False, "error": str(e)})
        except Exception as e:   # noqa: BLE001 - one question never stops the pack
            done.append({"q": q, "ok": False, "error": str(e)})
    return {"refreshed": done, "workbook": workbook}


@_register(kit.ADMIN)
async def fin_status(restart: bool = False) -> dict:
    """Health: version, atlas age, knowledge build (metrics, facts, drafts awaiting approval), where answer files go,
    and the doctor checks (config, Excel, Smart View, VBA access, EPM, drill browser, atlas, knowledge, output
    folder, HTTP API, update; only problems listed).
    restart=True after repeated errors or to load new code: the relay starts a fresh server process (its hidden
    Excel closes); run without the relay, only this server's own Excel is closed (it restarts on the next pull).
    A stale atlas: run `smartass atlas build` in a terminal."""
    import os

    from .core import config
    out: dict = {"version": __version__}
    if restart:
        if os.environ.get("SVKIT_PROXY"):
            out["restarted"] = "relay"             # run_server.py replaced the child before forwarding this call
        else:
            old, kit._host = kit._host, None
            if old is not None:
                await asyncio.to_thread(old.shutdown)
            out["restarted"] = "excel"
    with contextlib.closing(atlas.connect()) as db:
        st = atlas.status(db)
    from .knowledge.build import summary
    kn = summary(config.KNOWLEDGE_DB)
    from .ops import doctor
    checks = await asyncio.to_thread(doctor.run)         # network checks (EPM 5 s, HTTP 2 s) off the event loop
    return {**out, "behind_relay": bool(os.environ.get("SVKIT_PROXY")), "analyser": str(config.ANALYSER_DIR),
            "atlas": {"oldest_age_days": st["oldest_age_days"], "stale": st["stale"],
                      "dimensions_built": len(st["built"])},
            "metrics": kn.get("metrics", 0),
            "knowledge": {"built_at": kn.get("built_at"), "metrics": kn.get("metrics", 0),
                          "facts": kn.get("facts", 0), "drafts": kn.get("drafts", 0)},
            "doctor": {"overall": doctor.overall(checks),
                       "problems": [c.__dict__ for c in checks if c.status != "OK"]}}


@_register(kit.WRITE)
async def xl_snapshot(workbook: str | None = None, sheet: str | None = None, range: str | None = None) -> dict:
    """Save the current state of a range (default: the sheet's used range) of an open workbook, changing nothing.
    Call before any Excel MCP edit to a workbook the user has open; xl_revert(change_id) undoes the edit."""
    from .workbench import snapshot
    return await _user(lambda h: snapshot(workbook, sheet, range), "xl_snapshot")


@_register(kit.WRITE)
async def xl_revert(change_id: str | None = None, workbook: str | None = None, force: bool = False) -> dict:
    """Put a range back as xl_snapshot (or a v1 xl_* edit) recorded it. Default: the latest change. Not saved."""
    from .workbench import revert
    return await _user(lambda h: revert(change_id, workbook, force), "xl_revert")


@_register(kit.WRITE)
async def xl_tidy(workbook: str | None = None, sheet: str | None = None, rows: dict | None = None,
                  last_col: str | None = None) -> dict:
    """TN report style (when rows + last_col are given) and the tn-formelstil formula check; zero BLOCK findings
    before work is done. rows: {"1": "title", "4": "header", "5-12": "line", "13": "sum", ...}; last_col e.g. "F"."""
    from .workbench import tidy
    return await _user(lambda h: tidy(workbook, sheet, rows, last_col), "xl_tidy")


def main():
    mcp.run()
