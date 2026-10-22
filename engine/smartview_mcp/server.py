"""Read-only MCP server for Oracle Smart View (FinMax / FinLite / FinPlan on Telenor's EPM Cloud).

Tools never write to EPM: submit, calculate and business-rule functions are not exposed at all.
"""
from __future__ import annotations

import datetime as dt
import logging
import os
import re
import time
from contextlib import asynccontextmanager, contextmanager
from pathlib import Path
from typing import Literal

from mcp.server import MCPServer
from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import BaseModel, Field

from . import atlas, config, logsetup, ops, qlog
from .grid import Axis, to_markdown
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
            if _browser is not None:
                _browser.shutdown()


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


def tool(annotations: ToolAnnotations = READ):
    """mcp.tool + a guard: ANY exception reaches the model as readable text (the SDK otherwise reports only
    'Error executing tool X' for non-ToolError exceptions)."""
    import functools

    def deco(fn):
        @functools.wraps(fn)
        async def guarded(*args, **kwargs):
            try:
                return await fn(*args, **kwargs)
            except ToolError:
                raise
            except Exception as e:   # noqa: BLE001
                log.exception("tool %s failed", fn.__name__)
                raise ToolError(f"{type(e).__name__}: {e}") from None
        return mcp.tool(annotations=annotations)(guarded)
    return deco


async def _call(fn, timeout: float | None = None, what: str = "call"):
    """host().call with every failure surfaced as a readable ToolError."""
    try:
        return await host().call(fn, timeout=timeout, what=what)
    except (HostError, ToolError) as e:
        raise ToolError(str(e)) from None


def _safe_name(s: str) -> str:
    return re.sub(r"[^\w\-. ]+", "_", s).strip()[:80] or "query"


# ---------------------------------------------------------------- primary tools
@tool(annotations=LOCAL)
async def sv_status() -> dict:
    """Health of the Smart View engine: Excel pid, connected sheets, Smart View version, restarts, atlas age,
    and which drill browsers can start (drill_browser). Cheap; does not start Excel or a browser."""
    db = atlas.connect()
    st = atlas.status(db)
    from . import sso
    h = _host.status() if _host else {"excel_running": False, "smartview_version": sso.smartview_version()}
    from .drill_browser import browser_health
    try:
        drill = browser_health()
    except Exception as e:   # noqa: BLE001 - a health probe never fails the status call
        drill = {"error": str(e)[:200]}
    return {**h, "server_started": _STARTED, "behind_reload_proxy": bool(os.environ.get("SVKIT_PROXY")),
            "drill_browser": drill,
            "atlas": {"oldest_age_days": st["oldest_age_days"], "stale": st["stale"],
                      "dimensions_built": len(st["built"])},
            "query_log": str(config.QUERY_LOG), "output_dir": str(config.OUT_DIR)}


@tool(annotations=READ)
async def sv_dimensions(connection: Connection) -> dict:
    """List a cube's dimensions (names = top members) and their default axis. Use before sv_query.
    First call starts Excel and signs in via SSO (~15 s)."""
    try:
        return await _call(lambda h: ops.dimensions(h, connection), what="sv_dimensions")
    except (HostError, ToolError) as e:
        raise _err(e)


@tool(annotations=WRITE)
async def sv_query(connection: Connection, rows: list[AxisIn], columns: list[AxisIn], ctx: Context,
                   pov: dict[str, str] | None = None, suppress_missing: bool = True, suppress_zeros: bool = False,
                   save_xlsx: bool = False, title: str | None = None, large: bool = False,
                   max_preview_rows: int = config.PREVIEW_ROWS) -> dict:
    """Retrieve a data grid from a cube and return it as a table.

    rows/columns: one entry per dimension, each with explicit member names (cross product is built).
    pov: {dimension: member} for dimensions not on rows/columns; unspecified ones use the top member
    (e.g. {'All Scenarios':'Actual','All Currencies':'NOK','All Balance Types':'Accumulated'}).
    Example: rows=[{dimension:'All Accounts',members:['ac.C30','ac.EBITDA']}],
             columns=[{dimension:'All Time Periods',members:['Aug 26','Sep 26']}],
             pov={'All Entities':'un.SY.2WHS','All Scenarios':'Actual','All Currencies':'NOK',
                  'All Balance Types':'Accumulated'}.
    Unknown members are reported (Smart View silently ignores them) - fix with sv_find_member.
    save_xlsx writes the full result (with a POV/metadata sheet, label 'Internal') to the output folder.
    Grids over 50,000 cells need large=true."""
    rows_ax = [Axis(a.dimension, a.members) for a in rows]
    cols_ax = [Axis(a.dimension, a.members) for a in columns]
    n_cells = 1
    for a in rows_ax + cols_ax:
        n_cells *= max(1, len(a.members))
    if n_cells > config.MAX_CELLS and not large:
        raise ToolError(f"Grid has {n_cells:,} cells (> {config.MAX_CELLS:,}). Narrow the members, split the query, "
                        "or pass large=true if you really need it.")
    pov = pov or {}
    t0 = time.monotonic()
    await ctx.report_progress(1, 3, "retrieving")
    try:
        lay, parsed, secs = await _call(
            lambda h: ops.query(h, connection, pov, rows_ax, cols_ax, suppress_missing, suppress_zeros),
            timeout=max(config.CALL_TIMEOUT_S, n_cells / 200), what="sv_query")
    except (HostError, ToolError) as e:
        qlog.write({"tool": "sv_query", "connection": connection, "pov": pov, "rows": [r.model_dump() for r in rows],
                    "columns": [c.model_dump() for c in columns], "error": str(e)})
        raise _err(e)
    await ctx.report_progress(2, 3, "parsed")
    saved = None
    if save_xlsx:
        stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        path = config.OUT_DIR / f"{stamp}_{connection}_{_safe_name(title or 'query')}.xlsx"
        meta = [["Source", f"Oracle Smart View / EPM Cloud, connection {connection}"],
                ["Retrieved", dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                ["Title", title or ""]] + [["POV " + d, m] for d, m in lay.pov] + \
               [["Rows", "; ".join(f"{a.dimension}: {', '.join(a.members)}" for a in rows_ax)],
                ["Columns", "; ".join(f"{a.dimension}: {', '.join(a.members)}" for a in cols_ax)],
                ["Unknown members", ", ".join(parsed.unknown_members) or "none"]]
        label_cols = [f"{d} (label)" for d in parsed.columns[:lay.nr]]
        data = [parsed.columns + label_cols] + [r + lab for r, lab in zip(parsed.rows, parsed.row_labels)]
        try:
            saved = str(await _call(lambda h: ops.save_table(h, path, [("Data", data), ("POV", meta)]),
                                          what="save_xlsx"))
        except (HostError, ToolError) as e:
            saved = f"NOT SAVED: {e}"
    await ctx.report_progress(3, 3, "done")
    qlog.write({"tool": "sv_query", "connection": connection, "pov": dict(lay.pov),
                "rows": [r.model_dump() for r in rows], "columns": [c.model_dump() for c in columns],
                "cells": lay.n_cells, "rows_returned": len(parsed.rows), "retrieve_s": round(secs, 2),
                "total_s": round(time.monotonic() - t0, 2), "unknown": parsed.unknown_members, "saved": saved})
    warnings = parsed.pov_problems + ([f"{parsed.no_access_cells} cells are #NoAccess (no read access)"]
                                      if parsed.no_access_cells else [])
    if parsed.unknown_members:
        warnings.insert(0, "Unknown/unrecognised members (no data returned for them): "
                        + "; ".join(parsed.unknown_members) + ". Use sv_find_member to get the right names.")
    return {
        "connection": connection,
        "pov": dict(lay.pov),
        "rows_returned": len(parsed.rows),
        "rows_requested": parsed.total_rows,
        "warnings": warnings,
        "table": to_markdown(parsed.columns, parsed.rows, parsed.row_labels, limit=max_preview_rows),
        "saved_xlsx": saved,
        "retrieve_seconds": round(secs, 2),
    }


@tool(annotations=LOCAL)
async def sv_find_member(text: str, connection: Connection | None = None, dimension: str | None = None,
                         live: bool = False, limit: int = 25) -> dict:
    """Find member names by code or description, e.g. 'WHS', 'Wholesale', 'C30', 'EBITDA', 'Sep 26'.
    Searches the local atlas (fast, all three cubes). live=true (needs connection) asks EPM directly for one exact
    member name and returns its properties (dimension, parent, level, child count, aliases)."""
    out: dict = {}
    db = atlas.connect()
    out["atlas_hits"] = atlas.search(db, text, connection, dimension, limit)
    st = atlas.status(db)
    if st["stale"]:
        out["atlas_note"] = ("Atlas is missing or older than 30 days - run sv_atlas_build."
                             if st["built"] else "Atlas not built yet - run sv_atlas_build (one-off, minutes).")
    if live or (not out["atlas_hits"] and connection):
        if not connection:
            raise ToolError("live lookup needs a connection")
        try:
            out["live"] = await _call(lambda h: ops.find_member(h, connection, text), what="find_member")
        except (HostError, ToolError) as e:
            out["live"] = {"found": False, "error": str(e)}
    return out


@tool(annotations=LOCAL)
async def sv_members(connection: Connection, member: str,
                     relation: Literal["children", "descendants", "leaves", "ancestors"] = "children",
                     limit: int = 500) -> dict:
    """Navigate the hierarchy from the atlas: children / descendants / leaves (level-0 descendants) / ancestors of a
    member name (e.g. 'un.SY.2WHS', 'ac.C70', 'Yr 26'). Use the names in sv_query member lists."""
    db = atlas.connect()
    if relation == "children":
        res = atlas.children(db, connection, member, limit)
    elif relation == "ancestors":
        res = atlas.ancestors(db, connection, member)
    else:
        res = atlas.descendants(db, connection, member, leaves_only=(relation == "leaves"), limit=limit)
    if not res:
        hint = atlas.search(db, member, connection, None, 5)
        return {"member": member, "relation": relation, "result": [],
                "note": "Nothing in the atlas for this member" + (f"; did you mean: {[h['name'] for h in hint]}"
                                                                  if hint else "; build it with sv_atlas_build")}
    return {"member": member, "relation": relation, "count": len(res), "result": res}


@tool(annotations=ADMIN)
async def sv_atlas_build(ctx: Context, connections: list[Connection] | None = None,
                         dimensions: list[str] | None = None) -> dict:
    """(Re)build the local atlas: every dimension tree of the given cubes (default all three) via Smart View
    zoom-all-levels. Read-only against EPM; takes a few minutes. Rebuild when sv_status says it is stale."""
    conns = connections or ["FinMax", "FinLite", "FinPlan"]
    db = atlas.connect()
    done, step = [], 0
    for conn in conns:
        try:
            dims = (await _call(lambda h: ops.dimensions(h, conn), what="dimensions"))["dimensions"]
        except (HostError, ToolError) as e:
            done.append({"connection": conn, "error": str(e)})
            continue
        wanted = {ops.canon_dim(conn, d) for d in dimensions} if dimensions else None
        if wanted is None:
            atlas.purge(db, conn)                 # full rebuild of this cube: drop stale/renamed dimensions
        for dim in dims:
            if wanted and dim not in wanted:
                continue
            step += 1
            await ctx.report_progress(step, None, f"{conn}: {dim}")
            t0 = time.monotonic()
            try:
                members = await _call(lambda h: ops.dump_dimension(h, conn, dim), timeout=600,
                                            what=f"atlas {conn}/{dim}")
                atlas.store(db, conn, dim, members, time.monotonic() - t0)
                done.append({"connection": conn, "dimension": dim, "members": len(members),
                             "seconds": round(time.monotonic() - t0, 1)})
            except (HostError, ToolError) as e:
                atlas.store(db, conn, dim, [], time.monotonic() - t0, str(e))
                done.append({"connection": conn, "dimension": dim, "error": str(e)})
    try:
        from . import atlas_history
        atlas_history.ensure_schema(db)
        snaps = atlas_history.snapshot_all(db)
        done.append({"snapshots": len(snaps), "new": sum(1 for s in snaps if s.get("new"))})
    except Exception as e:   # noqa: BLE001 - history must never break a build
        done.append({"snapshot_error": str(e)})
    qlog.write({"tool": "sv_atlas_build", "result": done})
    return {"built": done, "atlas_db": str(config.ATLAS_DB)}


@tool(annotations=LOCAL)
async def sv_atlas_status() -> dict:
    """What the atlas contains (per cube and dimension), when each part was built, and whether it is stale."""
    return atlas.status(atlas.connect())


@tool(annotations=WRITE)
async def sv_refresh_workbook(path: str, output_name: str | None = None) -> dict:
    """Refresh every Smart View sheet of an existing workbook. Works on a COPY: the source file is never changed.
    The refreshed copy is saved (label 'Internal') to the output folder and the path is returned; open it with
    excel-mcp for further work. The source must be closed in Excel."""
    src = Path(path)
    if not src.exists():
        raise ToolError(f"File not found: {path}")
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    out = config.OUT_DIR / f"{stamp}_{output_name or ('refreshed_' + src.stem)}{src.suffix}"
    try:
        res = await _call(lambda h: ops.refresh_workbook(h, src, out), timeout=900, what="refresh_workbook")
    except (HostError, ToolError) as e:
        raise _err(e)
    qlog.write({"tool": "sv_refresh_workbook", **res})
    return res


@tool(annotations=LOCAL)
async def sv_show(visible: bool = True) -> dict:
    """Show or hide the Smart View Excel window (hidden by default). Showing lets the user watch."""
    await _call(lambda h: h.show(visible), what="show")
    return {"visible": visible}


@tool(annotations=ADMIN)
async def sv_restart() -> dict:
    """Close and restart the Smart View Excel (only this server's own instance). Use after repeated errors."""
    import asyncio
    global _host
    if _host is not None:
        old, _host = _host, None
        await asyncio.to_thread(old.shutdown)          # never block the event loop
    h = host()
    await _call(lambda _h: None, what="restart")
    return h.status()


# ---------------------------------------------------------------- advanced: full ad hoc API (read-only)
def _sheet(name: str) -> str:
    return ("X_" + _safe_name(name))[:31]


def _grid_preview(h: ExcelHost, sheet: str, max_rows: int = 80) -> dict:
    ws = h.wb.Worksheets(sheet)
    vals = ops._used_values(ws)
    rows = [list(r) for r in vals]
    return {"sheet": sheet, "used_range": ws.UsedRange.Address, "rows": len(rows),
            "grid": rows[:max_rows], "truncated": len(rows) > max_rows}


@tool(annotations=READ)
async def sv_adhoc_open(connection: Connection, sheet: str = "adhoc") -> dict:
    """Advanced: open a fresh ad hoc sheet on a cube with Smart View's default grid (POV in row 1). Then edit
    cells with sv_adhoc_set_cells and run sv_adhoc_action. Returns the grid (cell A1 = grid[0][0])."""
    s = _sheet(sheet)

    def job(h):
        ops.ensure_grid_sheet(h, connection, sheet=s, fresh=True)
        return _grid_preview(h, s)
    try:
        return await _call(job, what="adhoc_open")
    except (HostError, ToolError) as e:
        raise _err(e)


@tool(annotations=READ)
async def sv_adhoc_grid(sheet: str = "adhoc", max_rows: int = 80) -> dict:
    """Advanced: read the current contents of an ad hoc sheet."""
    s = _sheet(sheet)
    return await _call(lambda h: _grid_preview(h, s, max_rows), what="adhoc_grid")


@tool(annotations=READ)
async def sv_adhoc_set_cells(cells: dict[str, str], sheet: str = "adhoc", clear_first: bool = False) -> dict:
    """Advanced: type member names / text into cells of an ad hoc sheet (stored as text), e.g.
    {'B2':'Sep 26','A3':'ac.C30'}. Follow with sv_adhoc_action(action='retrieve')."""
    s = _sheet(sheet)

    def job(h):
        ws = h.wb.Worksheets(s)
        ws.Activate()
        if clear_first:
            ws.Cells.ClearContents()
        for addr, text in cells.items():
            r = ws.Range(addr)
            r.NumberFormat = "@"
            r.Value = text
        return _grid_preview(h, s)
    return await _call(job, what="adhoc_set_cells")


@tool(annotations=READ)
async def sv_adhoc_action(action: Literal["retrieve", "zoom_in", "zoom_out", "keep_only", "remove_only",
                                          "pivot_to_pov", "pivot_to_grid", "pivot"],
                          sheet: str = "adhoc", cell: str | None = None, level: Literal[
                              "next", "all", "bottom", "siblings", "same_level", "same_generation"] = "next",
                          dimension: str | None = None, end_cell: str | None = None) -> dict:
    """Advanced: run a Smart View ad hoc operation on a sheet (read-only; no submit exists).
    zoom_in(cell, level) · zoom_out(cell) · keep_only(cell or range) · remove_only(cell or range) ·
    pivot_to_pov(cell) · pivot_to_grid(dimension, cell) · pivot(cell -> end_cell) · retrieve."""
    s = _sheet(sheet)
    lv = {"next": 0, "all": 1, "bottom": 2, "siblings": 3, "same_level": 4, "same_generation": 5}[level]

    def job(h):
        h.wb.Worksheets(s).Activate()
        if action == "retrieve":
            h.run("B_Retrieve", s)
        elif action == "zoom_in":
            h.run("B_ZoomIn", s, cell, lv, False)
        elif action == "zoom_out":
            h.run("B_ZoomOut", s, cell)
        elif action == "keep_only":
            h.run("B_KeepOnly", s, cell)
        elif action == "remove_only":
            h.run("B_RemoveOnly", s, cell)
        elif action == "pivot_to_pov":
            h.run("B_PivotToPOV", s, cell)
        elif action == "pivot_to_grid":
            h.run("B_PivotToGrid", s, dimension, cell)
        elif action == "pivot":
            h.run("B_Pivot", s, cell, end_cell)
        return _grid_preview(h, s)
    if action != "retrieve" and not cell:
        raise ToolError(f"{action} needs a cell address")
    try:
        res = await _call(job, what=f"adhoc_{action}")
    except (HostError, ToolError) as e:
        raise _err(e)
    qlog.write({"tool": "sv_adhoc_action", "action": action, "sheet": s, "cell": cell, "level": level,
                "dimension": dimension})
    return res


@tool(annotations=WRITE)
async def sv_adhoc_option(option: int, value: str | None = None, sheet: str | None = "adhoc") -> dict:
    """Advanced: get (value omitted) or set a Smart View option. sheet=None means global (careful: global options
    are the user's preferences). Codes: 1 zoom level, 5 indentation (0 none,1 sub items,2 totals),
    6/41 suppress missing rows/cols, 7/42 suppress zero rows/cols, 13 missing label, 16 member display
    (0 name,1 name+desc,2 desc), 33 navigate without data, 44 suppress missing blocks."""
    s = _sheet(sheet) if sheet else ""

    def coerce(v):
        if v is None:
            return None
        if v.lower() in ("true", "false"):
            return v.lower() == "true"
        try:
            return int(v)
        except ValueError:
            return v

    def job(h):
        if value is None:
            r = h._run("B_GetOption", option, s)
            return {"option": option, "sheet": s or "(global)", "value": r[1] if r and r[0] == "OK" else None,
                    "status": r[0] if r else None}
        return {"option": option, "sheet": s or "(global)", "set": coerce(value),
                "status": h._run("B_SetOption", option, coerce(value), s)}
    return await _call(job, what="adhoc_option")


@tool(annotations=READ)
async def sv_member_info(connection: Connection, member: str) -> dict:
    """Advanced: all properties EPM returns for one exact member name (parent, level, generation, child count,
    first child, consolidation operator, storage, aliases, UDAs)."""
    try:
        return await _call(lambda h: ops.find_member(h, connection, member), what="member_info")
    except (HostError, ToolError) as e:
        raise _err(e)


# ---------------------------------------------------------------- analysis in the user's workbook
from . import answer as answer_mod            # noqa: E402
from . import atlas_history, templates, userexcel   # noqa: E402
from .timing import Timer, stage                    # noqa: E402

Target = Literal["active", "new"] | str
_sessions: dict[str, dict[int, dict[str, "userexcel.Pull"]]] = {}   # workbook path -> q -> sheet -> Pull


def _session(path: str, q: int) -> dict:
    return _sessions.setdefault(path.lower(), {}).setdefault(q, {})


def fin(v, kind: str = "heltall") -> str:
    """Financial text for chat, matching the sheet: 1 000 000 / (1 000 000) / – ; % with sign; FTE 1 decimal."""
    if v is None or v == "":
        return ""
    if not isinstance(v, (int, float)):
        return str(v)
    if kind.startswith("prosent"):
        return "–" if abs(v) < 5e-5 else f"{v * 100:+.1f} %".replace(".", ",")
    if abs(v) < 0.5 and kind != "tall1":
        return "–"
    s = f"{abs(v):,.1f}" if kind == "tall1" else f"{abs(v):,.0f}"
    s = s.replace(",", " ").replace(".", ",")
    return f"({s})" if v < 0 else s


def _answer_markdown(spec: dict, res: dict) -> str:
    cols = spec.get("columns", [])
    head = ["Line"] + [c.get("header", "") for c in cols]
    out = ["| " + " | ".join(h.replace("|", "/") for h in head) + " |", "|---|" + "---:|" * len(cols)]
    n = 0
    for sec in spec.get("sections", []):
        out.append(f"| **{sec.get('title', '')}** |" + " |" * len(cols))
        for ln in sec.get("lines", []):
            k = ln.get("key")
            if not k or k not in res["values"]:
                continue
            vals = res["values"][k]
            cells = []
            for j, v in enumerate(vals):
                kind = ln.get("num") or cols[j].get("kind") or "heltall"
                if j < len(ln.get("cells", [])) and isinstance(ln["cells"][j], dict) and "varpct" in ln["cells"][j]:
                    kind = "prosent1"
                cells.append(fin(v, kind))
            label = ln.get("label", k)
            if ln.get("kind") in ("sum", "grand"):
                label = f"**{label}**"
            out.append("| " + label + " | " + " | ".join(cells) + " |")
            n += 1
            if n >= 45:
                out.append("| … see the answer sheet for the rest |" + " |" * len(cols))
                return "\n".join(out)
    return "\n".join(out)


def _result(t, q: int, spec: dict, res: dict, pulls: list, warnings: list) -> dict:
    try:
        if userexcel.com(lambda: t.wb.ReadOnly, limit=10):
            warnings = [f"'{t.name}' is open READ-ONLY (someone or something else holds the file) - the analysis "
                        "is built but cannot be saved under this name."] + warnings
    except Exception:   # noqa: BLE001 - informational only
        pass
    return {
        "workbook": t.path, "question": f"Q{q}", "answer_sheet": res["sheet"],
        "data_sheets": [p.sheet for p in pulls] + ["Kilder"],
        "key_points": spec.get("key_points", []),
        "table": _answer_markdown(spec, res),
        "checks": res["checks"],
        "warnings": warnings + res["warnings"] + [f"formula error on line '{k}'" for k in res["formula_errors"]],
        "style_lint": res["lint"],
        # tn-formelstil: 'block' findings must be fixed (spec_patch) before the answer is presented as done
        "formula_style": res.get("formula_style"),
        "note": "Your workbook was NOT saved - review and save it yourself." if not t.created
        else f"New analysis workbook saved: {t.path}",
    }


def _tidy(t, sheet: str | None = None) -> None:
    """Fixed, neat sheet order (sheetorder: Svar first, other sheets, data per question, Kilder, hidden) and land
    on `sheet`. Our own analysis files get the full order; in a user's own workbook only our sheets are reordered
    among themselves. Cosmetic: never fails the tool."""
    from . import sheetorder
    try:
        with stage("arrange sheets"):
            mode = "full" if (t.created or _is_our_analysis_file(t.path)) else "active"
            sheetorder.arrange_sheets(t.wb, mode, activate=sheet)
    except Exception as e:   # noqa: BLE001
        log.warning("arrange sheets failed in %s: %s", t.name, e)
        if sheet:
            try:
                userexcel.com(t.wb.Worksheets(sheet).Activate)
            except Exception:   # noqa: BLE001
                pass


def _finish(t, q: int, spec: dict, res: dict, saved_new: bool):
    _tidy(t, res["sheet"])
    t.status(None)
    if saved_new:
        userexcel.com(t.wb.Save)


@tool(annotations=WRITE)
async def sv_analyse(question: str, template: Literal["division_pnl", "cost_base", "monthly_trend", "actual_vs_plan"],
                     params: dict, target: str = "session", allow_archive: bool = False, q: int | None = None) -> dict:
    """Answer a financial question as an analysis IN THE USER'S EXCEL: live Smart View data sheets ('Q<n> FinMax' ...),
    a formatted answer sheet 'Q<n> Svar' (TN report style, live formulas, key points, checks, variance drivers, chart)
    and a 'Kilder' log. target: 'session' (default: this session's own analysis workbook in
    the analyser folder (config.ANALYSER_DIR, set in config.toml)\\<project>, created on its first question and
    reused/reopened for every later one),
    'active' (the workbook the user is working in - only when they ask for it), 'new' (a new file that becomes the
    session's workbook), or a path. Archive files are refused unless allow_archive=true. q: extend an existing
    question number (follow-up), else the next free number. Templates and params:
      division_pnl:   entity (e.g. 'un.SY.2TEC'), period ('Sep 26'; default last closed), basis 'ytd'|'month',
                      compare 'ly'|None, drivers bool
      cost_base:      parent ('un.SY.2TEC'), accounts (['ac.C50S','ac.C60','ac.C70','ac.STA20Y']), period, top (25)
      monthly_trend:  entity, accounts, through ('Sep 26')
      actual_vs_plan: entity (FinPlan alias, e.g. '2TEC Technology'), month ('Sep'), year ('FY26'), basis, forecast
    Write the chat reply from the returned key_points/table/checks."""
    fn = templates.TEMPLATES[template]

    def job(h):
        tm = Timer.start()
        with stage("plan (template)"):
            plan = fn(h, **params)
        with stage("open target"):
            t = userexcel.resolve_target(h, target, allow_archive, title=plan.title)
            qn = q or userexcel.next_question(t)
        pulls_by_id, sess = {}, _session(t.path, qn)
        try:
            for ps in plan.pulls:
                with stage("pull"):
                    p = userexcel.pull(h, t, qn, ps.connection, ps.pov, ps.rows, ps.cols,
                                       title=f"Q{qn} {ps.connection}" + (f" {ps.title_suffix}" if ps.title_suffix else ""))
                pulls_by_id[ps.id] = p
                sess[p.sheet] = p
            spec = plan.answer(pulls_by_id)
            with stage("answer"):
                res = answer_mod.build(t, qn, spec, sess)
        except Exception as e:
            removed = userexcel.rollback(t)
            for name in removed:
                sess.pop(name, None)
            raise HostError(f"{e} (the analysis was not completed; removed the partial sheets: "
                            f"{', '.join(removed) or 'none'})") from None
        notes = "; ".join(spec.get("notes", [])[:1])
        with stage("log Kilder"):
            for p in pulls_by_id.values():
                try:
                    userexcel.log_source(t, qn, question, p, notes)
                except Exception as e:   # noqa: BLE001 - the source log must never undo a finished analysis
                    res["warnings"].append(f"could not log {p.sheet} in 'Kilder': {e}")
        warns = [w for p in pulls_by_id.values() for w in p.parsed.pov_problems] + \
                [f"unknown members in {p.sheet}: {', '.join(p.parsed.unknown_members)}"
                 for p in pulls_by_id.values() if p.parsed.unknown_members]
        t.created = t.created or _is_our_analysis_file(t.path)     # an analysis file of ours: save it too
        with stage("save + activate"):
            _finish(t, qn, spec, res, saved_new=t.created)
        return {**_result(t, qn, spec, res, list(pulls_by_id.values()), warns), "timings": tm.summary()}
    try:
        out = await _call_user(job, what=f"sv_analyse {template}")
    except (HostError, ToolError) as e:
        qlog.write({"tool": "sv_analyse", "template": template, "params": params, "error": str(e)})
        raise _err(e)
    qlog.write({"tool": "sv_analyse", "template": template, "params": params, "question": question,
                "workbook": out["workbook"], "q": out["question"], "checks": out["checks"]})
    return out


@tool(annotations=WRITE)
async def sv_ws_pull(question: str, connection: Connection, rows: list[AxisIn], columns: list[AxisIn],
                     pov: dict[str, str] | None = None, target: str = "session", allow_archive: bool = False,
                     q: int | None = None, sheet_title: str | None = None) -> dict:
    """Free-form analysis, step 1 (use sv_analyse when a template fits): put a LIVE Smart View grid (financial NOK
    format) into the user's workbook as 'Q<n> <cube>' and log it in 'Kilder'.
    For several pulls of ONE question: call again with target=<returned 'workbook'> and q=<returned 'q'> (otherwise
    target='new' would create a new file and q would advance). Then call sv_ws_answer(q, spec, target=<workbook>).
    Returns workbook, q, sheet, a data preview and the exact row_keys/col_keys to use in the answer spec."""
    rows_ax = [Axis(a.dimension, a.members) for a in rows]
    cols_ax = [Axis(a.dimension, a.members) for a in columns]

    def job(h):
        tm = Timer.start()
        with stage("open target"):
            t = userexcel.resolve_target(h, target, allow_archive, title=question[:50])
            qn = q or userexcel.next_question(t)
        try:
            with stage("pull"):
                p = userexcel.pull(h, t, qn, connection, pov or {}, rows_ax, cols_ax,
                                   title=sheet_title or f"Q{qn} {connection}")
        except Exception as e:
            removed = userexcel.rollback(t)
            raise HostError(f"{e} (removed the partial sheet: {', '.join(removed) or 'none'})") from None
        _session(t.path, qn)[p.sheet] = p
        with stage("log Kilder"):
            try:
                userexcel.log_source(t, qn, question, p)
            except Exception:   # noqa: BLE001 - the source log must never undo a finished pull
                pass
        _tidy(t, p.sheet)
        t.status(None)
        if t.created:
            with stage("save"):
                userexcel.com(t.wb.Save)
        return {"timings": tm.summary(), "workbook": t.path, "q": qn, "sheet": p.sheet,
                "table": to_markdown(p.parsed.columns, p.parsed.rows, p.parsed.row_labels, limit=40),
                "row_keys": [list(k) for k in p.layout.row_tuples][:200],
                "col_keys": [list(k) for k in p.layout.col_tuples],
                "warnings": p.parsed.pov_problems + ([f"unknown members: {', '.join(p.parsed.unknown_members)}"]
                                                     if p.parsed.unknown_members else []),
                "retrieve_seconds": round(p.seconds, 2)}
    try:
        return await _call_user(job, what="sv_ws_pull")
    except (HostError, ToolError) as e:
        raise _err(e)


SPEC_DOC = """Spec format (all numbers become live formulas):
      {title, subtitle ('NOK · FinMax · Actual · ...'), key_points: [your interpretation],
       auto_points: true (default: factual points drafted from the numbers come first) | false,
       columns: [{header, kind: heltall|tall1|prosent1}],
       sections: [{title, lines: [{key, label, kind: line|sum|grand, num?: tall1,
                   cells: [{pull: '<pull id or sheet>', row: [...], col: [...]} | {var: [i, j]} | {varpct: [i, j]} |
                           {expr: 'rev+cogs'} | {sum: [keys] | 'm*'} | {ref: {key, col}} | {rowsum: [cols]} |
                           {value} | null]}],
                   repeat: {over: ['Jan', ...] | 'months:Jan..Dec' | [{m: 'Jan', ...}, ...],
                            line: {key: 'm{i}', label: '{v}', cells: [{pull: 'a', row: [...], col: ['{v} 26']}, ...]}}}],
       checks: [{label, expr, columns: [..]} | {label, static: 'OK'|'DIFF', detail}], notes: [..], source,
       chart: {type: column|bar|line|waterfall, lines: [keys] | 'm*', column} | {type, row: key, columns: [..]} |
              {type, lines: [keys] | 'm*', columns: [i, j]} (one series per column, e.g. this year vs last year)}
    In a repeat, a pull cell whose member is not in the pull (e.g. 'Oct 26' before it closes) is left empty, and so
    are variances that depend on it. Column indexes in var/varpct/ref/rowsum are 0-based ANSWER columns.
    Formatting: full NOK #,##0, negatives in parentheses; FinMax signs kept; variance = current − comparison."""
SPEC_DOC += """
    Drill lines (after sv_drill): {drill: 'd', where: {COLUMN: 'exact value', ...}, agg: sum|count|debit|credit}
    sums the question's Linjer table (SUMIFS; an exact-match FILTER where Excel would coerce the key). where {} =
    all lines. Reconcile against the drill's live grid: {pull: '<cube_pull sheet>', row: [...], col: [...]}.
    Mixed signs (profile.signs_mixed): no share-of-total column; show debit and credit instead."""


def _load_session(t, q: int) -> dict:
    sess = _sessions.get(t.path.lower(), {}).get(q)
    if not sess:                                       # e.g. after a server restart: rebuild from the workbook
        with stage("reload pulls from workbook"):
            sess = userexcel.load_pulls(t, q)
        if sess:
            _sessions.setdefault(t.path.lower(), {})[q] = sess
    return sess or {}


def _has_drill_cells(spec) -> bool:
    if isinstance(spec, dict):
        return "drill" in spec or any(_has_drill_cells(v) for v in spec.values())
    if isinstance(spec, list):
        return any(_has_drill_cells(v) for v in spec)
    return False


def _drill_rows(workbook: str, q: int) -> dict[str, list[dict]]:
    """alias -> cached drill rows for this question (drafted key points evaluate drill cells from these)."""
    from . import drill_store
    try:
        db = drill_store.connect()
        try:
            out = {}
            for a, i in drill_store.find(db, workbook, q).items():
                meta, rows = drill_store.load(db, i)
                ec = (meta.get("rec") or {}).get("excel_columns")
                if ec and len(ec) < len(meta["columns"]):   # key columns only in Excel: evaluate like the sheet
                    rows = [{k: r.get(k) for k in ec} for r in rows]
                out[a] = rows
            return out
        finally:
            db.close()
    except Exception:   # noqa: BLE001 - drafted points are best effort
        return {}


def _answer_job(t, q: int, raw_spec: dict, ids: dict, sess: dict, tm, extra_warnings: list | None = None) -> dict:
    from . import specs
    with stage("prepare spec"):
        spec = specs.prepare(raw_spec, ids, sess,
                             drills=_drill_rows(t.path, q) if _has_drill_cells(raw_spec) else None)
    # lines Smart View rejected (rc=1000) are missing from the answer: always say so, in the result and on the sheet
    rejected = list(dict.fromkeys(w for p in sess.values() for w in p.parsed.pov_problems
                                  if w.startswith("Smart View rejected")))
    extra_warnings = list(dict.fromkeys((extra_warnings or []) + rejected))
    if rejected:
        spec = {**spec, "notes": list(spec.get("notes") or []) + [
            "Utelatt fordi Smart View avviste medlemmet: " + w.split(": ", 1)[-1].split(". Pick")[0]
            for w in rejected]}
    with stage("answer"):
        res = answer_mod.build(t, q, spec, sess)
    try:
        userexcel.save_spec(t, q, raw_spec, ids)
    except Exception:   # noqa: BLE001 - only needed for patch follow-ups
        pass
    t.created = t.created or _is_our_analysis_file(t.path)     # made by an earlier pull of ours: ours to save
    with stage("save + activate"):
        _finish(t, q, spec, res, saved_new=t.created)
    return {**_result(t, q, spec, res, list(sess.values()), extra_warnings), "timings": tm.summary()}


def _doc(extra: str):
    """Append shared reference text to a tool's docstring (applied before @tool reads it)."""
    def deco(fn):
        fn.__doc__ = (fn.__doc__ or "").rstrip() + "\n    " + extra
        return fn
    return deco


@tool(annotations=WRITE)
@_doc(SPEC_DOC)
async def sv_ws_answer(q: int, spec: dict | None = None, spec_patch: dict | None = None, target: str = "session",
                       allow_archive: bool = False) -> dict:
    """Free-form analysis, step 2 (sv_ws_build does pulls + answer in ONE call): build or rebuild 'Q<n> Svar' from a
    spec whose cells reference the pulls of question q (same target as the pulls, ideally their 'workbook').
    Follow-up: pass spec_patch instead of spec - it is deep-merged into the question's last spec (dicts merge,
    lists/values replace, null deletes), e.g. {"notes": [...]} or {"key_points": [...]}."""
    from . import specs

    def job(h):
        tm = Timer.start()
        with stage("open target"):
            t = userexcel.resolve_target(h, target, allow_archive)
        sess = _load_session(t, q)
        if not sess:
            raise HostError(f"No pulls for Q{q} in '{t.name}'. Run sv_ws_pull first with the same target.")
        stored = userexcel.load_spec(t, q)
        ids = stored[1] if stored else {}
        if spec is None:
            if not stored:
                raise HostError(f"No earlier answer spec for Q{q}; pass spec.")
            raw = specs.merge(stored[0], spec_patch or {})
        else:
            raw = specs.merge(spec, spec_patch or {})
        return _answer_job(t, q, raw, ids, sess, tm)
    try:
        return await _call_user(job, what="sv_ws_answer")
    except (HostError, ToolError) as e:
        raise _err(e)


def _build_job(h, tm, question: str, pulls: list, spec: dict, target: str, allow_archive: bool, q: int | None,
               replace: bool = False) -> dict:
    """All pulls as live Smart View sheets, then the answer (shared by sv_ws_build and sv_concept).
    replace=True (a whole question re-run, e.g. sv_concept with q): the question's old data sheets are parked first
    and deleted only after the new build succeeded (restored if it fails)."""
    with stage("open target"):
        t = userexcel.resolve_target(h, target, allow_archive, title=question[:50])
        qn = q or userexcel.next_question(t)
    parked = {}
    if replace and q:
        with stage("park old data sheets"):
            parked = userexcel.park_question_data(t, qn)
        _sessions.get(t.path.lower(), {}).pop(qn, None)
    sess, ids, warns = _session(t.path, qn), {}, []
    try:
        for i, pi in enumerate(pulls):
            default = f"Q{qn} {pi.connection}" + (f" {pi.id}" if i else "")
            with stage("pull"):
                p = userexcel.pull(h, t, qn, pi.connection, pi.pov or {},
                                   [Axis(a.dimension, a.members) for a in pi.rows],
                                   [Axis(a.dimension, a.members) for a in pi.columns], title=pi.title or default)
            sess[p.sheet] = p
            ids[pi.id] = p.sheet
            warns += p.parsed.pov_problems + ([f"unknown members in {p.sheet}: {', '.join(p.parsed.unknown_members)}"]
                                              if p.parsed.unknown_members else [])
        out = _answer_job(t, qn, spec, ids, sess, tm, warns)
    except Exception as e:
        removed = userexcel.rollback(t)
        for name in removed:
            sess.pop(name, None)
        if parked:
            userexcel.restore_parked(t, parked)
            _sessions.get(t.path.lower(), {}).pop(qn, None)      # reload the restored sheets from the workbook
        raise HostError(f"{e} (the analysis was not completed; removed the partial sheets: "
                        f"{', '.join(removed) or 'none'}" + ("; the previous version was kept" if parked else "")
                        + ")") from None
    if parked:
        with stage("drop old data sheets"):
            userexcel.drop_parked(t, parked)
    with stage("log Kilder"):
        for pi in pulls:
            try:
                userexcel.log_source(t, qn, question, sess[ids[pi.id]])
            except Exception as e:   # noqa: BLE001 - the source log must never undo a finished analysis
                out["warnings"].append(f"could not log {ids[pi.id]} in 'Kilder': {e}")
    _tidy(t, out["answer_sheet"])                     # again: the Kilder log may have added a sheet
    if t.created:
        userexcel.com(t.wb.Save)
    out["timings"] = tm.summary()
    return out


class PullIn(BaseModel):
    id: str = Field(description="Short id used in the spec's pull cells, e.g. 'm' (monthly) or 'ytd'")
    connection: Connection
    rows: list[AxisIn]
    columns: list[AxisIn]
    pov: dict[str, str] | None = None
    title: str | None = Field(default=None, description="Sheet name; default 'Q<n> <cube>' (+ ' <id>' if needed)")


@tool(annotations=WRITE)
@_doc(SPEC_DOC)
async def sv_ws_build(question: str, pulls: list[PullIn], spec: dict, target: str = "session",
                      allow_archive: bool = False, q: int | None = None) -> dict:
    """ONE call for a free-form analysis: every pull as a live Smart View sheet, then 'Q<n> Svar' from the spec, in
    the session's workbook (target 'session' default: one file per session, Q1, Q2 ... | 'active' = the user's open
    workbook, only when asked | 'new' | path; archive refused unless allow_archive). Pull cells in the
    spec use the pull ids. Follow-ups: sv_ws_answer(q, spec_patch=..., target=<workbook>); new data for the same
    question: sv_ws_build(..., q=<q>, target=<workbook>). Returns key points, table, checks, sheets and timings."""

    def job(h):
        return _build_job(h, Timer.start(), question, pulls, spec, target, allow_archive, q)
    try:
        out = await _call_user(job, what="sv_ws_build")
    except (HostError, ToolError) as e:
        qlog.write({"tool": "sv_ws_build", "question": question, "error": str(e)})
        raise _err(e)
    qlog.write({"tool": "sv_ws_build", "question": question, "workbook": out["workbook"], "q": out["question"],
                "checks": out["checks"], "timings": out["timings"]})
    return out


# ---------------------------------------------------------------- sv_drill: ledger lines behind a FinLite cell
_browser = None
_drill_lock = None          # asyncio.Lock: one browser fetch at a time, so the hang budget starts when it runs
_ALIAS = re.compile(r"^[A-Za-z]\w*$")


def browser():
    global _browser
    if _browser is None:
        from .drill_browser import DrillBrowser
        _browser = DrillBrowser()
    return _browser


@contextmanager
def _timed(tm, name: str):
    """Stage time recorded on THIS call's timer (stage() is thread-local, and sv_drill awaits on the event loop)."""
    t0 = time.monotonic()
    try:
        yield
    finally:
        tm.stages[name] = round(tm.stages.get(name, 0.0) + time.monotonic() - t0, 2)


async def _drill_fetch(url: str, max_lines: int) -> dict:
    """browser().fetch, one at a time in this server, with a hang guard: past the budget only OUR headless Chrome
    is killed (the next drill starts a fresh one) and the tool fails readably instead of waiting forever."""
    import asyncio
    import math
    from .drill_browser import DrillBusy
    global _drill_lock
    if _drill_lock is None:
        _drill_lock = asyncio.Lock()
    budget = 300 + 10 * math.ceil(max_lines / config.DRILL_PAGE)      # login + capture (+ one browser restart),
                                                                       # then ~10 s per page
    async with _drill_lock:
        try:
            return await asyncio.wait_for(asyncio.wrap_future(browser().fetch(url, max_lines=max_lines)), budget)
        except asyncio.TimeoutError:
            await asyncio.to_thread(browser().kill)
            raise ToolError(f"The drill page did not finish within {budget} s; the hidden browser was restarted. "
                            "Try again, or narrow the cell.") from None
        except DrillBusy as e:
            raise ToolError(str(e)) from None


_too_wide: set = set()      # cells EPM refused as too wide (EPMFDM-140373) in this server: split them right away


class _KnownTooWide(Exception):
    def __str__(self):
        return "EPMFDM-140373 earlier in this session"


def _too_wide_key(connection: str, cell: dict) -> tuple:
    return (connection, tuple(sorted(cell.items())))


async def _drill_cell(connection: str, cell: dict, max_lines: int, split: str) -> tuple[list[dict], float | None]:
    """Drill a cell (split into child cells when too big). Returns (pieces, the cell's own cube value)."""
    state = {"left": max_lines, "cube": None}
    pieces = await _drill_piece(connection, cell, max_lines, split, state)
    return pieces, state["cube"]


async def _drill_piece(connection: str, cell: dict, max_lines: int, split: str, state: dict,
                       depth: int = 0) -> list[dict]:
    """Drill one cell; if it holds more than DRILL_SPLIT_AT lines (and split='auto'), drill its children instead.
    state['left'] is the line budget still free (max_lines minus the lines already fetched): a cell bigger than
    that fails with its size BEFORE any split (spec 4). Returns pieces: [{cell, cube_value, items, view, secs}]."""
    from . import drill_split
    from .drill_browser import CellTooWide, TooManyLines
    link = await _call(lambda h: ops.drill_link(h, connection, cell), what="drill link")
    if depth == 0:
        state["cube"] = link["cube_value"]
    elif link["cube_value"] is None:                # an empty (#Missing) child of a split: nothing to drill
        return [{"cell": cell, "cube_value": link["cube_value"], "items": [], "view": None, "secs": 0,
                 "note": "empty cell (skipped)"}]
    if not link["url"]:
        return [{"cell": cell, "cube_value": link["cube_value"], "items": [], "view": None, "secs": 0,
                 "note": "no drill report on this cell (data not loaded through Data Integration)"}]
    can_split = split == "auto" and depth < 4
    left = state["left"]
    limit = min(config.DRILL_SPLIT_AT, left) if can_split else left
    try:
        if _too_wide_key(connection, cell) in _too_wide and can_split:
            raise _KnownTooWide()                   # seen before: skip the ~20 s round trip to EPM's refusal
        res = await _drill_fetch(link["url"], limit)
    except (CellTooWide, _KnownTooWide) as e:
        # too many dimensions open at once: narrow by entity (then account) - a period split would not help
        parts = drill_split.plan_split(cell, drill_split.atlas_children(connection),
                                       order=("All Entities", "All Accounts")) if can_split else None
        if not parts:
            raise ToolError(f"EPM refused the drill on {cell} as too wide ({e}). Choose a lower entity or "
                            "account.") from None
        _too_wide.add(_too_wide_key(connection, cell))
        out = []
        for pc in parts:
            out += await _drill_piece(connection, pc, max_lines, split, state, depth + 1)
        return out
    except TooManyLines as e:
        if e.total > left:
            raise ToolError(f"{cell} holds {e.total} lines, more than the {left} still allowed "
                            f"(max_lines={max_lines}). Narrow the cell or raise max_lines.") from None
        parts = drill_split.plan_split(cell, drill_split.atlas_children(connection)) if can_split else None
        if parts:                                   # too big: drill each child cell instead (recursively)
            out = []
            for pc in parts:
                out += await _drill_piece(connection, pc, max_lines, split, state, depth + 1)
            return out
        res = await _drill_fetch(link["url"], left)                       # leaf: page it all
    state["left"] -= len(res["items"])
    return [{"cell": cell, "cube_value": link["cube_value"], **res}]


def _lines_of(pieces: list[dict], view: dict, parent_cube: float | None = None) -> dict:
    """CPU work off the event loop: normalise each piece ONCE, reconcile each piece, and the whole against the
    PARENT cell's own cube value; split_check = sum of the piece cubes - parent (None without a split)."""
    from . import drill_lines as dl
    cols = dl.columns_from_view(view)
    fixes = dl.load_text_fixes(config.TEXT_FIXES)
    rows, piece_recs = [], []
    stats = {"fixed_mojibake": 0, "fixed_by_list": 0, "unfixed": set()}
    for p in pieces:
        pr, st = dl.normalize(p["items"], cols, fixes)
        rows += pr
        stats["fixed_mojibake"] += st["fixed_mojibake"]
        stats["fixed_by_list"] += st["fixed_by_list"]
        stats["unfixed"].update(st["unfixed"])
        piece_recs.append({"cell": p["cell"], **dl.reconcile(p["cube_value"], pr),
                           **({"note": p["note"]} if p.get("note") else {})})
    stats["unfixed"] = sorted(stats["unfixed"])
    piece_cubes = round(sum(p["cube_value"] or 0.0 for p in pieces), 2)
    parent = piece_cubes if parent_cube is None else parent_cube
    rec = dl.reconcile(parent, rows)
    split_check = round(piece_cubes - round(float(parent), 2), 2) if len(pieces) > 1 and parent is not None else None
    return {"cols": [c[0] for c in cols], "rows": rows, "stats": stats, "piece_recs": piece_recs, "rec": rec,
            "profile": dl.profile(rows), "split_check": split_check}


def _gap_warning(rec: dict) -> str:
    share = f" ({abs(rec['gap'] / rec['cube_value']) * 100:.1f} % of the cell)" if rec.get("cube_value") else ""
    return (f"Lines differ from the cube by {rec['gap']:,.2f} NOK{share} (data not loaded via Data Integration, "
            "or postings after the load).")


def _grid_key(connection: str, pov: dict, account: str, periods: list[str]) -> str:
    """Identity of a drill's cube grid: same key = refresh the existing sheet in place."""
    import json
    return json.dumps({"c": connection, "pov": sorted(pov.items()), "a": account, "p": list(periods)})


def _prev_drill(workbook: str, q: int, alias: str) -> dict | None:
    """Metadata of the earlier drill of this question/alias (the cache is saved last, so it is still the old one)."""
    from . import drill_store
    try:
        db = drill_store.connect()
        try:
            did = drill_store.find(db, workbook, q).get(alias)
            return drill_store.meta(db, did) if did else None
        finally:
            db.close()
    except Exception:   # noqa: BLE001 - no earlier drill known: build fresh
        return None


def _restore_lines(t, q: int, alias: str, prev: dict | None, source: str) -> str:
    """A re-drill failed while rewriting the reused table: put the previous lines back from drill.db; if that fails
    too, mark the sheet INCOMPLETE in red. Returns a note for the error message."""
    from . import drill_store, drill_xl
    try:
        if prev is None:
            raise LookupError("no earlier drill in the cache")
        db = drill_store.connect()
        try:
            _, old_rows = drill_store.load(db, prev["id"])
        finally:
            db.close()
        title = (f"Drill-through lines – {' × '.join(prev['cell'].values())} ({prev['connection']}) "
                 f"– forrige uttrekk {prev['created']}")
        if old_rows:
            drill_xl.write_lines(t, q, old_rows, prev["columns"], title=title, source=source, alias=alias)
        else:
            drill_xl.clear_lines(t, q, title=title, source=source, alias=alias)
        return "the previous lines were restored"
    except Exception as e:   # noqa: BLE001
        try:
            drill_xl.mark_incomplete(t, q, alias)
        except Exception:   # noqa: BLE001
            pass
        return f"the previous lines could not be restored ({e}); the sheet is marked INCOMPLETE"


def _cache_drill(**kw) -> str:
    from . import drill_store
    db = drill_store.connect()
    try:
        drill_store.prune(db)
        return drill_store.save(db, **kw)
    finally:
        db.close()


@tool(annotations=WRITE)
async def sv_drill(question: str, cell: dict[str, str], connection: Literal["FinLite"] = "FinLite", id: str = "d",
                   split: Literal["auto", "none"] = "auto", max_lines: int = config.DRILL_MAX_LINES,
                   target: str = "session", q: int | None = None, allow_archive: bool = False) -> dict:
    """Drill a FinLite cell down to its ledger lines (supplier, invoice, voucher, JE source, line text, product)
    through a hidden headless browser (Data Integration drill-through; nothing opens on screen).
    cell: {'All Accounts': 'ac.C40S', 'All Entities': 'un.SY.2WHS', 'All Time Periods': 'Q3 26', ...} - parents
    allowed; other dimensions default to Actual / NOK / Periodic / top member. Cells over 20 000 lines are split
    automatically (months, then cost centres). Writes a live FinLite grid of the cell + 'Q<n> Linjer' (Excel table
    Linjer_Q<n>) to the session workbook (target 'none' = cache only) and returns the reconciliation (cube vs
    lines), a profile (top suppliers/accounts/months/journals, signs_mixed) and text repairs. Then build the report
    with sv_ws_answer(q, spec, target=<workbook>) using {drill: 'd', where: {...}} cells. id: alias of this drill
    in the question (letters/digits/_; default 'd'). Re-drilling the same q and id refreshes its table in place."""
    if not _ALIAS.match(id or ""):
        raise ToolError(f"id '{id}' must start with a letter and contain only letters, digits and _ (e.g. 'd', 'b2').")
    from .drill_browser import redact
    try:
        return await _sv_drill(question, cell, connection, id, split, max_lines, target, q, allow_archive)
    except Exception as e:
        msg = redact(str(e) if isinstance(e, ToolError) else f"{type(e).__name__}: {e}")   # never an sso_token
        qlog.write({"tool": "sv_drill", "question": question, "cell": cell, "error": msg})
        raise ToolError(msg) from None


async def _sv_drill(question, cell, connection, id, split, max_lines, target, q, allow_archive) -> dict:
    import asyncio
    from . import drill_lines as dl
    from . import drill_xl
    tm = Timer()
    with _timed(tm, "drill"):
        pieces, parent_cube = await _drill_cell(connection, cell, max_lines, split)
    view = next((p["view"] for p in pieces if p.get("view")), None)
    if view is None:                                # spec 6: no drill report -> cube value, 0 lines, the reason
        total_cube = parent_cube if parent_cube is not None else sum(p["cube_value"] or 0.0 for p in pieces)
        reasons = sorted({p.get("note") or "the drill page returned no view" for p in pieces})
        out = {"question": question, "cell": cell, "reconciliation": dl.reconcile(total_cube, []),
               "pieces": [{"cell": p["cell"], "cube_value": p["cube_value"], "note": p.get("note")} for p in pieces]
               if len(pieces) > 1 else [], "warnings": ["No drill-through lines: " + "; ".join(reasons)],
               "timings": tm.summary()}
        qlog.write({"tool": "sv_drill", "question": question, "cell": cell, "lines": 0, "note": reasons})
        return out
    with _timed(tm, "normalise + profile"):
        L = await asyncio.to_thread(_lines_of, pieces, view, parent_cube)
    cols, rows, stats, piece_recs, rec = L["cols"], L["rows"], L["stats"], L["piece_recs"], L["rec"]
    out = {"question": question, "cell": cell, "reconciliation": rec, "pieces": piece_recs if len(pieces) > 1 else [],
           "profile": L["profile"], "text_repairs": stats, "columns": cols}
    if L["split_check"] is not None:
        out["split_check"] = L["split_check"]
    warnings = []
    if stats["unfixed"]:
        warnings.append("Text broken in EPM (letters lost at load): " + "; ".join(stats["unfixed"][:5])
                        + " - add corrections to knowledge\\text_fixes.json")
    if rec["status"] == "gap":
        warnings.append(_gap_warning(rec))
    if L["split_check"]:
        warnings.append(f"The split pieces' cube values differ from the cell by {L['split_check']:,.2f} NOK (the "
                        "atlas children may be capped, out of date or shared): the lines are reconciled to the cell "
                        "itself.")

    if target != "none":
        title = f"Drill-through lines – {' × '.join(cell.values())} ({connection})"
        source = "Kilde: Data Integration drill-through (FinLite)."

        def job(h):
            jt = Timer.start()                      # the job runs on the host thread (thread-local timer)
            with stage("open target"):
                t = userexcel.resolve_target(h, target, allow_archive, title=question[:50])
                qn = q or userexcel.next_question(t)
            pov = {**ops.DRILL_DEFAULTS, **{d: m for d, m in cell.items()
                                           if d not in ("All Accounts", "All Time Periods")}}
            periods = list(dict.fromkeys([cell["All Time Periods"]] + [pc["cell"]["All Time Periods"]
                                                                       for pc in piece_recs]))
            key = _grid_key(connection, pov, cell["All Accounts"], periods)
            prev = _prev_drill(t.path, qn, id)                 # the earlier drill of this q/alias, if any
            prev_grid = ((prev or {}).get("rec") or {}).get("grid") or {}
            sess, notes, parked, info = _load_session(t, qn), [], {}, None
            had_table = drill_xl.table_columns(t.wb, drill_xl.table_name(qn, id)) is not None
            lines_started = False
            try:
                with stage("cube grid"):
                    old = prev_grid.get("sheet")
                    if old and prev_grid.get("key") == key and old in sess:
                        p = userexcel.refresh_pull(t, sess[old])           # same grid: refresh in place
                        notes.append(f"grid {old} refreshed in place")
                    else:
                        parked = userexcel.park_sheets(t, [old]) if old else {}
                        for o in parked.values():
                            _session(t.path, qn).pop(o, None)
                        p = userexcel.pull(h, t, qn, connection, pov,
                                           [Axis("All Accounts", [cell["All Accounts"]])],
                                           [Axis("All Time Periods", periods)], title=f"Q{qn} {connection}")
                        if parked:
                            notes.append(f"the cell changed: grid {p.sheet} rebuilt - re-run sv_ws_answer for Q{qn} "
                                         "if its formulas used the old grid")
                _session(t.path, qn)[p.sheet] = p
                lines_started = True
                with stage("lines sheet"):
                    if rows:
                        info = drill_xl.write_lines(t, qn, rows, cols, title=title, source=source, alias=id)
                    else:                           # 0 lines: empty an existing table so Excel matches drill.db
                        info = drill_xl.clear_lines(t, qn, title=title, source=source, alias=id)
            except Exception as e:
                removed = userexcel.rollback(t)
                for name in removed:
                    _session(t.path, qn).pop(name, None)
                if parked:
                    userexcel.restore_parked(t, parked)
                    _sessions.get(t.path.lower(), {}).pop(qn, None)     # reload the restored grid from the workbook
                extra = ""
                if had_table and lines_started:     # the reused table may be half written: put the old lines back
                    extra = "; " + _restore_lines(t, qn, id, prev, source)
                raise HostError(f"{e} (removed the partial sheets: {', '.join(removed) or 'none'}{extra})") from None
            if parked:
                with stage("drop old grid"):
                    userexcel.drop_parked(t, parked)
            with stage("log Kilder"):
                try:
                    userexcel.log_source(t, qn, question, p, notes=f"drill: {rec['lines']} linjer, avvik {rec['gap']}"
                                         + (f"; ark {info['sheet']}" if info else "; ingen linjer"))
                except Exception:   # noqa: BLE001 - the source log must never undo a finished drill
                    pass
            t.status(None)
            t.created = t.created or _is_our_analysis_file(t.path)     # the session's own workbook: ours to save
            _tidy(t, info["sheet"] if info else p.sheet)
            if t.created:
                with stage("save"):
                    userexcel.com(t.wb.Save)
            if info and info.get("columns_changed"):
                notes.append(f"columns changed – re-run sv_ws_answer for Q{qn}")
            elif info and info.get("reused") and rows:
                notes.append(f"re-drill: table {info['table']} refreshed in place (answer formulas stay valid)")
            if info and len(info["columns"]) < len(cols) and rows:
                notes.append(f"{len(rows)} lines > {config.DRILL_XL_FULL}: only the key columns are in Excel "
                             "(excel_columns); where-filters on other columns are not possible in the answer.")
            if not rows:
                notes.append(f"0 lines: table {info['table']} emptied." if info else
                             "0 lines: no lines table was written.")
            return {"workbook": t.path, "q": qn, "lines_sheet": info["sheet"] if info else None,
                    "table": info["table"] if info else None, "excel_columns": info["columns"] if info else [],
                    "cube_pull": {"sheet": p.sheet, "row_keys": [list(k) for k in p.layout.row_tuples],
                                  "col_keys": [list(k) for k in p.layout.col_tuples]},
                    "notes": notes, "_grid": {"sheet": p.sheet, "key": key}, "_excel_stages": dict(jt.stages)}
        try:
            res = await _call_user(job, what="sv_drill")
        except (HostError, ToolError) as e:
            raise _err(e)
        for k, v in res.pop("_excel_stages", {}).items():
            tm.stages[f"excel.{k}"] = round(v, 2)
        grid = res.pop("_grid", None)
        out.update(res)
    else:
        grid = None
    with _timed(tm, "cache"):
        out["drill_id"] = await asyncio.to_thread(
            _cache_drill, connection=connection, cell=cell, rec={**rec, "grid": grid, "excel_columns": out.get("excel_columns")}, columns=cols, rows=rows,
            workbook=out.get("workbook"), q=out.get("q"), alias=id)
    out["warnings"] = warnings
    out["timings"] = tm.summary()
    qlog.write({"tool": "sv_drill", "question": question, "cell": cell, "lines": rec["lines"], "gap": rec["gap"],
                "workbook": out.get("workbook"), "q": out.get("q"), "timings": out["timings"]})
    return out


@tool(annotations=LOCAL)
async def sv_scope(text: str) -> dict:
    """SCOPING - call FIRST for product, profitability, business-area or 'what does X cost/earn' questions. Returns
    matching business concepts (with their blocks: booked / estimated / attributed / not in FinMax) and the account
    and product groups from the business map (knowledge\\) that match, with booking pattern and pitfalls. If a concept
    matches, answer with sv_concept. If none matches: research, PROPOSE the blocks to the user, and after they confirm
    add the concept to knowledge\\concepts.json."""
    from . import concepts as cmod
    return cmod.scope(text)


@tool(annotations=WRITE)
async def sv_concept(concept: str, period: str | None = None, compare: str | None = None, entity: str | None = None,
                     target: str = "session", allow_archive: bool = False, q: int | None = None) -> dict:
    """ONE call for a business-concept analysis (e.g. 'splitt' = Splitt profitability): builds every block of the
    concept from knowledge\\concepts.json - booked lines, estimated lines with their keys, attributed lines and what is
    not in FinMax - as live Smart View sheets + 'Q<n> Svar' in the session's workbook, YTD `period` (default: last
    closed month) vs `compare` (default: same month last year). Use sv_scope first to find the concept."""
    from . import concepts as cmod

    def job(h):
        tm = Timer.start()
        c = cmod.get(concept)
        with stage("plan (concept)"):
            p = period or templates.closed_month(h, "FinMax")
            mo, yy = templates.split_period(p)
            ly = compare or f"{templates.MONTHS[mo - 1]} {yy - 1:02d}"
            pulls, spec = cmod.plan(c, p, ly, entity)
        pull_objs = [PullIn.model_validate({k: v for k, v in pl.items() if v is not None}) for pl in pulls]
        return _build_job(h, tm, f"{c.get('title', concept)} {p}", pull_objs, spec, target, allow_archive, q,
                          replace=True)
    try:
        out = await _call_user(job, what=f"sv_concept {concept}")
    except (HostError, ToolError, ValueError) as e:
        qlog.write({"tool": "sv_concept", "concept": concept, "error": str(e)})
        raise _err(e)
    qlog.write({"tool": "sv_concept", "concept": concept, "workbook": out["workbook"], "q": out["question"],
                "checks": out["checks"], "timings": out["timings"]})
    return out


@tool(annotations=LOCAL)
async def sv_atlas_changes(connection: Connection, dimension: str | None = None, since: str | None = None) -> str:
    """Hierarchy changes between atlas builds: added / removed / moved (new parent) / relabelled members, e.g. cost
    centres moved by a reorganisation. since: ISO date (default: the previous snapshot). Snapshots are taken after
    every sv_atlas_build; rebuild monthly to see changes."""
    db = atlas.connect()
    atlas_history.ensure_schema(db)
    if not atlas_history.list_snapshots(db, connection):
        atlas_history.snapshot_all(db)
    return atlas_history.changes_markdown(atlas_history.changes(db, connection, dimension, since))


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


# ================================================================== xl_*: live work in the user's open workbooks
# Any open workbook, no Smart View involved. Active workbook by default (or name one). Every change is snapshotted
# (xl_revert undoes it - Ctrl+Z cannot undo automation), the screen stays frozen while writing and then jumps to the
# change, and the workbook is NEVER saved: the user reviews and saves. Formulas must pass tn-formelstil.
from . import livexl  # noqa: E402

XL_RULES = ("Formulas follow skill tn-formelstil: SUM over ranges (never a + chain of 3+ cells), label-keyed XLOOKUP "
            "(no INDEX/MATCH(1,1*...), no guessed ranges like $A$5:$Z$400), variance % "
            "=IF(D=0,\"\",(C-D)/ABS(D)), no hard-coded numbers. A result with formula_style.ok=false has BLOCK "
            "findings: fix them (write again) before calling the work done.")


async def _xl(fn, what: str):
    try:
        return await _call_user(lambda h: fn(), what=what)
    except (HostError, ToolError) as e:
        raise _err(e)


@tool(annotations=LOCAL)
async def xl_workbooks() -> list[dict]:
    """All workbooks open in the user's Excel window(s): name, path, sheets, active sheet, selection, unsaved
    changes, read-only, archive flag. Start here when the user talks about 'my workbook' / 'this sheet'."""
    return await _xl(livexl.workbooks, "xl_workbooks")


@tool(annotations=LOCAL)
async def xl_overview(workbook: str | None = None, sheet: str | None = None, preview_rows: int = 12) -> dict:
    """Map of an open workbook (default: the active one): per sheet the used range, formula/constant counts,
    tables, a value preview; plus defined names. Give sheet for a longer preview of one sheet."""
    return await _xl(lambda: livexl.overview(workbook, sheet, preview_rows), "xl_overview")


@tool(annotations=LOCAL)
async def xl_read(range: str | None = None, sheet: str | None = None, workbook: str | None = None,
                  what: Literal["values", "formulas", "both", "formats"] = "values",
                  max_cells: int = livexl.MAX_READ_CELLS) -> dict:
    """Read a range of an open workbook live (unsaved edits included). range: 'B5:F30', a name, or omitted/
    'selection' = what the user has selected. what='formulas' returns {cell: formula} for formula cells only."""
    if what == "formats":
        max_cells = min(max_cells, 600)
    return await _xl(lambda: livexl.read(workbook, sheet, range, what, max_cells), "xl_read")


@tool(annotations=WRITE)
@_doc(XL_RULES)
async def xl_write(range: str, sheet: str | None = None, workbook: str | None = None,
                   values: list[list] | None = None, formulas: list[list] | None = None, formula: str | None = None,
                   number_format: str | None = None, allow_archive: bool = False, label: str | None = None) -> dict:
    """Write into an open workbook live. ONE of: values (2D; text stays text, '=..' is a formula), formulas (2D),
    or formula (ONE formula written over the whole range - relative references adjust per cell, like a fill
    right/down). range: the target, or its top-left cell for values/formulas. number_format: US code, e.g.
    '#,##0;(#,##0);"–"'. Snapshot first -> change_id for xl_revert. Never saves. Archive files refused unless
    allow_archive (only when the user explicitly asked)."""
    return await _xl(lambda: livexl.write(workbook, sheet, range, values, formulas, formula, number_format,
                                          allow_archive, label), "xl_write")


@tool(annotations=WRITE)
async def xl_rows(action: Literal["insert_rows", "delete_rows", "insert_cols", "delete_cols"], at: str,
                  sheet: str | None = None, workbook: str | None = None, force: bool = False,
                  allow_archive: bool = False, label: str | None = None) -> dict:
    """Insert or delete whole rows ('12' or '12:14') or columns ('D' or 'D:E') live. A delete that would break
    formulas (#REF!) is refused unless force=true. Snapshot -> xl_revert. Never saves."""
    return await _xl(lambda: livexl.rows_cols(workbook, sheet, action, at, allow_archive, force, label), "xl_rows")


@tool(annotations=WRITE)
async def xl_add_sheet(name: str, after: str | None = None, workbook: str | None = None,
                       allow_archive: bool = False) -> dict:
    """Add an empty sheet to an open workbook (after the given sheet, default last). xl_revert deletes it."""
    return await _xl(lambda: livexl.add_sheet(workbook, name, after, allow_archive), "xl_add_sheet")


@tool(annotations=WRITE)
async def xl_format(range: str, style: str | None = None, sheet: str | None = None, workbook: str | None = None,
                    number_format: str | None = None, column_width: float | None = None,
                    allow_archive: bool = False, label: str | None = None) -> dict:
    """TN report style on a range: style = a TN cell style ('TN Tittel', 'TN Seksjon', 'TN Kolonne tall', 'TN Note'
    ...) or a number kind with optional emphasis: 'heltall', 'heltall.sum', 'prosent1.hovedsum', 'etikett.sum',
    'tall1', 'faktor', 'belop.input' ... Styles are installed in the workbook if missing. Snapshot -> xl_revert."""
    return await _xl(lambda: livexl.format_range(workbook, sheet, range, style, number_format, column_width,
                                                 allow_archive, label), "xl_format")


@tool(annotations=WRITE)
async def xl_format_sheet(rows: dict[str, str], last_col: str, sheet: str | None = None, workbook: str | None = None,
                          label_cols: list[str] | None = None, num: dict[str, str] | None = None,
                          col_widths: dict[str, float] | None = None, freeze_rows: int | None = None,
                          allow_archive: bool = False) -> dict:
    """Whole sheet in TN report style (skill tn-rapportstil; the same engine as the smartview answers), contents
    untouched. rows: {"1": "title", "2": "subtitle", "4": "header", "5-12": "line", "13": "sum", "15": "grand",
    "17": "note", "3": "section", "6": "text"}; unlisted rows with content = line, empty = blank. last_col: right
    edge of the table ('F'). label_cols default ['B']. num: number kind per value column {"C": "heltall",
    "F": "prosent1"} (default heltall). col_widths {"A": 2.5, "B": 46, "C:F": 17}. Returns style_lint. Snapshot."""
    return await _xl(lambda: livexl.format_sheet(workbook, sheet, last_col, rows, label_cols, num, col_widths,
                                                 freeze_rows, allow_archive), "xl_format_sheet")


@tool(annotations=LOCAL)
async def xl_trace(cell: str, direction: Literal["precedents", "dependents"] = "precedents",
                   sheet: str | None = None, workbook: str | None = None) -> dict:
    """Explain a cell: its formula and value, with row label / column header; precedents = every reference in the
    formula with its value (or sum), formula and labels - enough to explain the number in words; dependents = cells
    on this and other sheets that use it."""
    return await _xl(lambda: livexl.trace(workbook, sheet, cell, direction), "xl_trace")


@tool(annotations=LOCAL)
async def xl_check(sheet: str | None = None, workbook: str | None = None) -> dict:
    """tn-formelstil check of the live formulas in one sheet or the whole open workbook: BLOCK findings (+ chains,
    guessed fixed ranges, legacy INDEX/MATCH lookups, 0 on a zero base, hard-coded numbers) and warnings."""
    return await _xl(lambda: livexl.check(workbook, sheet), "xl_check")


@tool(annotations=LOCAL)
async def xl_changes(workbook: str | None = None, limit: int = 20) -> list[dict]:
    """Recent xl_* changes (newest first, kept 7 days): id, time, workbook, sheet, operation, range, reverted."""
    return livexl.changes(workbook, limit)


@tool(annotations=WRITE)
async def xl_revert(change_id: str | None = None, workbook: str | None = None, force: bool = False) -> dict:
    """Undo an xl_* change from its snapshot (values, formulas, formats, column widths; inserted rows/sheets are
    removed, deleted rows restored). Default: the newest unreverted change in the workbook. Refuses when the cells
    were edited afterwards or later changes on the sheet are not reverted yet, unless force=true."""
    return await _xl(lambda: livexl.revert(change_id, workbook, force), "xl_revert")


def main():
    mcp.run()


if __name__ == "__main__":
    main()
