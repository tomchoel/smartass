"""Cube query tools: sv_dimensions, sv_query, sv_find_member, sv_members, sv_refresh_workbook, sv_member_info, sv_scope. Registered on the shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

import datetime as dt
import logging
import time
from pathlib import Path
from typing import Literal

import anyio.lowlevel
import anyio.to_thread
from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from . import atlas, config, logsetup, ops, qlog
from .grid import Axis, to_markdown
from .timing import Timer, stage
from .host import HostError, InputError
from .server import LOCAL, READ, WRITE, AxesIn, Connection, axes_in, _call, _err, _safe_name, tool

log = logging.getLogger(__name__)


@tool(annotations=READ)
async def sv_dimensions(connection: Connection) -> dict:
    """List a cube's dimensions (names = top members) and their default axis. Use before sv_query.
    First call starts Excel and signs in via SSO (~15 s)."""
    try:
        return await _call(lambda h: ops.dimensions(h, connection), what="sv_dimensions")
    except (HostError, ToolError) as e:
        raise _err(e)


async def _retrieve(connection, pov, rows_ax, cols_ax, suppress_missing, suppress_zeros, source: str, timeout: float,
                    *, n_cells: int = 0, ctx=None):
    """(layout, parsed, seconds, route, note). Smart View's HTTP protocol first (no Excel) unless source='com'; with
    source='auto' an HTTP failure that is not an input error falls back to Smart View in Excel (note says why), and a
    grid too big for one HTTP request's time budget goes straight to Excel."""
    from . import svhttp
    note = None
    too_big = source == "auto" and n_cells / 200 > config.SVHTTP_TIMEOUT_S
    if source != "com" and not too_big:
        try:
            lay, parsed, secs = await anyio.to_thread.run_sync(lambda: svhttp.shared().retrieve(
                connection, pov, rows_ax, cols_ax, suppress_missing, suppress_zeros), abandon_on_cancel=True)
            return lay, parsed, secs, "http", None
        except InputError as e:
            raise _err(e) from None
        except Exception as e:   # noqa: BLE001 - any HTTP-route failure: the Excel route, or a clear error for 'http'
            msg = f"{type(e).__name__}: {logsetup.redact(e)}"[:300]
            if source == "http":
                raise ToolError(f"HTTP route failed - {msg}. Next step: call sv_query with source='com'.") from None
            log.warning("svhttp failed, using Smart View in Excel: %s", msg,
                        exc_info=not isinstance(e, svhttp.SvHttpError))
            note = f"HTTP route failed ({msg}); the grid was read through Smart View in Excel instead."
            if ctx is not None:
                await ctx.report_progress(1.5, 3, "HTTP route failed; using Smart View in Excel")   # after 1 'retrieving'
    await anyio.lowlevel.checkpoint()          # a pending cancel raises here, before any Excel job is submitted
    try:
        lay, parsed, secs = await _call(
            lambda h: ops.query(h, connection, pov, rows_ax, cols_ax, suppress_missing, suppress_zeros),
            timeout=timeout, what="sv_query")
    except ToolError as e:
        if note is None:
            raise
        err = ToolError(f"{e} (HTTP route also failed: {msg})")
        err.http_note = note
        raise err from None
    return lay, parsed, secs, "com", note


@tool(annotations=WRITE)
async def sv_query(connection: Connection, rows: AxesIn, columns: AxesIn, ctx: Context,
                   pov: dict[str, str] | None = None, suppress_missing: bool = True, suppress_zeros: bool = False,
                   save_xlsx: bool = False, title: str | None = None, large: bool = False,
                   max_preview_rows: int = config.PREVIEW_ROWS,
                   source: Literal["auto", "http", "com"] = "auto") -> dict:
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
    Grids over 50,000 cells need large=true.
    source: 'auto' (default) reads over Smart View's HTTP protocol (no Excel, ~1 s; signs in through the hidden
    browser) and falls back to Smart View in Excel; 'com' forces Excel; 'http' never falls back."""
    rows, columns = axes_in(rows), axes_in(columns)
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
        lay, parsed, secs, route, http_note = await _retrieve(
            connection, pov, rows_ax, cols_ax, suppress_missing, suppress_zeros, source,
            timeout=max(config.CALL_TIMEOUT_S, n_cells / 200), n_cells=n_cells, ctx=ctx)
    except (HostError, ToolError) as e:
        qlog.write({"tool": "sv_query", "connection": connection, "pov": pov, "rows": [r.model_dump() for r in rows],
                    "columns": [c.model_dump() for c in columns], "source": source, "error": str(e),
                    "http_note": getattr(e, "http_note", None)})
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
                "route": route, "cells": lay.n_cells, "rows_returned": len(parsed.rows), "retrieve_s": round(secs, 2),
                "total_s": round(time.monotonic() - t0, 2), "unknown": parsed.unknown_members, "saved": saved,
                "http_note": http_note})
    warnings = parsed.pov_problems + ([f"{parsed.no_access_cells} cells are #NoAccess (no read access)"]
                                      if parsed.no_access_cells else [])
    if parsed.suppressed_rows:
        n = len(parsed.suppressed_rows)
        warnings.append(f"{n} requested row(s) had no data and were left out (suppress_missing): "
                        + "; ".join(parsed.suppressed_rows[:10]) + (" ..." if n > 10 else "")
                        + ". They are valid members with nothing booked - report them as 0 / empty.")
    if parsed.unknown_members:
        warnings.insert(0, "Unknown/unrecognised members (no data returned for them): "
                        + "; ".join(parsed.unknown_members) + ". Use sv_find_member to get the right names.")
    if http_note:
        warnings.append(http_note)
    return {
        "connection": connection,
        "route": route,
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
            raise ToolError("live lookup needs a connection; pass connection= FinMax, FinLite or FinPlan.")
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


@tool(annotations=WRITE)
async def sv_refresh_workbook(path: str, output_name: str | None = None, ctx: Context | None = None) -> dict:
    """Refresh every Smart View sheet of an existing workbook. Works on a COPY: the source file is never changed.
    The refreshed copy is saved (label 'Internal') to the output folder and the path is returned; open it with
    excel-mcp for further work. The source must be closed in Excel."""
    src = Path(path)
    if not src.exists():
        raise ToolError(f"File not found: {path}. Check the path (absolute, to a closed workbook).")
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    out = config.OUT_DIR / f"{stamp}_{output_name or ('refreshed_' + src.stem)}{src.suffix}"
    try:
        def job(h):
            Timer.start()                         # its own timer: progress notes while the refresh runs
            with stage("refresh workbook"):
                return ops.refresh_workbook(h, src, out)
        res = await _call(job, timeout=900, what="refresh_workbook")
    except (HostError, ToolError) as e:
        raise _err(e)
    qlog.write({"tool": "sv_refresh_workbook", **res})
    return res


@tool(annotations=READ)
async def sv_member_info(connection: Connection, member: str) -> dict:
    """Advanced: all properties EPM returns for one exact member name (parent, level, generation, child count,
    first child, consolidation operator, storage, aliases, UDAs)."""
    try:
        return await _call(lambda h: ops.find_member(h, connection, member), what="member_info")
    except (HostError, ToolError) as e:
        raise _err(e)


@tool(annotations=LOCAL)
async def sv_scope(text: str) -> dict:
    """SCOPING - call FIRST for product, profitability, business-area or 'what does X cost/earn' questions. Returns
    matching business concepts (with their blocks: booked / estimated / attributed / not in FinMax) and the account
    and product groups from the business map (knowledge\\) that match, with booking pattern and pitfalls. If a concept
    matches, answer with sv_concept. If none matches: research, PROPOSE the blocks to the user, and after they confirm
    add the concept to knowledge\\concepts.json."""
    from . import concepts as cmod
    return cmod.scope(text)
