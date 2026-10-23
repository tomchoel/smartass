"""Cube query tools: sv_dimensions, sv_query, sv_find_member, sv_members, sv_refresh_workbook, the advanced ad hoc API, sv_member_info, sv_scope. Registered on the shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

import datetime as dt
import time
from pathlib import Path
from typing import Literal

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from . import atlas, config, ops, qlog
from .grid import Axis, to_markdown
from .timing import Timer, stage
from .host import ExcelHost, HostError
from .server import LOCAL, READ, WRITE, AxesIn, Connection, axes_in, _call, _err, _safe_name, tool

@tool(annotations=READ)
async def sv_dimensions(connection: Connection) -> dict:
    """List a cube's dimensions (names = top members) and their default axis. Use before sv_query.
    First call starts Excel and signs in via SSO (~15 s)."""
    try:
        return await _call(lambda h: ops.dimensions(h, connection), what="sv_dimensions")
    except (HostError, ToolError) as e:
        raise _err(e)


@tool(annotations=WRITE)
async def sv_query(connection: Connection, rows: AxesIn, columns: AxesIn, ctx: Context,
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
        raise ToolError(f"{action} needs a cell address; pass cell= (e.g. 'B3').")
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


@tool(annotations=LOCAL)
async def sv_scope(text: str) -> dict:
    """SCOPING - call FIRST for product, profitability, business-area or 'what does X cost/earn' questions. Returns
    matching business concepts (with their blocks: booked / estimated / attributed / not in FinMax) and the account
    and product groups from the business map (knowledge\\) that match, with booking pattern and pitfalls. If a concept
    matches, answer with sv_concept. If none matches: research, PROPOSE the blocks to the user, and after they confirm
    add the concept to knowledge\\concepts.json."""
    from . import concepts as cmod
    return cmod.scope(text)
