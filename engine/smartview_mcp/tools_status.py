"""Status / atlas / admin tools: sv_status, sv_atlas_build, sv_atlas_status, sv_restart, sv_atlas_changes. Registered on the shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

import os
import time

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from smartass.knowledge import atlas
from . import atlas_history, qlog
from smartass.ops import doctor
from smartass.epm import ops
from smartass.core import config
from . import server as core                      # live module ref: read core._host at call time
from smartass.xl.host import HostError
from .server import _STARTED, ADMIN, LOCAL, Connection, _call, host, tool

# ---------------------------------------------------------------- primary tools
@tool(annotations=LOCAL)
async def sv_status() -> dict:
    """Health of the Smart View engine: Excel pid, connected sheets, Smart View version, restarts, atlas age,
    which drill browsers can start (drill_browser), and the doctor checks (config, Excel, Smart View, EPM, output
    folder, HTTP API, update). Cheap; does not start Excel or a browser."""
    import asyncio
    doc = await asyncio.to_thread(doctor.run)      # network checks (EPM 5 s, HTTP 2 s) off the event loop
    db = atlas.connect()
    st = atlas.status(db)
    from smartass.epm import sso
    h = core._host.status() if core._host else {"excel_running": False, "smartview_version": sso.smartview_version()}
    from smartass.epm.drill_browser import browser_health
    try:
        drill = browser_health()
    except Exception as e:   # noqa: BLE001 - a health probe never fails the status call
        drill = {"error": str(e)[:200]}
    return {**h, "server_started": _STARTED, "behind_reload_proxy": bool(os.environ.get("SVKIT_PROXY")),
            "drill_browser": drill,
            "doctor": doctor.as_dict(doc),
            "atlas": {"oldest_age_days": st["oldest_age_days"], "stale": st["stale"],
                      "dimensions_built": len(st["built"])},
            "query_log": str(config.QUERY_LOG), "output_dir": str(config.OUT_DIR)}


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


@tool(annotations=ADMIN)
async def sv_restart() -> dict:
    """Close and restart the Smart View Excel (only this server's own instance). Use after repeated errors."""
    import asyncio
    if core._host is not None:
        old, core._host = core._host, None
        await asyncio.to_thread(old.shutdown)          # never block the event loop
    h = host()
    await _call(lambda _h: None, what="restart")
    return h.status()


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
