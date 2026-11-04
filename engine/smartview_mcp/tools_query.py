"""Member lookups for writing answer specs: sv_dimensions, sv_find_member, sv_members, sv_member_info. Registered on the shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

from typing import Literal

from mcp.server.mcpserver.exceptions import ToolError

from smartass.knowledge import atlas
from . import ops
from .host import HostError
from .server import LOCAL, READ, Connection, _call, _err, tool


@tool(annotations=READ)
async def sv_dimensions(connection: Connection) -> dict:
    """List a cube's dimensions (names = top members) and their default axis.
    First call starts Excel and signs in via SSO (~15 s)."""
    try:
        return await _call(lambda h: ops.dimensions(h, connection), what="sv_dimensions")
    except (HostError, ToolError) as e:
        raise _err(e)


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
    member name (e.g. 'un.SY.2WHS', 'ac.C70', 'Yr 26'). Use the names in answer specs."""
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


@tool(annotations=READ)
async def sv_member_info(connection: Connection, member: str) -> dict:
    """Advanced: all properties EPM returns for one exact member name (parent, level, generation, child count,
    first child, consolidation operator, storage, aliases, UDAs)."""
    try:
        return await _call(lambda h: ops.find_member(h, connection, member), what="member_info")
    except (HostError, ToolError) as e:
        raise _err(e)
