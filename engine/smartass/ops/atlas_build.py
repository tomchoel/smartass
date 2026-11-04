"""`smartass atlas build` / `smartass atlas status`: (re)build the local atlas (every dimension tree of FinMax,
FinLite and FinPlan) through Smart View zoom-all-levels in the engine's own hidden Excel. Read-only against EPM;
takes a few minutes. Run it when `fin_status` or `smartass doctor` says the atlas is stale and the release
folder has no fresher copy (`smartass setup` copies the monthly one)."""
from __future__ import annotations

import asyncio
import contextlib
import time
from typing import Awaitable, Callable

from ..core import config
from ..knowledge import atlas

CUBES = ("FinMax", "FinLite", "FinPlan")
DUMP_TIMEOUT_S = 600

Call = Callable[..., Awaitable]


async def build(call: Call, connections: list[str] | None = None, dimensions: list[str] | None = None,
                db=None, progress: Callable[[str], None] = print) -> list[dict]:
    """call(fn, timeout=..., what=...) runs fn(host) in the hidden Excel (ExcelHost.call). A full build of a cube
    (no `dimensions`) purges it first, so renamed or dropped dimensions do not linger. Errors are recorded per
    dimension and never stop the build."""
    from ..epm import ops
    from ..xl.host import HostError, InputError
    own = db is None
    db = db or atlas.connect()
    done: list[dict] = []
    try:
        for conn in connections or CUBES:
            try:
                dims = await call(lambda h, c=conn: ops.dimensions(h, c), what=f"atlas {conn} dimensions")
            except (HostError, InputError) as e:
                done.append({"connection": conn, "error": str(e)})
                progress(f"{conn}: {e}")
                continue
            wanted = {ops.canon_dim(conn, d) for d in dimensions} if dimensions else None
            if wanted is None:
                atlas.purge(db, conn)
            for dim in dims:
                if wanted and dim not in wanted:
                    continue
                t0 = time.monotonic()
                try:
                    members = await call(lambda h, c=conn, d=dim: ops.dump_dimension(h, c, d),
                                         timeout=DUMP_TIMEOUT_S, what=f"atlas {conn}/{dim}")
                    atlas.store(db, conn, dim, members, time.monotonic() - t0)
                    done.append({"connection": conn, "dimension": dim, "members": len(members),
                                 "seconds": round(time.monotonic() - t0, 1)})
                    progress(f"{conn}/{dim}: {len(members)} members")
                except (HostError, InputError) as e:
                    atlas.store(db, conn, dim, [], time.monotonic() - t0, str(e))
                    done.append({"connection": conn, "dimension": dim, "error": str(e)})
                    progress(f"{conn}/{dim}: {e}")
    finally:
        if own:
            db.close()
    return done


def run_build(connections: list[str] | None = None, dimensions: list[str] | None = None) -> list[dict]:
    """The CLI path: its own hidden Excel for the build, closed at the end (Smart View options restored)."""
    from ..xl.host import ExcelHost

    async def main():
        h = ExcelHost(visible=False)
        try:
            return await build(h.call, connections, dimensions)
        finally:
            await asyncio.to_thread(h.shutdown)
    return asyncio.run(main())


def status() -> dict:
    with contextlib.closing(atlas.connect()) as db:
        return {**atlas.status(db), "atlas_db": str(config.ATLAS_DB)}
