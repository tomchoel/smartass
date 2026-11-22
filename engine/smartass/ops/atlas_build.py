"""`smartass atlas build` / `smartass atlas status`: (re)build the local atlas (every dimension tree of FinMax,
FinLite and FinPlan). Two routes, both read-only against EPM: over the token session (build_rest: Planning REST's Get
Member, epm/members.py; about 10 s for both pods, no Excel) and, until the old route is deleted, Smart View
zoom-all-levels in the engine's own hidden Excel (build; a few minutes). The REST build writes a new file and swaps it
in only when every dimension was built (build_into). Run it when `fin_status` or `smartass doctor` says the atlas is
stale and the release folder has no fresher copy (`smartass setup` copies the monthly one)."""
from __future__ import annotations

import asyncio
import contextlib
import sqlite3
import time
from pathlib import Path
from typing import Awaitable, Callable

from ..core import config
from ..knowledge import atlas

CUBES = ("FinMax", "FinLite", "FinPlan")
SWAP_TIMEOUT_S = 30.0          # build_into waits at most this long for another process's lock on the atlas
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


def _canon(name: str) -> str:
    return name.removeprefix("All ").replace(" ", "").casefold()


def build_rest(source, connections: list[str] | None = None, dimensions: list[str] | None = None, db=None,
               progress: Callable[[str], None] = print) -> list[dict]:
    """build's shape over the token session. source: members.Source (dimensions(conn) -> [(REST name, atlas name)],
    tree(conn, REST name) -> Get Member's tree). `dimensions`: any spelling of a dimension (its atlas name, its REST
    name, with or without 'All ' and spaces). A full build of a cube purges it first; errors (members.TreeError,
    svhttp.SvHttpError) are recorded per dimension and never stop the build; oauth.NotSignedIn stops it."""
    from ..epm import members, svhttp
    own = db is None
    db = db or atlas.connect()
    done: list[dict] = []
    wanted = {_canon(d) for d in dimensions} if dimensions else None
    try:
        for conn in connections or CUBES:
            cube = svhttp.CUBES[conn][2]
            try:
                dims = source.dimensions(conn)
            except (members.TreeError, svhttp.SvHttpError) as e:
                done.append({"connection": conn, "error": str(e)})
                progress(f"{conn}: {e}")
                continue
            if wanted is None:
                atlas.purge(db, conn)
            for rest_name, dim in dims:
                if wanted and not wanted & {_canon(rest_name), _canon(dim)}:
                    continue
                t0 = time.monotonic()
                try:
                    rows = members.rows(source.tree(conn, rest_name), cube)
                    atlas.store(db, conn, dim, rows, time.monotonic() - t0)
                    done.append({"connection": conn, "dimension": dim, "members": len(rows),
                                 "seconds": round(time.monotonic() - t0, 1)})
                    progress(f"{conn}/{dim}: {len(rows)} members")
                except (members.TreeError, svhttp.SvHttpError) as e:
                    atlas.store(db, conn, dim, [], time.monotonic() - t0, str(e))
                    done.append({"connection": conn, "dimension": dim, "error": str(e)})
                    progress(f"{conn}/{dim}: {e}")
    finally:
        if own:
            db.close()
    return done


def build_into(target: Path, run: Callable[[sqlite3.Connection], list[dict]],
               progress: Callable[[str], None] = print) -> tuple[list[dict], bool]:
    """run(db) on a NEW file - a copy of target (its cards and every other table kept) - which replaces target only
    when run returned and recorded no error:
    never a half-built atlas in place. (done, replaced). An error, or run raising, leaves target untouched (and
    says so); the new file is removed. The swap is SQLite's backup into target, one transaction: a reader that has
    target open sees the old atlas or the new one; a lock another process holds longer than SWAP_TIMEOUT_S leaves it
    unchanged (not replaced, said so)."""
    target = Path(target)
    new = target.with_name(target.name + ".new")
    new.unlink(missing_ok=True)
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        if target.exists():
            try:
                _copy(target, new)
            except (sqlite3.Error, TimeoutError) as e:
                progress(f"atlas not built: {target} could not be read ({e}); it is unchanged")
                return [{"connection": "-", "error": f"the atlas could not be read: {e}"}], False
        with contextlib.closing(atlas.connect(new)) as db:
            done = run(db)
        failed = [d for d in done if d.get("error")]
        if failed or not done:
            why = (f"{len(failed)} {'dimension' if len(failed) == 1 else 'dimensions'} or cubes failed" if failed
                   else "nothing was built (no such dimension?)")
            progress(f"atlas not replaced: {why}; {target} is unchanged")
            return done, False
        try:
            _copy(new, target)
        except (sqlite3.Error, TimeoutError) as e:
            progress(f"atlas not replaced: {target} could not be written ({e}); it is unchanged")
            return [*done, {"connection": "-", "error": f"the atlas could not be replaced: {e}"}], False
        return done, True
    finally:
        new.unlink(missing_ok=True)


def _copy(src_path: Path, dst_path: Path) -> None:
    """SQLite's online backup of src into dst, one step: dst is replaced in one transaction. Python's backup retries
    for as long as another connection holds a lock; this gives up after SWAP_TIMEOUT_S (TimeoutError), dst unchanged."""
    deadline = time.monotonic() + SWAP_TIMEOUT_S

    def waiting(status, remaining, total):
        if time.monotonic() > deadline:
            raise TimeoutError(f"in use by another process for over {SWAP_TIMEOUT_S:g} s")
    with contextlib.closing(sqlite3.connect(src_path, timeout=1.0)) as src, \
            contextlib.closing(sqlite3.connect(dst_path, timeout=1.0)) as dst:
        src.backup(dst, progress=waiting)


def run_build(connections: list[str] | None = None, dimensions: list[str] | None = None, excel: bool = False
              ) -> list[dict]:
    """The CLI path (and `smartass knowledge refresh`'s): over the token session into a new file swapped in when
    complete (build_into) - the default since the gate (docs/audit/2026-11-22-atlas-rest.md); excel: the old route,
    its own hidden Excel for the build, closed at the end (Smart View options restored; deleted in Task 6). A full
    build (no `dimensions`) then derives the behaviour cards of the cubes built (cards_after_build) - after the
    swap."""
    if not excel:
        from ..epm import members, oauth
        try:
            done, replaced = build_into(config.ATLAS_DB, lambda db: build_rest(
                members.Source(), connections, dimensions, db))
        except oauth.NotSignedIn:
            return [{"connection": "-", "error": "not signed in to EPM: run `smartass login`; the atlas is unchanged"}]
        if replaced and not dimensions:
            cards_after_build(done)
        return done
    from ..xl.host import ExcelHost

    async def main():
        h = ExcelHost(visible=False)
        try:
            return await build(h.call, connections, dimensions)
        finally:
            await asyncio.to_thread(h.shutdown)
    done = asyncio.run(main())
    if not dimensions:
        cards_after_build(done)
    return done


def cards_after_build(done: list[dict], derive: Callable | None = None, progress: Callable[[str], None] = print
                      ) -> list[dict]:
    """Behaviour cards for the cubes this build stored without error (FinLite copies FinMax). A card failure is a
    warning only: the atlas is kept and the build's result does not change."""
    ok = sorted({d["connection"] for d in done if d.get("dimension") and not d.get("error")})
    if not ok:
        return []
    if derive is None:
        from ..knowledge import cards_fetch
        derive = cards_fetch.run
    try:
        res = derive(ok)
    except Exception as e:   # noqa: BLE001 - cards never fail the atlas build
        progress(f"warning: behaviour cards not derived ({e}); the atlas is kept")
        return []
    for d in res:
        if d.get("error"):
            progress(f"warning: cards {d['connection']}/{d['dimension']}: {d['error']}")
    return res


def status() -> dict:
    with contextlib.closing(atlas.connect()) as db:
        return {**atlas.status(db), "cards": atlas.cards_status(db), "atlas_db": str(config.ATLAS_DB)}
