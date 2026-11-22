"""`smartass atlas build` / `smartass atlas status`: (re)build the local atlas (every dimension tree of FinMax,
FinLite and FinPlan). Two routes, both read-only against EPM: over the token session (build_rest: Planning REST's Get
Member, epm/members.py; about 10 s for both pods, no Excel) and, until the old route is deleted, Smart View
zoom-all-levels in the engine's own hidden Excel (build; a few minutes). The REST build writes a new file and swaps it
in only when every dimension was built (build_into). Run it when `fin_status` or `smartass doctor` says the atlas is
stale and the release folder has no fresher copy (`smartass setup` copies the monthly one)."""
from __future__ import annotations

import asyncio
import contextlib
import os
import sqlite3
import tempfile
import time
from pathlib import Path
from typing import Awaitable, Callable

from ..core import config
from ..knowledge import atlas

CUBES = ("FinMax", "FinLite", "FinPlan")
SWAP_TIMEOUT_S = 30.0          # build_into waits at most this long for another process's lock on the atlas
# A dimension EPM lists that the allowlist has no member-tree rule for (epmhttp.MEMBER_TREES is fixed: each rule a
# whole path, no pattern): skipped, its old rows kept, said in the result, the CLI's last line and `atlas status`.
NEW_DIMENSION = "dimension {dim} is new in EPM; this release cannot read it - update smartass"
NOT_SIGNED_IN = "not signed in to EPM - run `smartass login`"   # why a REST build without a token replaced nothing
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
    name, with or without 'All ' and spaces). A full build of a cube purges it first, except a dimension EPM lists
    that the allowlist cannot read: that one is skipped ({"skipped": NEW_DIMENSION}), its rows kept. Errors
    (members.TreeError, svhttp.SvHttpError) are recorded per dimension and never stop the build; oauth.NotSignedIn
    stops it."""
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
            app = svhttp.CUBES[conn][1]
            unread = {dim for rest_name, dim, *_ in dims if not members.readable(app, rest_name)}
            if wanted is None:
                atlas.purge(db, conn, keep=unread)
            for rest_name, dim, *kind in dims:
                if wanted and not wanted & {_canon(rest_name), _canon(dim)}:
                    continue
                if dim in unread:                     # no rule for it: never read, never an error, its rows kept
                    done.append({"connection": conn, "dimension": dim, "skipped": NEW_DIMENSION.format(dim=dim)})
                    progress(f"{conn}/{dim}: {done[-1]['skipped']}")
                    continue
                t0 = time.monotonic()
                try:
                    rows = members.rows(source.tree(conn, rest_name), cube, period=bool(kind and kind[0] == "Period"))
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
               progress: Callable[[str], None] = print, refused: dict[type, str] | None = None
               ) -> tuple[list[dict], bool]:
    """run(db) on a NEW file - a copy of target (its cards and every other table kept) - which replaces target only
    when run returned and recorded no error: never a half-built atlas in place. (done, replaced). When nothing is
    replaced, done ends with {"connection": "-", "error": "atlas not replaced: <why>"} and progress says so; an error,
    or run raising, leaves target untouched; the new file is removed. Every attempt is recorded next to target
    (last_attempt: `atlas status`). The swap is SQLite's backup into target, one transaction: a reader that has
    target open sees the old atlas or the new one; a lock another process holds longer than SWAP_TIMEOUT_S leaves it
    unchanged. refused: {exception type: why} - run raising one of them is a reason like the others (not replaced,
    said and recorded), not an exception."""
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    lock = _lock(target)
    if lock is None:
        _record(target, False, "another atlas build is running", [])
        progress(f"atlas not replaced: another atlas build is running; {target} is unchanged")
        return [{"connection": "-", "error": "atlas not replaced: another atlas build is running"}], False
    new: Path | None = None
    done: list[dict] = []

    def outcome(replaced: bool, why: str = "") -> tuple[list[dict], bool]:
        _record(target, replaced, why, done)
        if not replaced:
            done.append({"connection": "-", "error": f"atlas not replaced: {why}"})
            progress(f"atlas not replaced: {why}; {target} is unchanged")
        return done, replaced
    try:
        fd, name = tempfile.mkstemp(dir=target.parent, prefix=target.name + ".", suffix=".new")   # a name of its own
        os.close(fd)
        new = Path(name)
        new.unlink()
        if target.exists():
            try:
                _copy(target, new)
            except (sqlite3.Error, TimeoutError) as e:
                return outcome(False, f"{target.name} could not be read ({e})")
        try:
            with contextlib.closing(atlas.connect(new)) as db:
                done = run(db)
        except BaseException as e:
            why = next((w for t, w in (refused or {}).items() if isinstance(e, t)), None)
            if why is not None:
                return outcome(False, why)
            _record(target, False, f"the build stopped ({type(e).__name__})", done)
            raise
        failed = [d for d in done if d.get("error")]
        if failed or not done:
            return outcome(False, f"{len(failed)} {'dimension or cube' if len(failed) == 1 else 'dimensions or cubes'} "
                                  "failed" if failed else "nothing was built (no such dimension?)")
        try:
            _copy(new, target)
        except (sqlite3.Error, TimeoutError) as e:
            return outcome(False, f"{target.name} could not be written ({e})")
        return outcome(True)
    finally:
        if new is not None:
            new.unlink(missing_ok=True)
        _unlock(lock)


def _lock(target: Path):
    """The build's lock next to the atlas (one build at a time); None when another build holds it. Windows releases
    it when the holder's process ends, so a crashed build leaves no stale lock."""
    import msvcrt
    fh = open(Path(target).with_name(Path(target).name + ".build.lock"), "a+b")   # noqa: SIM115 - held until _unlock
    try:
        fh.seek(0)
        msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        fh.close()
        return None
    return fh


def _unlock(fh) -> None:
    import msvcrt
    with contextlib.suppress(OSError):
        fh.seek(0)
        msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
    fh.close()


def _last_path(target: Path) -> Path:
    return Path(target).with_name(Path(target).name + ".last.json")


def _record(target: Path, replaced: bool, why: str, done: list[dict]) -> None:
    """The attempt, for `atlas status`: when, replaced or not, why, the dimensions this release could not read."""
    import datetime as dt
    import json
    rec = {"at": dt.datetime.now().isoformat(timespec="seconds"), "replaced": replaced, "why": why,
           "skipped": [f"{d['connection']}: {d['skipped']}" for d in done if d.get("skipped")]}
    with contextlib.suppress(OSError):
        _last_path(target).write_text(json.dumps(rec, indent=1), encoding="utf-8")


def last_attempt(target: Path | None = None) -> dict | None:
    """The last REST build's attempt on target (default the atlas): {at, replaced, why, skipped}; None before any."""
    import json
    try:
        return json.loads(_last_path(target or config.ATLAS_DB).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


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
        done, replaced = build_into(config.ATLAS_DB, lambda db: build_rest(
            members.Source(), connections, dimensions, db), refused={oauth.NotSignedIn: NOT_SIGNED_IN})
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
    ok = sorted({d["connection"] for d in done if d.get("dimension") and not d.get("error") and not d.get("skipped")})
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
        return {**atlas.status(db), "cards": atlas.cards_status(db), "atlas_db": str(config.ATLAS_DB),
                "last_build": last_attempt()}
