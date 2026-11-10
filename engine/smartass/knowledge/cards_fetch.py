"""Read value slices through SvHttp and store derived behaviour cards in atlas.db (plan phase B, Tasks 3-4).
Read-only: ordinary retrieve grids only. One value per member per slice, in chunked grids (about 0.3-0.5 s per
2 000 rows once the session is open); every parent is then compared with its children offline (cards.py)."""
from __future__ import annotations

import datetime as dt
import sqlite3
import time
from collections import defaultdict
from dataclasses import dataclass
from typing import Callable

from ..epm.grid import INVALID, NOACCESS, Axis
from . import atlas
from .cards import account_types, operators, settle_time_balance, time_balance

PNL_TYPES = ("revenue", "expense")

CHUNK = 2000
ACCOUNT_DIMS = {"All Accounts", "Account"}


@dataclass(frozen=True)
class Slice:
    """One POV with the column dimension and its members; every column is one value slice."""
    pov: dict
    cdim: str
    cols: tuple[str, ...]


def last_full_year() -> int:
    return dt.date.today().year - 1


def slices(conn: str, dim: str, year: int) -> list[Slice]:
    """The value slices per cube and dimension for the last full year `year`. A later slice only fills operators
    the earlier ones left unknown (a child without data there)."""
    yy = year % 100
    fm = {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"}
    fp = {"Scenario": "Actual", "Version": "Working", "Years": f"FY{yy:02d}", "All Balance Types": "Periodic"}
    target = {**fp, "Scenario": "Target", "Version": "Final", "Years": f"FY{yy + 1:02d}"}   # FinPlan volumes live here
    table = {
        ("FinMax", "All Accounts"): [Slice({**fm, "All Entities": "un.SY"}, "All Time Periods",
                                           (f"Yr {yy:02d}", f"Yr {yy - 1:02d}"))],
        ("FinMax", "All Entities"): [Slice({**fm, "All Time Periods": f"Yr {yy:02d}"}, "All Accounts",
                                           ("ac.C70", "ac.C30"))],
        ("FinPlan", "Account"): [Slice({**fp, "Entity": "BANORWAY"}, "Period", ("YearTotal",)),
                                 Slice({**target, "Entity": "BANORWAY"}, "Period", ("YearTotal",))],
        ("FinPlan", "Entity"): [Slice({**fp, "Period": "YearTotal"}, "Account", ("OPEX", "Revenue"))],
    }
    return table[(conn, dim)]


DERIVED = {"FinMax": ("All Accounts", "All Entities"), "FinPlan": ("Account", "Entity")}   # accounts first


# FinPlan (live 10.11): 'Found duplicate aliases or member name conflicting with alias under parent ...' and
# 'Duplicate name <m> found in grid; cannot resolve. Please replace it with qualified name.'
AMBIGUOUS = ("duplicate aliases", "qualified name")


def grid(sv, conn: str, pov: dict, dim: str, names: list[str], cdim: str, cols: list[str],
         chunk: int = CHUNK, clashes: set[str] | None = None
         ) -> tuple[dict[str, dict[str, float | None]], dict[str, set[str]]]:
    """({member: {column: value or None}}, {column: members with no known value}) over chunked grids of `names` x
    `cols`. #Missing is None (never 0). No known value: #NoAccess / #Invalid, and a member the provider cannot
    resolve because its name clashes with an alias (FinPlan, live 10.11: the whole grid fails; the chunk is split
    until the clashing members are found and left out). `clashes`: members already known to clash (not asked
    again); the ones found here are added to it, so the next slice of the dimension skips the search."""
    vals: dict[str, dict[str, float | None]] = {}
    blocked: dict[str, set[str]] = {c: set() for c in cols}

    def part(ns: list[str]) -> None:
        try:
            f = sv.fetch(conn, pov, [Axis(dim, ns)], [Axis(cdim, list(cols))])
        except Exception as e:   # noqa: BLE001 - only the alias clash is handled, anything else is raised
            if not any(m in str(e) for m in AMBIGUOUS):
                raise
            if len(ns) == 1:
                known.add(ns[0])
                return
            part(ns[:len(ns) // 2])
            part(ns[len(ns) // 2:])
            return
        lay, v = f.layout, f.values
        for j, n in enumerate(ns):
            row = vals.setdefault(n, {})
            for k, c in enumerate(cols):
                x = v[1 + lay.nc + j][lay.nr + k]
                if isinstance(x, (int, float)) and not isinstance(x, bool):
                    row[c] = float(x)
                else:
                    row[c] = None
                    if isinstance(x, str) and x.strip() in NOACCESS | INVALID:
                        blocked[c].add(n)

    known = clashes if clashes is not None else set()
    ask = [n for n in names if n not in known]
    for i in range(0, len(ask), chunk):
        part(ask[i:i + chunk])
    for n in names:
        if n in known:
            vals[n] = dict.fromkeys(cols)
            for c in cols:
                blocked[c].add(n)
    return vals, blocked


def tree(db: sqlite3.Connection, conn: str, dim: str):
    """(pairs [(name, parent)] in atlas order, names, kids {parent: [child]}, parents_of {name: [parent]})."""
    pairs = db.execute("SELECT name, parent FROM member WHERE connection=? AND dimension=? ORDER BY ord",
                       (conn, dim)).fetchall()
    names = list(dict.fromkeys(n for n, _ in pairs))
    kids: dict[str, list[str]] = defaultdict(list)
    parents_of: dict[str, list[str]] = defaultdict(list)
    for n, p in pairs:
        if p:
            kids[p].append(n)
            parents_of[n].append(p)
    return pairs, names, dict(kids), parents_of


def ancestry(name: str, parents_of: dict[str, list[str]]) -> dict[str, int]:
    """{member or ancestor: distance} over every hierarchy the member is shared into, breadth first (the nearest
    occurrence counts; atlas.ancestors follows one parent only)."""
    dist, front = {name: 0}, [name]
    while front:
        nxt = []
        for x in front:
            for p in parents_of.get(x, []):
                if p not in dist:
                    dist[p] = dist[x] + 1
                    nxt.append(p)
        front = nxt
    return dist


def derive_dimension(sv, db: sqlite3.Connection, conn: str, dim: str, chunk: int = CHUNK,
                     year: int | None = None, clashes: set[str] | None = None) -> int:
    """Operators, account type and shared flag for every (member, parent) of one dimension; replaces its cards
    (time balance is filled by derive_time_balance). `clashes`: see grid. Returns the number of card rows."""
    pairs, names, kids, parents_of = tree(db, conn, dim)
    if not names:                         # a failed atlas dump: keep the last good cards (and FinLite's copy of them)
        raise ValueError(f"no {conn}/{dim} members in the atlas; the cards are kept")
    vals, blocks = [], []
    clashes = set() if clashes is None else clashes
    for s in slices(conn, dim, year or last_full_year()):
        got, blocked = grid(sv, conn, s.pov, dim, names, s.cdim, list(s.cols), chunk, clashes)
        for c in s.cols:
            vals.append({n: got[n][c] for n in names})
            blocks.append(blocked[c])
    ops = operators(kids, *vals, blocked=blocks)

    def sign_value(n: str) -> float | None:
        return next((v[n] for v in vals if v.get(n)), None)

    types = {}
    if dim in ACCOUNT_DIMS:
        types = account_types(kids, {n: ancestry(n, parents_of) for n in names}, {n: sign_value(n) for n in names})
    rows = [(n, p, ops.get((p, n)), None, types.get(n), int(len(parents_of[n]) > 1)) for n, p in pairs if p]
    atlas.store_cards(db, conn, dim, rows)
    return len(rows)


MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def time_slices(conn: str, year: int) -> list[Slice]:
    """(pov, Period, 12 months + the year total) for the account time balance; later slices only for the accounts
    the earlier ones left undecided. FinMax: one Periodic grid is enough - Accumulated Jan equals Periodic Jan and
    Accumulated Dec equals Periodic Yr (probe 10.11), so no Accumulated retrieves."""
    yy = year % 100
    if conn == "FinMax":
        acc = slices("FinMax", "All Accounts", year)[0]
        return [Slice(acc.pov, acc.cdim, (*(f"{m} {yy:02d}" for m in MONTHS), f"Yr {yy:02d}"))]
    return [Slice(s.pov, "Period", (*MONTHS, "YearTotal")) for s in slices("FinPlan", "Account", year)]


def derive_time_balance(sv, db: sqlite3.Connection, conn: str, chunk: int = CHUNK, year: int | None = None,
                        clashes: set[str] | None = None) -> int:
    """cards.time_balance for every account of FinMax or FinPlan, written to every card row of the member (run after
    derive_dimension). Returns the number of accounts with a time balance."""
    dim = "All Accounts" if conn == "FinMax" else "Account"
    types = dict(db.execute("SELECT name, max(account_type) FROM card WHERE connection=? AND dimension=? GROUP BY name",
                            (conn, dim)).fetchall())
    names = list(types)
    found: dict[str, str | None] = dict.fromkeys(names)
    for s in time_slices(conn, year or last_full_year()):
        todo = [n for n in names if found[n] is None and types[n] not in PNL_TYPES]
        if not todo:
            break
        got, _ = grid(sv, conn, s.pov, dim, todo, s.cdim, list(s.cols), chunk, clashes)
        for n in todo:
            row = got[n]
            found[n] = time_balance(row[s.cols[-1]], [row[c] for c in s.cols[:-1]], movements=conn == "FinMax")
    found = {n: settle_time_balance(tb, types[n]) for n, tb in found.items()}
    with db:
        db.execute("UPDATE card SET time_balance=NULL WHERE connection=? AND dimension=?", (conn, dim))
        db.executemany("UPDATE card SET time_balance=? WHERE connection=? AND dimension=? AND name=?",
                       [(tb, conn, dim, n) for n, tb in found.items() if tb])
    return sum(1 for tb in found.values() if tb)


def copy_finlite(db: sqlite3.Connection) -> int:
    """FinLite shares FinMax's members: its cards are FinMax's wherever the FinLite atlas has the same (dimension,
    name, parent). Replaces every FinLite card; returns the rows copied. Run after FinMax's time balance."""
    with db:
        db.execute("DELETE FROM card WHERE connection='FinLite'")
        cur = db.execute(
            "INSERT INTO card SELECT 'FinLite', c.dimension, c.name, c.parent, c.operator, c.time_balance, "
            "c.account_type, c.shared, c.derived_at FROM card c WHERE c.connection='FinMax' AND EXISTS ("
            "SELECT 1 FROM member m WHERE m.connection='FinLite' AND m.dimension=c.dimension AND m.name=c.name "
            "AND m.parent=c.parent)")
    return cur.rowcount


def derive_all(sv, db: sqlite3.Connection, connections: list[str] | None = None,
               progress: Callable[[str], None] = print, year: int | None = None) -> list[dict]:
    """Per wanted cube: operators per dimension, then the account time balance; last the FinLite copy (so it gets
    FinMax's time balance too). Errors are recorded per step and never stop the rest."""
    want = set(connections or ("FinMax", "FinLite", "FinPlan"))
    done: list[dict] = []

    def step(conn: str, what: str, fn: Callable[[], int]) -> None:
        t0 = time.monotonic()
        try:
            done.append({"connection": conn, "dimension": what, "rows": fn(),
                         "seconds": round(time.monotonic() - t0, 1)})
        except Exception as e:   # noqa: BLE001 - one step's failure must not stop the others
            done.append({"connection": conn, "dimension": what, "error": str(e)[:300]})
        progress(_line(done[-1]))
    for conn, dims in DERIVED.items():
        if conn not in want:
            continue
        clashes: dict[str, set[str]] = {d: set() for d in dims}
        for dim in dims:
            step(conn, dim, lambda c=conn, d=dim: derive_dimension(sv, db, c, d, year=year, clashes=clashes[d]))
        step(conn, f"{dims[0]} (time balance)",
             lambda c=conn, d=dims[0]: derive_time_balance(sv, db, c, year=year, clashes=clashes[d]))
    if want & {"FinMax", "FinLite"}:
        done.append({"connection": "FinLite", "dimension": "(copy of FinMax)", "rows": copy_finlite(db)})
        progress(_line(done[-1]))
    return done


def _line(d: dict) -> str:
    if d.get("error"):
        return f"cards {d['connection']}/{d['dimension']}: {d['error']}"
    return f"cards {d['connection']}/{d['dimension']}: {d['rows']} rows" + (
        f" ({d['seconds']} s)" if "seconds" in d else "")


def run(connections: list[str] | None = None, progress: Callable[[str], None] = print) -> list[dict]:
    """The CLI path: the shared headless HTTP route on the real atlas, closed at the end."""
    import contextlib

    from ..epm import drill_browser, svhttp
    try:
        with contextlib.closing(atlas.connect()) as db:
            return derive_all(svhttp.shared(), db, connections, progress)
    finally:
        drill_browser.shutdown_shared()
