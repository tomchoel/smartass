r"""The gate of phase 2, Task 4 (deleted with the hidden-Excel atlas build in Task 6): the atlas built over the token
session (ops/atlas_build.build_rest: Planning REST's Get Member, epm/members.py) against today's atlas (built by the
hidden Excel's zoom), live and read-only, without Excel. The owner's atlas.db and knowledge.db are NEVER written:
today's atlas is copied (read-only) into the work folder, the build goes to a copy there, every consumer check reads
those copies (this script sets config.ATLAS_DB / KNOWLEDGE_DB to them before anything opens a database) and the
script refuses a work folder that holds the owner's files.

Parts (run one or all; build first - the others read its files):
  build    today's atlas copied to <work>/atlas_today.db; a full REST build into <work>/atlas_rest.db (a copy of it,
           build_into); seconds per pod;
  compare  per connection and dimension: member counts, the keys (name, parent), and for every key its depth, order
           and label. Classes: added since / removed since (counts; codes when fewer than 20), alias changed since
           (a code-keyed member's label: codes; a label-keyed member - its key IS its label - renamed in place: its
           key and its children's parent, paired by place), order of the common keys; anything else is a real
           difference (exit 1);
  labels   a sample of the code-keyed members whose alias changed, fetched over the provider route (live): its label
           there must be the rebuilt atlas's, not today's (the class proven, not assumed);
  knowledge  knowledge/build from each atlas into <work>: errors, warnings and every count;
  offline  the offline test set (golden numbers) with the rebuilt atlas and its knowledge.db (a child process);
  specs    every stored spec of answers.db planned against each atlas: members resolved / not (codes of a miss);
  smoke    grid_smoke: server._prepare for test cases b01 / b07 / b16 against the rebuilt atlas (live, read-only);
  routes   scripts/compare_grid_routes.py grids against the rebuilt atlas (both engines on real grids).
The build part also checks the schema (the member table's seven columns, the side table) and runs release 0.6.0's
writer on a copy of the rebuilt atlas (it must work; the side rows it leaves are then not trusted).

Prints counts, booleans, seconds, Oracle's property values and member CODES only (a label-keyed member - FinPlan's,
periods, scenarios - is counted, never named) - no label, no alias text, no figure, no token, cookie or URL query; an
error is its class (and HTTP status). A missing sign-in stops the run (exit 2). Exit 1 on a real difference.

Usage (from engine/):  uv run --frozen python ../scripts/compare_atlas.py <work folder> [build|compare|labels|knowledge|
offline|specs|smoke|all]
"""
from __future__ import annotations

import contextlib
import json
import re
import sqlite3
import subprocess
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

from smartass.core import config

WORK = Path(sys.argv[1]).resolve()
OWNER = {Path(config.ATLAS_DB).resolve(), Path(config.KNOWLEDGE_DB).resolve()}
TODAY, REBUILT = WORK / "atlas_today.db", WORK / "atlas_rest.db"
K_TODAY, K_REBUILT = WORK / "knowledge_today.db", WORK / "knowledge_rest.db"
if WORK.parent == Path(config.HOME).resolve() or WORK == Path(config.HOME).resolve() or OWNER & {
        p.resolve() for p in (TODAY, REBUILT, K_TODAY, K_REBUILT)}:
    sys.exit("refused: the work folder holds or is the owner's smartass home")
OWNER_ATLAS = Path(config.ATLAS_DB)
# every consumer below opens these, never the owner's (the defaults bind at import: set them first); a child process
# is given its pair
CHILD = len(sys.argv) > 4 and sys.argv[2].startswith("_")
config.ATLAS_DB, config.KNOWLEDGE_DB = (Path(sys.argv[3]), Path(sys.argv[4])) if CHILD else (REBUILT, K_REBUILT)
if OWNER & {Path(config.ATLAS_DB).resolve(), Path(config.KNOWLEDGE_DB).resolve()}:
    sys.exit("refused: the owner's database")

from smartass.epm import members, oauth, svhttp  # noqa: E402
from smartass.epm.grid import Axis, name_from_label  # noqa: E402
from smartass.ops import atlas_build  # noqa: E402

CODES_UNDER = 20
CODE = re.compile(r"[a-z]{2}\.[A-Za-z0-9_.&-]+")       # a member name that is a code (ac.C30, un.SY.2WHS)
SAMPLE = 5
failed = False


def say(line: str) -> None:
    print(line, flush=True)


def why(e: Exception) -> str:
    m = re.search(r"HTTP (\d{3})", str(e))
    return type(e).__name__ + (f" HTTP {m.group(1)}" if m else "")


def ro(path: Path) -> sqlite3.Connection:
    return sqlite3.connect(f"file:{path.as_posix()}?mode=ro", uri=True)


def label_keyed(member: str) -> str:
    """A member whose name is its label, as a digest: compared between the two atlases, never printed."""
    import hashlib
    return "#" + hashlib.sha256(member.encode()).hexdigest()[:16]


def code_keyed(name: str, label: str) -> bool:
    """The key was derived from a 'code alias (prefix)' label: the key is a code, safe to print."""
    return name_from_label(label) is not None and name_from_label(label) == name


# ---------------------------------------------------------------- build
def part_build() -> bool:
    WORK.mkdir(parents=True, exist_ok=True)
    for p in (TODAY, REBUILT):
        p.unlink(missing_ok=True)
    with contextlib.closing(ro(OWNER_ATLAS)) as src, contextlib.closing(sqlite3.connect(TODAY)) as dst:
        src.backup(dst)
    with contextlib.closing(ro(TODAY)) as src, contextlib.closing(sqlite3.connect(REBUILT)) as dst:
        src.backup(dst)
    source = members.Source()
    secs: Counter = Counter()
    read = source.tree

    def timed(conn, dim):
        t0 = time.monotonic()
        try:
            return read(conn, dim)
        finally:
            secs[svhttp.CUBES[conn][1]] += time.monotonic() - t0
    source.tree = timed
    t0 = time.monotonic()
    done, replaced = atlas_build.build_into(REBUILT, lambda db: atlas_build.build_rest(
        source, db=db, progress=lambda m: None), progress=lambda m: say("build: " + m.split(";")[0]))
    total = time.monotonic() - t0
    errors = [d for d in done if d.get("error")]
    skipped = [d for d in done if d.get("skipped")]
    say(f"build: {len(done)} dimensions, {len(errors)} errors {[why(Exception(d['error'])) for d in errors]}, "
        f"{len(skipped)} skipped, replaced {replaced}, {total:.1f} s in all; Get Member seconds per application (pod): "
        f"{ {k: round(v, 1) for k, v in secs.items()} }")
    return replaced and not errors and schema() and older_release()


MEMBER_COLUMNS = ["connection", "dimension", "name", "label", "parent", "depth", "ord"]


def exposure(db) -> dict[tuple[str, str], tuple[int, int]]:
    """Task 4b review, Important 1: per (connection, dimension), the keys that a trusted side row records as the real
    name of a member behind ANOTHER key, and how many of those are not among their own real names. Counts only."""
    owners: dict[tuple[str, str], dict[str, set[str]]] = {}
    for c, d, n, rn in db.execute("SELECT r.connection, r.dimension, r.name, r.real_name FROM member_real r "
                                  "JOIN build b ON b.connection = r.connection AND b.dimension = r.dimension AND "
                                  "b.built_at = r.built_at"):
        owners.setdefault((c, d), {}).setdefault(rn.casefold(), set()).add(n)
    out: dict[tuple[str, str], tuple[int, int]] = {}
    for c, d, n in db.execute("SELECT DISTINCT connection, dimension, name FROM member"):
        keys = owners.get((c, d), {}).get(n.casefold(), set())     # the keys whose members are named n in EPM
        if keys - {n}:
            hit, alone = out.get((c, d), (0, 0))
            out[c, d] = (hit + 1, alone + (n not in keys))
    return out


def schema() -> bool:
    """The member table has exactly the seven columns every release knows; the side table holds every real member."""
    with contextlib.closing(ro(REBUILT)) as db:
        cols = [r[1] for r in db.execute("PRAGMA table_info(member)")]
        side = db.execute("SELECT count(*) FROM member_real").fetchone()[0]
        trusted = db.execute("SELECT count(*) FROM member_real r JOIN build b ON b.connection = r.connection AND "
                             "b.dimension = r.dimension AND b.built_at = r.built_at").fetchone()[0]
        twins = db.execute("SELECT count(*) FROM (SELECT 1 FROM member_real GROUP BY connection, dimension, name, "
                           "parent HAVING count(*) > 1)").fetchone()[0]
        rows = db.execute("SELECT count(*) FROM member").fetchone()[0]
        bare = db.execute("SELECT count(*) FROM (SELECT DISTINCT m.connection, m.dimension, m.name FROM member m WHERE "
                          "NOT EXISTS (SELECT 1 FROM member_real r JOIN build b ON b.connection = r.connection AND "
                          "b.dimension = r.dimension AND b.built_at = r.built_at WHERE r.connection = m.connection "
                          "AND r.dimension = m.dimension AND r.name = m.name))").fetchone()[0]
        shared = exposure(db)
    say(f"schema: member columns are the seven {cols == MEMBER_COLUMNS} | member rows {rows} | side rows {side}, "
        f"trusted {trusted} | (key, parent) pairs with more than one real member {twins} | member keys with no "
        f"trusted side row {bare}")
    say("schema: keys that are EPM's name of a member behind another key (REST answering such a key as itself is "
        "refused), per dimension: " + (", ".join(f"{c}/{d} {n} (not among their own real names {o})"
                                                 for (c, d), (n, o) in sorted(shared.items())) or "none"))
    return cols == MEMBER_COLUMNS and side == trusted and side >= rows


def older_release() -> bool:
    """Release 0.6.0's writer on a copy of the rebuilt atlas: its purge and its positional 7-value INSERT of FinPlan
    (the rows as they are, a new build stamp). It must work; FinPlan's side rows are then no longer trusted, FinMax's
    still are."""
    from smartass.knowledge import atlas
    old = WORK / "atlas_old_writer.db"
    old.unlink(missing_ok=True)
    with contextlib.closing(ro(REBUILT)) as src, contextlib.closing(sqlite3.connect(old)) as dst:
        src.backup(dst)
    with contextlib.closing(sqlite3.connect(old)) as db:
        dims = [d for d, in db.execute("SELECT dimension FROM build WHERE connection='FinPlan'")]
        before = {d: db.execute("SELECT * FROM member WHERE connection='FinPlan' AND dimension=?", (d,)).fetchall()
                  for d in dims}
        try:
            with db:                                   # atlas.purge as 0.6.0 has it
                db.execute("DELETE FROM member WHERE connection=?", ("FinPlan",))
                db.execute("DELETE FROM build WHERE connection=?", ("FinPlan",))
            for d, rows_ in before.items():            # atlas.store as 0.6.0 has it
                with db:
                    db.execute("DELETE FROM member WHERE connection=? AND dimension=?", ("FinPlan", d))
                    db.executemany("INSERT OR IGNORE INTO member VALUES (?,?,?,?,?,?,?)", rows_)
                    db.execute("INSERT OR REPLACE INTO build VALUES (?,?,?,?,?,?)",
                               ("FinPlan", d, time.time() + 1, len(rows_), 1.0, None))
            worked = True
        except sqlite3.Error as e:
            worked = False
            say(f"older release: its writer FAILED ({type(e).__name__})")
        same = all(db.execute("SELECT count(*) FROM member WHERE connection='FinPlan' AND dimension=?",
                              (d,)).fetchone()[0] == len(r) for d, r in before.items())
        finplan = sum(len(atlas.real_names(db, "FinPlan", d, {r[2] for r in rows_})) for d, rows_ in before.items())
        fm = db.execute("SELECT dimension, name FROM member WHERE connection='FinMax' LIMIT 1").fetchone()
        finmax = bool(atlas.real_names(db, "FinMax", fm[0], {fm[1]}))
    old.unlink(missing_ok=True)
    say(f"older release (0.6.0's purge and positional INSERT on a copy, FinPlan): worked {worked}, rows as before "
        f"{same}, FinPlan keys with a trusted real name after it {finplan}, FinMax's still trusted {finmax}")
    return worked and same and finplan == 0 and finmax


# ---------------------------------------------------------------- compare
def rows(db, conn, dim) -> dict:
    return {(n, p): (lab, d, o) for n, lab, p, d, o in db.execute(
        "SELECT name, label, parent, depth, ord FROM member WHERE connection=? AND dimension=?", (conn, dim))}


def renames(t: dict, r: dict) -> dict[str, str]:
    """Label-keyed members renamed in place (old key -> new key): a name only today's atlas has and a name only the
    rebuilt one has, with the same parents (after the renames found so far), depth and children - paired when the
    match is unique. Iterated, so a renamed parent's renamed child pairs too."""
    def sig(rows_: dict, names: set[str], m: dict) -> dict:
        parents, kids, depth = defaultdict(set), defaultdict(set), {}
        for (n, p), (_, d, _) in rows_.items():
            parents[n].add(m.get(p, p))
            depth[n] = d
            if p is not None:
                kids[p].add(m.get(n, n))
        out = defaultdict(list)
        for n in names:
            out[(frozenset(parents[n]), depth[n], frozenset(kids[n]))].append(n)
        return out
    m: dict[str, str] = {}
    for _ in range(4):
        old = {n for n, _ in t} - {n for n, _ in r} - set(m)
        new = {n for n, _ in r} - {n for n, _ in t} - set(m.values())
        a, b = sig(t, old, m), sig(r, new, {})
        found = {x[0]: y[0] for k, x in a.items() if len(x) == 1 and len(b.get(k, ())) == 1 for y in [b[k]]}
        if not found:
            break
        m.update(found)
    return m


def misplaced(seq_t: list, seq_r: list) -> list:
    """The fewest common keys whose removal leaves the two orders equal (the longest increasing subsequence)."""
    import bisect
    pos = {k: i for i, k in enumerate(seq_r)}
    xs = [pos[k] for k in seq_t]
    tails, idx, prev = [], [], [-1] * len(xs)
    for i, x in enumerate(xs):
        j = bisect.bisect_left(tails, x)
        if j == len(tails):
            tails.append(x)
            idx.append(i)
        else:
            tails[j], idx[j] = x, i
        prev[i] = idx[j - 1] if j else -1
    keep, k = set(), idx[-1] if idx else -1
    while k >= 0:
        keep.add(k)
        k = prev[k]
    return [seq_t[i] for i in range(len(seq_t)) if i not in keep]


alias_changed: dict[str, list[tuple[str, str]]] = defaultdict(list)     # conn -> [(dimension, code)]
renamed: dict[str, list[tuple[str, str, str]]] = defaultdict(list)      # conn -> [(dimension, old key, new key)]


def shown(keys) -> str:
    codes = sorted({n for n, _, lab in keys if code_keyed(n, lab)})
    return f" codes {codes}" if keys and len(keys) < CODES_UNDER and len(codes) == len(keys) else (
        f" ({len(codes)} code-keyed)" if keys else "")


def part_compare() -> bool:
    good = True
    with contextlib.closing(ro(TODAY)) as dt, contextlib.closing(ro(REBUILT)) as dr:
        dims = sorted(set(dt.execute("SELECT connection, dimension FROM build").fetchall())
                      | set(dr.execute("SELECT connection, dimension FROM build").fetchall()))
        totals = Counter()
        for conn, dim in dims:
            t, r = rows(dt, conn, dim), rows(dr, conn, dim)
            nt, nr = (n_built(d, conn, dim) for d in (dt, dr))
            m = renames(t, r)
            tm = {(m.get(n, n), m.get(p, p)): v for (n, p), v in t.items()}
            removed = [(n, p, t[n, p][0]) for (n, p) in t if (m.get(n, n), m.get(p, p)) not in r]
            added = [(n, p, r[n, p][0]) for (n, p) in r if (n, p) not in tm]
            common = [k for k in r if k in tm]
            depth = sum(1 for k in common if tm[k][1] != r[k][1])
            code_alias = [(k[0], tm[k][0]) for k in common if tm[k][0] != r[k][0] and code_keyed(k[0], tm[k][0])
                          and code_keyed(k[0], r[k][0])]
            renamed_labels = {k for k in common if k[0] in m.values()}
            other_label = [k for k in common if tm[k][0] != r[k][0] and not (code_keyed(k[0], tm[k][0])
                           and code_keyed(k[0], r[k][0])) and k not in renamed_labels]
            reparented = {n for n, _, _ in removed} & {n for n, _, _ in added}     # one code, another parent
            added = [x for x in added if x[0] not in reparented]
            removed = [x for x in removed if x[0] not in reparented]
            below = descendants(r, reparented)
            depth_moved = [k for k in common if tm[k][1] != r[k][1]]
            depth = sum(1 for k in depth_moved if k[0] not in below)
            seq_t = sorted(common, key=lambda k: tm[k][2])
            seq_r = sorted(common, key=lambda k: r[k][2])
            moved = misplaced(seq_t, seq_r) if seq_t != seq_r else []
            unexplained = [k for k in moved if k[0] not in below]
            alias_changed[conn] += sorted({(dim, c) for c, _ in code_alias})
            renamed[conn] += [(dim, old, new) for old, new in m.items()]
            real = bool(depth or other_label or unexplained)
            good &= not real
            totals.update(added=len(added), removed=len(removed), code_alias=len({c for c, _ in code_alias}),
                          renamed=len(m), reparented=len(reparented), out_of_order=len(moved))
            say(f"{conn}/{dim}: n today {nt} rest {nr} | rows today {len(t)} rest {len(r)} | keys equal "
                f"{set(t) == set(r)} | added since {len(added)}{shown(added)} | removed since {len(removed)}"
                f"{shown(removed)} | alias changed: code-keyed {len(code_alias)}"
                f"{' codes ' + str(sorted({c for c, _ in code_alias})) if code_alias else ''}, label-keyed renamed "
                f"{len(m)} | re-parented since {len(reparented)}"
                f"{' codes ' + str(sorted(reparented)) if reparented and all(map(CODE.fullmatch, reparented)) else ''}"
                f" | of "
                f"{len(common)} common: depth differs {len(depth_moved)} ({depth} not below a re-parented member), "
                f"label differs otherwise {len(other_label)}, out of order {len(moved)} ({len(unexplained)} not "
                f"below a re-parented member){' | REAL DIFFERENCE' if real else ''}")
        say(f"compare: {dict(totals)}; real differences: {not good}")
    return good


def n_built(db, conn: str, dim: str) -> int:
    return (db.execute("SELECT n FROM build WHERE connection=? AND dimension=?", (conn, dim)).fetchone() or [0])[0]


def descendants(r: dict, names: set[str]) -> set[str]:
    """The names below any of names in the rebuilt atlas (a re-parented member's subtree keeps its keys; its place in
    the order and its depth follow the member)."""
    kids = defaultdict(set)
    for n, p in r:
        kids[p].add(n)
    out, todo = set(), list(names)
    while todo:
        for c in kids.get(todo.pop(), ()):
            if c not in out:
                out.add(c)
                todo.append(c)
    return out


# ---------------------------------------------------------------- labels: the alias-changed class proven live
def asked(sv, conn: str, dim: str, other: str, member: str) -> str | None:
    """The label the provider route shows for member (asked by its label: the atlas's derived key is not always EPM's
    name - live, most FinMax projects); None when it refuses it."""
    try:
        f = sv.fetch(conn, None, [Axis(dim, [member])], [Axis(other, [other])])
        return f.values[1 + f.layout.nc][0]
    except oauth.NotSignedIn:
        raise
    except svhttp.SvHttpError:
        return None


def part_labels() -> bool:
    """Per connection, a sample of the members whose alias changed (code-keyed: the same key, another label;
    label-keyed: the key was the label): the provider route must show the REBUILT label when asked by it, and must not
    know today's label any more."""
    good = True
    sv = svhttp.shared()
    with contextlib.closing(ro(TODAY)) as dt, contextlib.closing(ro(REBUILT)) as dr:
        if not alias_changed and not renamed:
            part_compare_quiet()
        samples = [(conn, dim, old, new, f"code {new}") for conn, pairs in alias_changed.items()
                   for dim, new in pairs[:SAMPLE] for old in [new]]
        samples += [(conn, dim, old, new, f"label-keyed #{i + 1}") for conn, pairs in renamed.items()
                    for i, (dim, old, new) in enumerate(pairs[:SAMPLE])]
        for conn, dim, old, new, shown_as in samples:
            other = next(d for d, in dr.execute("SELECT DISTINCT dimension FROM build WHERE connection=? AND "
                                                "dimension<>? ORDER BY dimension", (conn, dim)))
            lab_t = dt.execute("SELECT label FROM member WHERE connection=? AND dimension=? AND name=?",
                               (conn, dim, old)).fetchone()[0]
            lab_r = dr.execute("SELECT label FROM member WHERE connection=? AND dimension=? AND name=?",
                               (conn, dim, new)).fetchone()[0]
            now, before = asked(sv, conn, dim, other, lab_r), asked(sv, conn, dim, other, lab_t)
            ok = now == lab_r and before != lab_t
            say(f"labels {conn}/{dim} {shown_as}: the rebuilt label is EPM's {now == lab_r}; today's label refused "
                f"{before is None}, answered as itself {before == lab_t}")
            good &= ok
    say(f"labels: alias-changed class proven on the sample: {good}")
    return good


def part_compare_quiet() -> None:
    global say
    keep, say = say, lambda line: None
    try:
        part_compare()
    finally:
        say = keep


# ---------------------------------------------------------------- knowledge.db from each atlas
def part_knowledge() -> bool:
    from smartass.knowledge.build import build
    got = {}
    for name, a, k in (("today", TODAY, K_TODAY), ("rest", REBUILT, K_REBUILT)):
        k.unlink(missing_ok=True)
        rep = build(Path(config.SEMANTIC_DIR), a, k)
        with contextlib.closing(ro(k)) as db:
            tables = {t: db.execute(f"SELECT count(*) FROM {t}").fetchone()[0] for t, in db.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'member_fts_%'")}
        got[name] = {"built": rep.built, "errors": len(rep.errors), "warnings": len(rep.warnings),
                     "metrics": rep.metrics, "facts": rep.facts, "concepts": rep.concepts, "rules": rep.rules,
                     "entities": rep.entities, "kpis": rep.kpis, "vocab": rep.vocab, "notes": rep.notes, **tables}
        say(f"knowledge {name}: {got[name]}")
    diff = {k: (got['today'][k], got['rest'][k]) for k in got["today"] if got["today"][k] != got["rest"].get(k)}
    say(f"knowledge: differs {diff}")
    return got["rest"]["built"] and not got["rest"]["errors"]


# ---------------------------------------------------------------- the offline test set (a child process each)
def part_offline() -> bool:
    good = True
    for name, a, k in (("today", TODAY, K_TODAY), ("rest", REBUILT, K_REBUILT)):
        out = subprocess.run([sys.executable, __file__, str(WORK), "_offline", str(a), str(k)], capture_output=True,
                             text=True, timeout=1800)
        line = (out.stdout.strip().splitlines() or ["(no output)"])[-1]
        say(f"offline {name}: {line}")
        good &= name != "rest" or (out.returncode == 0)
    return good


def _offline(a: str, k: str) -> int:
    from smartass.testset import runner
    res = runner.run_offline(runner.load_cases())
    ok = sum(1 for t in res if t.get("ok") is True)
    bad = [t["id"] for t in res if t.get("ok") is False]
    none = sum(1 for t in res if t.get("ok") is None)
    print(f"golden {ok} passed, {len(bad)} failed {bad}, {none} without a golden")
    return 1 if bad else 0


# ---------------------------------------------------------------- stored specs resolve against each atlas
def part_specs() -> bool:
    """Passed when the rebuilt atlas misses no member today's atlas resolves (a miss in both is a stored answer that
    names a member EPM no longer has - Task 3b's class, not the atlas's)."""
    out = {}
    for name, a, k in (("today", TODAY, K_TODAY), ("rest", REBUILT, K_REBUILT)):
        r = subprocess.run([sys.executable, __file__, str(WORK), "_specs", str(a), str(k)], capture_output=True,
                           text=True, timeout=1800)
        lines = r.stdout.strip().splitlines() or ["(no output)", "[]"]
        say(f"specs {name}: {lines[-2] if len(lines) > 1 else lines[0]}")
        out[name] = {tuple(x) for x in json.loads(lines[-1])} if r.returncode == 0 else None
    if out["today"] is None or out["rest"] is None:
        return False
    def shown_(xs):
        return sorted((c, d, m if CODE.fullmatch(m) else "(label-keyed)") for c, d, m in xs)
    say(f"specs: missed only with the rebuilt atlas {shown_(out['rest'] - out['today'])}, only with today's "
        f"{shown_(out['today'] - out['rest'])}")
    return not (out["rest"] - out["today"])


def _specs(a: str, k: str) -> int:
    """Every member a stored spec's grid requests name: in the atlas or not, per connection and dimension."""
    import importlib.util
    spec = importlib.util.spec_from_file_location("cgr", Path(__file__).with_name("compare_grid_routes.py"))
    cgr = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(cgr)
    from smartass.data.provider import SvHttpProvider, grids
    from smartass.knowledge import atlas
    from smartass.knowledge.resolver import open_resolver
    closed = SvHttpProvider(svhttp.shared()).closed_months()
    specs = cgr.stored_specs(100_000)
    resolved, missed, failed_plans = 0, Counter(), 0
    with contextlib.closing(open_resolver(closed["FinPlan"][0] % 100, [])) as resolver, \
            contextlib.closing(atlas.connect(Path(a))) as db:
        for s in specs:
            try:
                reqs = cgr.requests_of(s, resolver, closed)
            except Exception:   # noqa: BLE001 - counted
                failed_plans += 1
                continue
            for req in reqs:
                for pov, rows_, cols, _ in grids(req):
                    for d, mm in [*((x.dimension, m) for x in rows_ + cols for m in x.members),
                                  *((pov or {}).items())]:
                        if atlas.has_member(db, req.cube, d, mm) or mm == d:
                            resolved += 1
                        else:
                            missed[(req.cube, d, mm if CODE.fullmatch(mm) else label_keyed(mm))] += 1
    shown_ = Counter()
    for (c, d, mm), n in missed.items():
        shown_[(c, d, mm if CODE.fullmatch(mm) else "(label-keyed)")] += n
    print(f"stored specs {len(specs)}, not planned {failed_plans}, members asked {resolved + sum(missed.values())}, "
          f"resolved {resolved}, not {sum(missed.values())}: {sorted(shown_.items())}")
    print(json.dumps(sorted(missed)))
    return 0


# ---------------------------------------------------------------- Task 3's grid comparison against the rebuilt atlas
def part_routes() -> bool:
    """scripts/compare_grid_routes.py grids (both engines on the stored answers' grids, Task 3b's wrong-member class)
    with config pointed at the rebuilt atlas and its knowledge.db: the REST engine's real-name check on real grids."""
    r = subprocess.run([sys.executable, __file__, str(WORK), "_routes", str(REBUILT), str(K_REBUILT)],
                       capture_output=True, text=True, timeout=3600)
    lines = [x for x in r.stdout.strip().splitlines() if not re.match(r"\d\d ", x)]   # summaries, not each grid
    for line in lines:
        say(f"routes: {line}")
    say(f"routes: grids OK {sum(1 for x in r.stdout.splitlines() if re.match(r'\d\d .*\| OK$', x))}, exit "
        f"{r.returncode}")
    return r.returncode == 0


def _routes(a: str, k: str) -> int:
    import importlib.util
    spec = importlib.util.spec_from_file_location("cgr", Path(__file__).with_name("compare_grid_routes.py"))
    cgr = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(cgr)
    return cgr.main(["grids"])


# ---------------------------------------------------------------- grid_smoke against the rebuilt atlas
def part_smoke() -> bool:
    r = subprocess.run([sys.executable, __file__, str(WORK), "_smoke", str(REBUILT), str(K_REBUILT)],
                       capture_output=True, text=True, timeout=1800)
    lines = r.stdout.strip().splitlines()
    for line in lines:
        say(f"smoke: {line}")
    return r.returncode == 0 and len(lines) == 3 and all("prepare OK" in x for x in lines)


def _smoke(a: str, k: str) -> int:
    import asyncio
    from smartass import server
    from smartass.model.spec import AnalysisSpec
    from smartass.testset import runner

    async def main() -> int:
        cases = {c["id"]: c for c in runner.load_cases()}
        bad = 0
        for cid in ("b01", "b07", "b16"):
            spec = AnalysisSpec.model_validate(cases[cid]["spec"])
            t0 = time.perf_counter()
            try:
                pre, notes = await server._prepare(spec.model_dump(), server._provider(), spec, complete=True)
                lines = getattr(pre.frame, "lines", {}) or {}
                print(f"{cid}: prepare OK in {time.perf_counter() - t0:.1f} s | queries {len(spec.queries)} | "
                      f"line sets {len(lines)} | notes {len(notes)}", flush=True)
            except Exception as e:   # noqa: BLE001 - the class only
                bad += 1
                print(f"{cid}: FAILED {why(e)}", flush=True)
        return 1 if bad else 0
    return asyncio.run(main())


def main(argv: list[str]) -> int:
    part = argv[0] if argv else "all"
    if CHILD:                            # a child process: config already points at the files given
        return {"_offline": _offline, "_specs": _specs, "_smoke": _smoke, "_routes": _routes}[part](argv[1], argv[2])
    try:
        good = True
        for name, fn in (("build", part_build), ("compare", part_compare), ("labels", part_labels),
                         ("knowledge", part_knowledge), ("offline", part_offline), ("specs", part_specs),
                         ("smoke", part_smoke), ("routes", part_routes)):
            if part in (name, "all"):
                t0 = time.monotonic()
                ok = fn()
                say(f"== {name}: {'passed' if ok else 'FAILED'} ({time.monotonic() - t0:.1f} s)")
                good &= ok
    except oauth.NotSignedIn:
        say("not signed in: run `smartass login` (no sign-in is started here)")
        return 2
    return 0 if good else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[2:]))
