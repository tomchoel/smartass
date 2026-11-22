r"""The gate of phase 2, Task 3 (deleted in Task 7): the second engine (epm/rest.py, Planning REST exportdataslice)
against the Smart View provider route (epm/svhttp.py), live, read-only, without Excel. Every grid is fetched on both
routes and the two Fetched compared: the layout, every cell (labels included) in its place, numbers within the
provider's own precision (equal as floats, or within 1e-9 relative).

Parts (run one or all):
  order   members the atlas keeps under their label (periods, FinPlan accounts), asked AGAINST outline order (the
          third, the first, then the second child of a parent) on rows and on columns, per application: the provider
          and REST must agree cell for cell, the provider's rows / columns must differ (else the check could not see a
          swap), and the names REST answers must be in the order asked (each member read alone first, for its name);
  grids   20 different grid requests taken from the stored answers (answers.db: the spec of every answer the run log
          names), fetched on both routes; the largest also over a lowered cell budget, so it is split and stitched;
  probes  what REST answers for an empty cell and #NoAccess (tallied over the grids), for an unknown member, and the
          largest grid one request takes (sequential, doubling, stopped at the first failure).

Prints counts, booleans, seconds and member CODES only - no figure, no label, no business text, no token, cookie or
URL query; an error is shown as its class (and HTTP status) only. A missing sign-in stops the run (exit 2): no
sign-in is started. Exit 1 on any difference.

Usage (from engine/):  uv run --frozen python ../scripts/compare_grid_routes.py [order|grids|probes|all]
"""
from __future__ import annotations

import contextlib
import dataclasses
import json
import re
import sqlite3
import sys
import time

from smartass import answers
from smartass.data.provider import SvHttpProvider, facts_from_parsed, grids
from smartass.epm import oauth, rest, svhttp
from smartass.epm.grid import Axis, parse
from smartass.knowledge import atlas
from smartass.knowledge.resolver import open_resolver
from smartass.model.analyze import plan_requests
from smartass.model.chain import complete_chain, restore_chain
from smartass.model.requests import GridRequest
from smartass.model.spec import AnalysisSpec

N_GRIDS = 20
CANDIDATES = 2000                       # stored answers looked at (newest first) to choose the 20 from
REL = 1e-9
BUDGET_STEPS = (25_000, 50_000, 100_000, 200_000, 400_000, 800_000, 1_600_000)
PROVIDER_SAYS = ("different layout", "appended POV", "Grid is Invalid", "do not match", "out of place")
failed = False
sv = svhttp.shared()


def say(line: str) -> None:
    print(line, flush=True)


def why(e: Exception) -> str:
    """An error as its class and, for an HTTP answer, its status - never its text (it may carry labels)."""
    m = re.search(r"HTTP (\d{3})", str(e))
    return type(e).__name__ + (f" HTTP {m.group(1)}" if m else "")


# ---------------------------------------------------------------- comparing two Fetched
def same(x, y) -> str:
    if x == y and type(x) is type(y):
        return "equal"
    if isinstance(x, float) and isinstance(y, float) and abs(x - y) <= REL * max(1.0, abs(x)):
        return "close"
    return f"differ ({type(x).__name__}/{type(y).__name__})"


def compare(p: svhttp.Fetched, r: svhttp.Fetched) -> dict:
    lay = (p.layout.pov, p.layout.row_tuples, p.layout.col_tuples) == (r.layout.pov, r.layout.row_tuples,
                                                                        r.layout.col_tuples)
    shape = len(p.values) == len(r.values) and all(len(a) == len(b) for a, b in zip(p.values, r.values))
    tally: dict[str, int] = {}
    for a, b in zip(p.values, r.values):
        for x, y in zip(a, b):
            k = same(x, y)
            tally[k] = tally.get(k, 0) + 1
    return {"layout": lay, "shape": shape, **tally}


def ok(c: dict) -> bool:
    return c["layout"] and c["shape"] and not any(k.startswith("differ") for k in c)


def data(f: svhttp.Fetched) -> list[list]:
    lay = f.layout
    return [row[lay.nr:lay.nr + len(lay.col_tuples)] for row in f.values[1 + lay.nc:]]


# ---------------------------------------------------------------- the routes
def rest_session(conn: str):
    return sv.rest_session(conn)


def provider(conn, pov, rows, cols) -> svhttp.Fetched:
    return sv.fetch(conn, pov, rows, cols)


def over_rest(conn, pov, rows, cols, db) -> svhttp.Fetched:
    return rest.fetch_grid(rest_session(conn), conn, pov, rows, cols, db)


class Heard:
    """What REST answered (names per row and column) and the raw form of its non-number cells, tallied."""
    def __init__(self):
        self.answers: list[tuple] = []
        self.raw: dict[str, int] = {}

    def __enter__(self):
        self._answer, self._cell = rest._answer, rest._cell

        def answer(j, lay, rows):
            out = self._answer(j, lay, rows)
            self.answers.append(out)
            return out

        def cell(v):
            got = self._cell(v)
            if not isinstance(got, float):
                kind = "null" if v is None else repr(v) if isinstance(v, str) and len(v) <= 12 else type(v).__name__
                shown = "None (blank)" if got is None else got if got in ("#Missing", "#NoAccess", "#Invalid") else "text"
                key = f"{kind} -> {shown}"
                self.raw[key] = self.raw.get(key, 0) + 1
            return got
        rest._answer, rest._cell = answer, cell
        return self

    def __exit__(self, *exc):
        rest._answer, rest._cell = self._answer, self._cell


# ---------------------------------------------------------------- grid requests from the stored answers
def stored_specs(limit: int) -> list[dict]:
    """The latest stored spec of each answer (workbook, Q) the run log names, newest first."""
    from smartass.core import config
    named = set()
    with open(config.RUN_LOG, encoding="utf-8") as f:
        for line in f:
            with contextlib.suppress(ValueError):
                j = json.loads(line)
                if j.get("workbook") and j.get("q") is not None:
                    named.add((answers._key(j["workbook"]), int(j["q"])))
    db = sqlite3.connect(f"file:{answers.db_path()}?mode=ro", uri=True)
    try:
        rows = db.execute("SELECT wb_key, q, spec FROM answer a WHERE id = (SELECT max(id) FROM answer b "
                          "WHERE b.wb_key = a.wb_key AND b.q = a.q) ORDER BY id DESC").fetchall()
    finally:
        db.close()
    out = []
    for wb, q, spec in rows:
        if (wb, q) in named:
            with contextlib.suppress(ValueError):
                out.append(json.loads(spec))
        if len(out) >= limit:
            break
    return out


def requests_of(raw: dict, resolver, closed) -> list:
    """pipeline.prepare's steps up to the grid requests (without fetching)."""
    raw = restore_chain(raw)
    spec = complete_chain(AnalysisSpec.model_validate(raw), resolver)
    for q in spec.queries:
        if resolver.exists(q.entity, q.cube):
            q.entity = resolver.canonical(q.entity, q.cube)
        q.products = [resolver.canonical(x, q.cube) if resolver.exists(x, q.cube) else x for x in q.products]
    return plan_requests(spec, resolver, closed)


def cells_of(req) -> int:
    return sum(len(c[0].members) * eval_rows(r) for _, r, c, _ in grids(req))


def eval_rows(rows) -> int:
    n = 1
    for a in rows:
        n *= len(a.members)
    return n


# ---------------------------------------------------------------- part 1: order
def siblings(db, conn: str, dim: str, member: str) -> list[str]:
    """The member's parent's children in outline order (the atlas's ord), if there are at least three."""
    for parent, in db.execute("SELECT parent FROM member WHERE connection=? AND dimension=? AND name=? AND "
                              "parent IS NOT NULL ORDER BY depth DESC", (conn, dim, member)).fetchall():
        kids = [n for n, in db.execute("SELECT name FROM member WHERE connection=? AND dimension=? AND parent=? "
                                       "ORDER BY ord", (conn, dim, parent))]
        if len(kids) >= 3:
            return kids
    return []


def against(kids: list[str]) -> list[str]:
    return [kids[2], kids[0], kids[1]]


def ords(db, conn, dim, members) -> list[int]:
    """Each member's place under the parent the three were taken from (they share it)."""
    parent = db.execute("SELECT parent FROM member WHERE connection=? AND dimension=? AND name IN (?, ?, ?) "
                        "GROUP BY parent HAVING count(*) = 3", (conn, dim, *members)).fetchone()[0]
    return [db.execute("SELECT ord FROM member WHERE connection=? AND dimension=? AND name=? AND parent=?",
                       (conn, dim, m, parent)).fetchone()[0] for m in members]


def name_alone(conn, pov, rows, cols, axis: str, member: str, db) -> str:
    """The name REST answers for one member asked alone on that axis."""
    if axis == "rows":
        rows = [Axis(rows[0].dimension, [member])]
    else:
        cols = [Axis(cols[0].dimension, [member])]
    with Heard() as h:
        over_rest(conn, pov, rows, cols, db)
    _, c, heads, _ = h.answers[0]
    return heads[0][0] if axis == "rows" else c[0]


def order_case(tag, conn, pov, rows, cols, axis: str, db) -> bool:
    try:
        return _order_case(tag, conn, pov, rows, cols, axis, db)
    except oauth.NotSignedIn:
        raise
    except Exception as e:   # noqa: BLE001 - shown as its class only
        say(f"order {tag} {conn} {axis}: FAILED {why(e)} | DIFF")
        return False


def _order_case(tag, conn, pov, rows, cols, axis: str, db) -> bool:
    asked = (rows if axis == "rows" else cols)[0]
    o = ords(db, conn, asked.dimension, asked.members)
    t0 = time.monotonic()
    p = provider(conn, pov, rows, cols)
    t1 = time.monotonic()
    with Heard() as h:
        r = over_rest(conn, pov, rows, cols, db)
    t2 = time.monotonic()
    d = data(p)
    distinct = (len({tuple(x) for x in d}) == len(d)) if axis == "rows" else \
        (len({tuple(x) for x in zip(*d)}) == len(d[0]))
    _, c, heads, _ = h.answers[0]
    got = [x[0] for x in heads] if axis == "rows" else list(c)
    names = [name_alone(conn, pov, rows, cols, axis, m, db) for m in asked.members]
    cmp = compare(p, r)
    good = ok(cmp) and distinct and got == names and o != sorted(o)
    say(f"order {tag} {conn} {axis}: members {len(asked.members)} | against outline order {o != sorted(o)} | "
        f"provider {axis} distinct {distinct} | REST names in the order asked {got == names} | layout equal "
        f"{cmp['layout']} | cells {sum(v for k, v in cmp.items() if k not in ('layout', 'shape'))} "
        f"equal {cmp.get('equal', 0)} close {cmp.get('close', 0)} differ "
        f"{sum(v for k, v in cmp.items() if k.startswith('differ'))} | provider {t1 - t0:.1f}s rest {t2 - t1:.1f}s"
        f" | {'OK' if good else 'DIFF'}")
    return good


def part_order(reqs, db) -> bool:
    good = True
    fm = next(r for r in reqs if r.cube == "FinMax")
    pov = {"All Scenarios": fm.scenario, "All Currencies": "NOK", "All Balance Types": fm.balance,
           "All Entities": fm.entities[0]}
    months = against(siblings(db, "FinMax", "All Time Periods", fm.periods[0]))
    acct = [fm.accounts[0]]
    say(f"order FinMax: entity {fm.entities[0]} account {acct[0]} periods {', '.join(months)}")
    good &= order_case("A", "FinMax", pov, [Axis("All Time Periods", months)], [Axis("All Accounts", acct)], "rows",
                       db)
    good &= order_case("B", "FinMax", pov, [Axis("All Accounts", acct)], [Axis("All Time Periods", months)], "cols",
                       db)
    fp = next(r for r in reqs if r.cube == "FinPlan")
    (fpov, _, fcols, _), = grids(fp)[:1]
    fpov = {**fpov, "Entity": fp.entities[0]}
    pmonths = against(siblings(db, "FinPlan", "Period", fcols[0].members[0]))
    parents_seen, tried = set(), 0
    parents = [p for p, in db.execute("SELECT parent FROM member WHERE connection='FinPlan' AND dimension='Account' "
                                      "AND parent IS NOT NULL GROUP BY parent HAVING count(*) BETWEEN 3 AND 30 "
                                      "ORDER BY min(depth)")]
    for parent in parents:
        kids = [n for n, in db.execute("SELECT name FROM member WHERE connection='FinPlan' AND dimension='Account' "
                                       "AND parent=? ORDER BY ord", (parent,))]
        if tuple(kids) in parents_seen or tried >= 25:
            continue
        parents_seen.add(tuple(kids))
        tried += 1
        try:
            rows = data(provider("FinPlan", fpov, [Axis("Account", kids)], [Axis("Period", pmonths)]))
        except oauth.NotSignedIn:
            raise
        except Exception:   # noqa: BLE001 - a parent the provider cannot show (an alias clash): another one
            continue
        with_data = [k for k, x in zip(kids, rows) if sum(isinstance(v, float) for v in x) == len(x)
                     and len(set(x)) == len(x)]
        distinct = list({tuple(rows[kids.index(k)]): k for k in with_data}.values())
        if len(distinct) < 3:
            continue                              # fewer than three children with distinct numbers: another parent
        accts = against(sorted(distinct[:3], key=kids.index))
        say(f"order FinPlan: parents tried {tried}")
        good &= order_case("C", "FinPlan", fpov, [Axis("Account", accts)], [Axis("Period", pmonths)], "rows", db)
        good &= order_case("D", "FinPlan", fpov, [Axis("Account", accts)], [Axis("Period", pmonths)], "cols", db)
        break
    else:
        say(f"order FinPlan: no parent with three children with distinct numbers among {tried} tried")
        good = False
    return good


# ---------------------------------------------------------------- part 2: twenty grids
def features(req, fetched: list[svhttp.Fetched]) -> dict:
    missing = noacc = all_missing = 0
    for f in fetched:
        for row in data(f):
            m = sum(1 for v in row if v in ("#Missing", None))
            missing += m
            noacc += sum(1 for v in row if v == "#NoAccess")
            all_missing += m == len(row)
    return {"cube": req.cube, "grids": len(fetched), "cells": sum(len(data(f)) * len(f.layout.col_tuples)
                                                                    for f in fetched),
            "rows": sum(len(f.layout.row_tuples) for f in fetched), "products": bool(req.products),
            "balance": req.balance, "missing": missing, "noaccess": noacc, "all_missing_rows": all_missing}


def choose(cands: list[tuple]) -> list[tuple]:
    """20 requests that cover the classes the gate names, then as different as possible (cube, shape, size)."""
    picked: list[tuple] = []

    def take(pred):
        for c in sorted(cands, key=lambda c: -c[2]["cells"]):
            if c not in picked and pred(c[2]):
                picked.append(c)
                return
    take(lambda f: f["cube"] == "FinPlan" and f["grids"] > 1)          # several years
    take(lambda f: True)                                                # the largest (split below)
    take(lambda f: f["all_missing_rows"] > 0)                           # a missing / suppressed row
    take(lambda f: f["missing"] > 0 and f["all_missing_rows"] == 0)     # empty cells among numbers
    take(lambda f: f["noaccess"] > 0)
    take(lambda f: f["products"])
    for cube in ("FinMax", "FinLite", "FinPlan"):
        take(lambda f, cube=cube: f["cube"] == cube)
    shapes = set()
    for c in sorted(cands, key=lambda c: (-c[2]["cells"])):
        f = c[2]
        key = (f["cube"], f["rows"] > 20, f["grids"], f["products"], f["balance"], len(str(f["cells"])))
        if len(picked) < N_GRIDS and c not in picked and key not in shapes:
            picked.append(c)
            shapes.add(key)
    for c in cands:
        if len(picked) < N_GRIDS and c not in picked:
            picked.append(c)
    return picked[:N_GRIDS]


def part_grids(reqs_by_spec, db, raw_tally: dict) -> bool:
    good = True
    seen, cands = set(), []
    for reqs in reqs_by_spec:
        for req in reqs:
            key = (req.cube, req.accounts, req.entities, req.periods, req.scenario, req.version, req.balance,
                   req.products)
            if key in seen:
                continue
            seen.add(key)
            try:
                t0 = time.monotonic()
                got = [provider(req.cube, pov, rows, cols) for pov, rows, cols, _ in grids(req)]
                cands.append((req, got, features(req, got), time.monotonic() - t0))
            except oauth.NotSignedIn:
                raise
            except Exception as e:   # noqa: BLE001 - a candidate the provider cannot fetch is not compared
                say(f"candidate {req.cube}: provider failed ({why(e)}) - not chosen")
    say(f"grids: {len(seen)} different requests from {len(reqs_by_spec)} answers, {len(cands)} fetched by the "
        "provider")
    chosen = choose(cands)
    if not any(c[0].cube == "FinLite" for c in chosen):         # the run log holds none: one FinMax request on FinLite
        fm = next((c[0] for c in chosen if c[0].cube == "FinMax" and c[2]["rows"] > 1), None)
        if fm is not None:
            lite = dataclasses.replace(fm, cube="FinLite")
            got = [provider("FinLite", pov, rows, cols) for pov, rows, cols, _ in grids(lite)]
            chosen.append((lite, got, features(lite, got), 0.0))
            say("extra: the run log holds no FinLite request - the last line is a FinMax request of the run log asked "
                "on FinLite (the same members)")
    largest = max(chosen, key=lambda c: c[2]["cells"])
    for i, (req, got, f, secs) in enumerate(chosen, 1):
        per_row = got[0].layout.nr + len(got[0].layout.col_tuples)          # as EPM counts a row
        budgets = [rest.MAX_CELLS] + ([max(per_row, f["rows"] * per_row // 3)] if (req, got, f, secs) == largest
                                      else [])
        for budget in budgets:
            saved = rest.MAX_CELLS
            rest.MAX_CELLS = budget
            try:
                t0 = time.monotonic()
                with Heard() as h:
                    mine = [over_rest(req.cube, pov, rows, cols, db) for pov, rows, cols, _ in grids(req)]
                rs = time.monotonic() - t0
                for k, v in h.raw.items():
                    raw_tally[k] = raw_tally.get(k, 0) + v
                cmps = [compare(p, r) for p, r in zip(got, mine)]
                same_n = len(mine) == len(got)
                eq = same_n and all(ok(c) for c in cmps)
                tot = {k: sum(c.get(k, 0) for c in cmps) for k in {k for c in cmps for k in c} - {"layout", "shape"}}
                warn = sum(len(r.warnings) for r in mine)
                parts = len(h.answers)
                err = ""
            except oauth.NotSignedIn:
                raise
            except Exception as e:   # noqa: BLE001 - a grid REST cannot fetch is a difference
                eq, tot, warn, parts, rs, err = False, {}, 0, 0, time.monotonic() - t0, f" | REST failed: {why(e)}"
            finally:
                rest.MAX_CELLS = saved
            good &= eq
            differ = sum(v for k, v in tot.items() if k.startswith("differ"))
            say(f"{i:02d} {f['cube']}{' (budget ' + str(budget) + ')' if budget != rest.MAX_CELLS else ''}: grids "
                f"{f['grids']} | rows {f['rows']} | cells {f['cells']} | products {f['products']} | balance "
                f"{f['balance']} | missing cells {f['missing']} | all-missing rows {f['all_missing_rows']} | "
                f"#NoAccess {f['noaccess']} | REST requests {parts} | warnings {warn} | equal {tot.get('equal', 0)} "
                f"close {tot.get('close', 0)} differ {differ} | provider {secs:.1f}s rest {rs:.1f}s | "
                f"{'OK' if eq else 'DIFF'}{err}")
            if differ:
                say(f"   differences: {json.dumps({k: v for k, v in tot.items() if k.startswith('differ')})}")
    return good


# ---------------------------------------------------------------- part 3: probes
def part_probes(reqs, db, raw_tally: dict) -> bool:
    fm = next(r for r in reqs if r.cube == "FinMax")
    pov = {"All Scenarios": fm.scenario, "All Currencies": "NOK", "All Balance Types": fm.balance}
    fake = "ac.ZZ_NO_SUCH_MEMBER"
    for where, rows in (("rows", [Axis("All Accounts", [fm.accounts[0], fake]), Axis("All Entities", fm.entities[:1])]),
                        ("POV", [Axis("All Accounts", fm.accounts[:1]), Axis("All Entities", fm.entities[:1])])):
        p2 = {**pov, "All Currencies": fake} if where == "POV" else pov
        try:
            p = provider("FinMax", p2, rows, [Axis("All Time Periods", list(fm.periods[:1]))])
            pv = f"answered ({len(p.values)} rows)"
        except Exception as e:   # noqa: BLE001
            pv = f"failed: {why(e)} ({'; '.join(p for p in PROVIDER_SAYS if p in str(e)) or 'other'})"
        try:
            over_rest("FinMax", p2, rows, [Axis("All Time Periods", list(fm.periods[:1]))], db)
            rv = "answered"
        except Exception as e:   # noqa: BLE001
            rv = f"failed: {why(e)}; names the member {fake in str(e)}"
        say(f"probe unknown member {fake} in {where}: provider {pv} | REST {rv}")
    # the largest grid one request takes: level-0 accounts and entities under the request's, twelve months
    accts = list(dict.fromkeys(m["name"] for m in atlas.descendants(db, "FinMax", "All Accounts", leaves_only=True,
                                                                     limit=5000) if m["dimension"] == "All Accounts"))
    ents = list(dict.fromkeys(m["name"] for m in atlas.descendants(db, "FinMax", "All Entities", leaves_only=True,
                                                                    limit=5000) if m["dimension"] == "All Entities"))
    year = fm.periods[0].split()[-1]
    months = [f"{m} {year}" for m in ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov",
                                      "Dec")]
    say(f"probe budget: level-0 accounts {len(accts)}, entities {len(ents)}")
    saved = rest.MAX_CELLS
    rest.MAX_CELLS = 10 ** 9                      # one request per grid: the budget itself is what is probed
    largest = None
    try:
        for cells in BUDGET_STEPS:
            n_rows = cells // len(months)
            ne = min(len(ents), max(1, n_rows // len(accts) + (n_rows % len(accts) > 0)))
            na = min(len(accts), n_rows // ne)
            rows = [Axis("All Accounts", accts[:na]), Axis("All Entities", ents[:ne])]
            t0 = time.monotonic()
            try:
                real = rest._json
                size = []
                rest._json = lambda text: (size.append(len(text.encode("utf-8"))), real(text))[1]
                try:
                    f = over_rest("FinMax", pov, rows, [Axis("All Time Periods", months)], db)
                finally:
                    rest._json = real
                say(f"probe budget {na * ne * len(months)} cells ({na} x {ne} x {len(months)}): answered in "
                    f"{time.monotonic() - t0:.1f}s ({len(f.values)} rows, answer {size[-1] / 1e6:.1f} MB)")
                largest = rows
            except Exception as e:   # noqa: BLE001
                say(f"probe budget {na * ne * len(months)} cells: failed in {time.monotonic() - t0:.1f}s: {why(e)}")
                break
            if na * ne * len(months) < 0.9 * cells:      # no more members to make it bigger
                break
    finally:
        rest.MAX_CELLS = saved
    if largest is not None:                       # the largest that answered, over the default budget: split
        cols = [Axis("All Time Periods", months)]
        t0 = time.monotonic()
        p = provider("FinMax", pov, largest, cols)
        t1 = time.monotonic()
        with Heard() as h:
            r = over_rest("FinMax", pov, largest, cols, db)
        rs = time.monotonic() - t1
        c = compare(p, r)
        labels = sum(1 for i, (a, b) in enumerate(zip(p.values, r.values)) for j, (x, y) in enumerate(zip(a, b))
                     if same(x, y).startswith("differ") and (i < 1 + p.layout.nc or j < p.layout.nr))
        # the facts an answer is computed from, as data/provider.py makes them from either route's grid
        req = GridRequest("probe", "FinMax", tuple(largest[0].members), tuple(largest[1].members), tuple(months),
                          pov["All Scenarios"], "", pov["All Balance Types"])
        facts = [facts_from_parsed(req, parse(f.layout, f.values, False, False), {m: m for m in months})
                 for f in (p, r)]
        say(f"probe split: the largest grid over MAX_CELLS {rest.MAX_CELLS}: REST requests {len(h.answers)} | layout "
            f"equal {c['layout']} | equal {c.get('equal', 0)} close {c.get('close', 0)} differ "
            f"{sum(v for k, v in c.items() if k.startswith('differ'))} (of them member labels {labels}) | facts "
            f"{len(facts[0][0])}, equal {facts[0] == facts[1]} | provider {t1 - t0:.1f}s rest {rs:.1f}s | "
            f"{'OK' if ok(c) else 'DIFF'}")
        return ok(c)
    return True


def main(argv: list[str]) -> int:
    part = argv[0] if argv else "all"
    try:
        closed = SvHttpProvider(sv).closed_months()
        db = atlas.connect()
        with contextlib.closing(open_resolver(closed["FinPlan"][0] % 100, [])) as resolver:
            specs = stored_specs(CANDIDATES)
            reqs_by_spec = []
            for s in specs:
                with contextlib.suppress(Exception):
                    reqs_by_spec.append(requests_of(s, resolver, closed))
        reqs = [r for rs in reqs_by_spec for r in rs]
        say(f"stored answers {len(specs)} | planned {len(reqs_by_spec)} | requests {len(reqs)}")
        good, tally = True, {}
        if part in ("order", "all"):
            good &= part_order(reqs, db)
        if part in ("grids", "all"):
            good &= part_grids(reqs_by_spec, db, tally)
            say(f"raw REST cells that are not numbers, over the grids: {json.dumps(tally) or 'none'}")
        if part in ("probes", "all"):
            good &= part_probes(reqs, db, tally)
    except oauth.NotSignedIn:
        say("not signed in: stopped (no sign-in is started)")
        return 2
    except Exception as e:   # noqa: BLE001 - its text may carry EPM's labels: the class only
        say(f"stopped: {why(e)}")
        return 3
    say(f"result: {'all equal' if good else 'DIFFERENCES'}")
    return 0 if good else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
