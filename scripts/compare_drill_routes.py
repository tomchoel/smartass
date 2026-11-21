r"""Live gate of the drill over the token (plan 2026-11-19, Task 7): the same FinLite cells drilled on two routes,
by default the browser route (SV_DRILL_ROUTE=browser: drill_direct's warm tab, the page route behind it) and the
token route (SV_DRILL_ROUTE=token: drill_token over epmhttp), one fresh Python process per route because config reads
the route at import, and compared line for line. The browser route is deleted (Task 9) only after the gate in
docs/audit/2026-11-19-drill-token-vs-browser.md has passed.

Usage (from the repo root, signed in on plan1; no Excel is used: drillstep's Excel fallbacks raise here):
    engine\.venv\Scripts\python.exe scripts\compare_drill_routes.py [--cell N ...] [--split off] [--routes token,token]
    engine\.venv\Scripts\python.exe scripts\compare_drill_routes.py --probe renewal
--split off drills each cell without splitting it, so a piece over one getData page (config.DRILL_PAGE) is paged by
offset (the paging probe); --routes compares any two routes, e.g. the token route against itself (EPM's offset pages
have no stable order, so a route's own run-to-run difference is the baseline). --probe renewal runs the token route
alone: cell 1, cell 2, then cell 2 again with the session cookies dropped just before its second getData (EPM then
treats the key in flight as lost, as when its server session ends), one getData with the old key on the new session,
and cell 1 once more.

Per cell one line: "<cell>: lines N/N, same lines: True, tie browser/token: exact/exact, cube value equal: True,
seconds B/T" ("same lines": the rows as sorted lists of json.dumps(row, sort_keys=True), types kept, and the columns
are equal; the cube values agree within 0.005), then per route what it took (pieces, too-wide refusals, derived
siblings, the token route's key requests, renewals, requests by path and status, getData offsets, anything epmhttp
refused - by its kind). Before them, the view definition of the first cell on both routes, diffed key by key.
Exit code 0: every cell and the view's columns the same; 1: a difference, a failed route or a token route that
started a browser (--probe: the probe failed); 2: a route could not sign in (BLOCKED: nothing is retried, no
sign-in is attempted).

Prints and keeps no figure and no business text: counts, booleans, seconds, key paths, value types and query item
names. What crosses from a route's process to this one, through a temporary folder deleted at the end, is hashes:
of the rows, of the view's leaf values and of the drill's columns, with the view's key paths and value types and the
rows' type names per column. Read-only towards EPM: every request goes through the engine's own routes (the probe's
one extra request is a getData read with EPM's own query). Each route's process closes the hidden browser it started
(drill_browser.shutdown_shared) and only that one; the PC's remembered too-wide cells are not read or written (each
route meets EPM's refusal itself)."""
from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import time
from collections import Counter
from pathlib import Path
from urllib.parse import unquote_plus, urlsplit

from smartass.epm import epmhttp
from smartass.epm.drillcore import redact

CELLS = (
    {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"},   # one piece
    {"All Accounts": "ac.C60", "All Entities": "un.SY", "All Time Periods": "Sep 26"},        # over 5 000: split
    {"All Accounts": "ac.C60080", "All Entities": "un.SY", "All Time Periods": "Yr 26"},      # too wide: derived
    {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Q3 26"},    # over 5 000 in one
)                                                   # query: split by the line count (cell 2 is now refused as too wide)
ROUTES = ("browser", "token")
CELL_S = 1500                       # one cell's budget inside a route's process (it closes its browser after)
BLOCKED = 2
TRAP_AT = 2                         # the renewal probe drops the session cookies before this getData of cell 2
_PRINTABLE = re.compile(r"[\x21-\x7e]*")


# ---------------------------------------------------------------- the judge (pure: tests/test_compare_drill_routes.py)
def _typed(o) -> dict:
    return {"$type": type(o).__name__, "$value": str(o)}


def _hash(value) -> str:
    return hashlib.sha256(json.dumps(value, sort_keys=True, default=_typed, ensure_ascii=False).encode("utf-8")) \
        .hexdigest()


def row_keys(rows: list[dict]) -> list[str]:
    """The rows as a sorted list of json.dumps(row, sort_keys=True), each hashed so no line leaves its process; a
    value JSON has no type for (a Decimal, a date) keeps its type name, so it never equals a string."""
    return sorted(_hash(r) for r in rows)


def column_types(rows: list[dict]) -> dict[str, list[str]]:
    """The type names of each column's values."""
    out: dict[str, set] = {}
    for r in rows:
        for k, v in r.items():
            out.setdefault(k, set()).add(type(v).__name__)
    return {k: sorted(v) for k, v in out.items()}


def _close(a, b) -> bool:
    if a is None or b is None:
        return a is None and b is None
    return abs(float(a) - float(b)) < 0.005


def compare(a: dict, b: dict) -> dict:
    """One cell on two routes: same lines (rows and columns), the same value types, the same tie, the same cube
    value; how the rows differ as counts (only on one side, paged twice on one side)."""
    both = bool(a.get("ok") and b.get("ok"))
    ra, rb = (Counter(s["rows"]) for s in (a, b)) if both else (Counter(), Counter())
    ta, tb = (a.get("types") or {}, b.get("types") or {}) if both else ({}, {})
    types = sorted(f"{c} ({', '.join(ta.get(c, []))}/{', '.join(tb.get(c, []))})"
                   for c in ta.keys() | tb.keys() if ta.get(c) != tb.get(c))
    same_rows = both and ra == rb
    same_columns = both and a["columns"] == b["columns"]
    same = bool(same_rows and same_columns)
    tie = both and a["status"] == b["status"]
    cube = both and _close(a["cell_value"], b["cell_value"])
    return {"a": a, "b": b, "same_lines": same, "same_rows": bool(same_rows), "same_columns": bool(same_columns),
            "tie_equal": bool(tie), "cube_equal": bool(cube), "type_differs": types,
            "only_a": sum((ra - rb).values()), "only_b": sum((rb - ra).values()),
            "dupes_a": sum(ra.values()) - len(ra), "dupes_b": sum(rb.values()) - len(rb),
            "ok": bool(same and tie and cube and not types)}


def label(cell: dict) -> str:
    return " / ".join(cell.values())


def line(cell: dict, c: dict, names: tuple[str, str] = ROUTES) -> str:
    """The brief's line: counts, booleans, the tie and seconds only; a failed route adds its error kind."""
    a, b = c["a"], c["b"]

    def n(s):
        return str(len(s["rows"])) if s.get("ok") else "-"

    def secs(s):
        return f"{s['secs']:.1f}" if s.get("ok") else "-"
    text = (f"{label(cell)}: lines {n(a)}/{n(b)}, same lines: {c['same_lines']}, tie {names[0]}/{names[1]}: "
            f"{a.get('status', '-')}/{b.get('status', '-')}, cube value equal: {c['cube_equal']}, "
            f"seconds {secs(a)}/{secs(b)}")
    for name, s in zip(names, (a, b)):
        if not s.get("ok"):
            text += f"; {name} failed: {s.get('error', 'not run')}"
    return text


def _walk(doc, path: str = "$", out: dict | None = None) -> dict:
    """path -> (type name, the value of a leaf) of every node; a string holding JSON (the view's viewDefinition) is
    walked into as '<json>'."""
    out = {} if out is None else out
    if isinstance(doc, str) and doc[:1] in "{[":
        try:
            inner = json.loads(doc)
        except ValueError:
            inner = None
        if isinstance(inner, (dict, list)):
            out[path] = ("json str", None)
            return _walk(inner, path + "<json>", out)
    out[path] = (type(doc).__name__, None if isinstance(doc, (dict, list)) else doc)
    if isinstance(doc, dict):
        for k, v in doc.items():
            _walk(v, f"{path}.{k}", out)
    elif isinstance(doc, list):
        for i, v in enumerate(doc):
            _walk(v, f"{path}[{i}]", out)
    return out


def _items(query: str) -> list[tuple[str, str, str]]:
    """(decoded name, raw name, raw value) of every query item."""
    out = []
    for item in (p for p in query.split("&") if p):
        k, _, v = item.partition("=")
        out.append((unquote_plus(k), k, v))
    return out


def _link(value) -> dict | None:
    """A link leaf by its parts, hashed: where it points (host and path) and each query item's value."""
    if not (isinstance(value, str) and value.startswith(("http://", "https://", "/"))):
        return None
    u = urlsplit(value)
    return {"where": _hash([u.netloc, u.path]), "query": {n: _hash(v) for n, _, v in _items(u.query)}}


def _columns(view) -> list | None:
    from smartass.epm import drill_lines
    try:
        return drill_lines.columns_from_view(view)
    except Exception:   # noqa: BLE001 - not a drill view
        return None


def view_digest(view) -> dict:
    """What crosses between the processes of a view definition: its key paths with their types and hashed leaf
    values (a link by its hashed parts), and the drill's columns as a count and a hash. No value in plain text."""
    cols = _columns(view)
    return {"paths": {p: [t, None if v is None else _hash(v), _link(v)] for p, (t, v) in _walk(view).items()},
            "columns": None if cols is None else len(cols), "columns_hash": None if cols is None else _hash(cols)}


def view_diff(a: dict, b: dict) -> dict:
    """Two view digests key by key: paths only in one, paths whose type differs, leaf paths whose value differs
    (a link: by part). Paths, type names and query item names only."""
    wa, wb = a["paths"], b["paths"]
    both = wa.keys() & wb.keys()

    def note(p):
        la, lb = wa[p][2], wb[p][2]
        if not la or not lb:
            return ""
        qa, qb = la["query"], lb["query"]
        names = ", ".join(sorted(n for n in qa.keys() | qb.keys() if qa.get(n) != qb.get(n))) or "none"
        return f" (a link: same path {la['where'] == lb['where']}, query items that differ: {names})"
    return {"only_browser": sorted(wa.keys() - wb.keys()), "only_token": sorted(wb.keys() - wa.keys()),
            "type_differs": sorted(f"{p} ({wa[p][0]}/{wb[p][0]})" for p in both if wa[p][0] != wb[p][0]),
            "value_differs": sorted(p + note(p) for p in both if wa[p][0] == wb[p][0] and wa[p][1] != wb[p][1])}


def view_lines(a: dict, b: dict, cell: dict | None = None) -> tuple[list[str], bool]:
    """Step 1's report and whether it passes: the diff's size, the columns the drill takes from each view (a view
    the drill cannot read is a failure), then the differing paths with list indices folded ([*]) and counted."""
    d = view_diff(a, b)
    ca, cb = a["columns"], b["columns"]
    same = ca is not None and cb is not None and a["columns_hash"] == b["columns_hash"]
    out = [f"view of {label(cell) if cell else 'the cell'}: {len(a['paths'])} paths on one route, "
           f"{len(b['paths'])} on the other; same documents: {not any(d.values())}",
           f"  columns the drill takes from it: {'-' if ca is None else ca}/{'-' if cb is None else cb}, same labels "
           f"and fields in the same order: {same}" + (" (not a drill view)" if ca is None or cb is None else "")]
    for kind, paths in d.items():
        if paths:
            folded = Counter(re.sub(r"\[\d+\]", "[*]", p) for p in paths)
            shown = ", ".join(f"{p} x{n}" if n > 1 else p for p, n in sorted(folded.items()))
            out.append(f"  {kind.replace('_', ' ')} ({len(paths)}): {shown[:1500]}")
    return out, same


def refused_kind(message: str, url: str, body: str | None) -> str:
    """What epmhttp refused, by its kind: the unlisted query items' names, the names of items with a character a URL
    cannot carry, or the first keyword of a drill query - never a value."""
    message = redact(message)
    parts = urlsplit(url)
    items = _items(parts.query)
    if "query item this request does not take" in message:
        rule = epmhttp.ALLOWED.get(parts.path)
        extra = sorted({n for n, _, _ in items if rule is None or n not in rule.query})
        return f"{message} (unlisted: {', '.join(extra)})"
    if "non-ASCII" in message:
        bad = sorted({n for n, k, v in items if not _PRINTABLE.fullmatch(k + v)})
        where = f"the query items {', '.join(bad)}" if bad else "the path"
        return f"{message} (in {where})"
    if "read-only SELECT" in message:
        try:
            query = json.loads(body or "")["QUERY"]
            word = re.match(r"\s*([A-Za-z]+)", query).group(1).upper()[:12]
        except (ValueError, KeyError, TypeError, AttributeError):
            word = "?"
        return f"{message} (the query starts with {word})"
    return message


def error_kind(e: BaseException) -> str:
    """An error as its class and a fixed kind (EPM's refusals with their EPMFDM codes): never its text, which can
    quote a member, a supplier or an amount."""
    from smartass.epm import oauth
    from smartass.epm.drillcore import DrillBusy, SignedOut, TooManyLines
    text = str(e)
    codes = list(dict.fromkeys(re.findall(r"EPMFDM-\d+", text)))
    if isinstance(e, oauth.OAuthError) or "Sign-in to EPM did not complete" in text:
        kind = "sign-in"
    elif isinstance(e, DrillBusy) or "another fin_drill is running" in text:
        kind = "another drill holds the browser profile"
    elif isinstance(e, epmhttp.Refused):
        kind = "refused by the session's allowlist"
    elif isinstance(e, SignedOut):
        kind = "the EPM session was lost"
    elif codes or "EPM refused" in text or "too wide" in text:
        kind = "EPM refused" + (f" ({', '.join(codes)})" if codes else "")
    elif isinstance(e, TooManyLines) or ("more than" in text and "lines" in text):
        kind = "too many lines"
    elif isinstance(e, (TimeoutError, asyncio.TimeoutError)):
        kind = "timeout"
    elif "this script never uses Excel" in text:
        kind = "Excel asked (refused)"
    else:
        kind = "other"
    return f"{type(e).__name__}: {kind}"


def detail(route: str, s: dict) -> str:
    """What a route took for one cell: counts only."""
    k = s.get("counts") or {}
    parts = [f"pieces {s.get('pieces', '-')}", f"too wide {k.get('too_wide', 0)}",
             f"too many lines {k.get('too_many', 0)}",
             f"derived used/tried {k.get('derived_used', 0)}/{k.get('derived_tried', 0)}",
             f"result warnings {s.get('warnings', '-')}", f"log warnings {k.get('log_warnings', 0)}",
             f"wall {s.get('wall', 0):.1f} s"]
    if route == "browser":
        parts.append(f"page route {k.get('page_route', 0)}")
    if k.get("excel"):
        parts.append(f"Excel asked {k['excel']} (refused)")
    if route == "token":
        reqs = ", ".join(f"{name[4:]} x{n}" for name, n in sorted(k.items()) if name.startswith("req "))
        offs = ", ".join(f"{o} x{n}" if n > 1 else o for o, n in sorted((s.get("offsets") or {}).items(),
                                                                         key=lambda x: int(x[0])))
        parts += [f"key requests {s.get('key_requests', '-')}", f"renewals {k.get('renewals', 0)}",
                  f"requests: {reqs or '-'}", f"getData offsets: {offs or '-'}", f"refused {k.get('refused', 0)}"]
    text = f"  {route}: " + ", ".join(parts)
    for r in dict.fromkeys(s.get("refusals") or []):
        text += f"\n    refused: {r}"
    return text


def rows_detail(c: dict, names: tuple[str, str]) -> str | None:
    """How two routes' rows differ, as counts and column names; None when they do not."""
    if not (c["a"].get("ok") and c["b"].get("ok")) or (c["same_rows"] and not c["type_differs"]):
        return None
    da, db = (len(set(s["rows"])) for s in (c["a"], c["b"]))
    text = (f"    rows: only on {names[0]} {c['only_a']}, only on {names[1]} {c['only_b']}, paged twice "
            f"{c['dupes_a']}/{c['dupes_b']}, distinct {da}/{db}")
    if c["type_differs"]:
        text += "; value types differ: " + "; ".join(c["type_differs"])
    return text


def old_key_timing(requests: list[dict], old: int | None, renew_at: float | None,
                   new_key_at: float | None) -> dict:
    """The requests that carried the old key, by when they left: before the renewal began (in flight then, or already
    answered), between its start and the new key, after the new key."""
    inf = float("inf")
    r0, r1 = inf if renew_at is None else renew_at, inf if new_key_at is None else new_key_at
    out = {"before": 0, "between": 0, "after": 0}
    for q in requests:
        if old is not None and q.get("gen") == old:
            out["before" if q["t"] < r0 else "between" if q["t"] < r1 else "after"] += 1
    return out


def renewal_passed(p: dict) -> bool:
    """The renewal probe passes with a renewal, the baseline's lines, an exact tie, no old key sent after the new one
    was taken, and a later drill in the same process that still works and gives its earlier lines."""
    return bool(p.get("ok") and p.get("renewals", 0) >= 1 and p.get("same_rows") and p.get("tie") == "exact"
                and p.get("cube_equal") and p.get("old_after_new_key") == 0 and p.get("later_ok")
                and p.get("later_same_rows"))


# ---------------------------------------------------------------- one route's process
def _blocked(e: BaseException) -> bool:
    from smartass.epm import oauth
    from smartass.epm.drillcore import DrillBusy
    text = str(e)
    return (isinstance(e, (oauth.OAuthError, DrillBusy)) or "Sign-in to EPM did not complete" in text
            or "another fin_drill is running" in text)


class _Route:
    """This process's route with the engine patched to count (and the token route's session wrapped): one drill at
    a time, what it took in self.now."""

    def __init__(self, route: str, split: str):
        import logging

        from smartass.core import config
        from smartass.epm import drill_split, drill_token, drillstep
        from smartass.epm.drillcore import CellTooWide, TooManyLines
        if config.DRILL_ROUTE != route:
            raise RuntimeError(f"SV_DRILL_ROUTE is {config.DRILL_ROUTE!r}, not {route!r}")
        self.split, self.drillstep = split, drillstep
        self.now: dict = {}
        self.gens: dict[str, int] = {}             # key -> its number (keys stay in this process, never written)
        self._lock = threading.Lock()
        now = self.now
        route_self = self

        class _Count(logging.Handler):          # every warning counted; none shown (they may name a cell's text)
            def emit(self, record):
                route_self.now["counts"]["log_warnings"] += 1
        logging.getLogger().addHandler(_Count(logging.WARNING))

        async def no_excel(fn, what="call"):
            route_self.now["counts"]["excel"] += 1
            raise RuntimeError(f"this script never uses Excel ({what})")
        drillstep._call = no_excel

        too_wide: set = set()                   # this process only: each route meets EPM's refusal itself
        drill_split.known_too_wide = lambda: too_wide
        drill_split.remember_too_wide = lambda key: too_wide.add(key)

        real_fetch, real_derived, real_page, real_lines_of = (drillstep._drill_fetch, drillstep._derived_piece,
                                                              drillstep._page_fetch, drillstep._lines_of)

        async def drill_fetch(url, max_lines):
            try:
                return await real_fetch(url, max_lines)
            except CellTooWide:
                now["counts"]["too_wide"] += 1
                raise
            except TooManyLines:
                now["counts"]["too_many"] += 1
                raise

        async def derived_piece(cell, cube, template, body, state):
            now["counts"]["derived_tried"] += 1
            piece = await real_derived(cell, cube, template, body, state)
            now["counts"]["derived_used"] += piece is not None
            return piece

        async def page_fetch(url, max_lines):
            now["counts"]["page_route"] += 1
            return await real_page(url, max_lines)

        def lines_of(pieces, view, parent_cube=None):
            now["view"] = view
            return real_lines_of(pieces, view, parent_cube)
        drillstep._drill_fetch = drill_fetch
        drillstep._derived_piece = derived_piece
        drillstep._page_fetch = page_fetch
        drillstep._lines_of = lines_of

        self.td = drill_token.shared() if route == "token" else None
        if self.td is not None:
            self._wrap_token(self.td)

    def _wrap_token(self, td) -> None:
        now, real_request, real_renew, real_take = self.now, td.http.request, td._renew, td._take_key

        def request(method, url, headers=None, body=None, timeout=120.0):
            parts = urlsplit(url)
            name = parts.path.rsplit("/", 1)[-1]
            key = (headers or {}).get("x-appbuilder-client-id")
            fire = False
            with self._lock:
                gen = None if key is None else self.gens.setdefault(key, len(self.gens))
                if name == "getData":
                    off = next((v for n, _, v in _items(parts.query) if n == "offset"), "?")
                    now["offsets"][off] += 1
                    trap = now.get("trap")
                    if trap is not None:
                        trap["n"] += 1
                        fire = trap["n"] == trap["at"]
                        if fire:
                            now["trap"] = None
            if fire:                            # the server session is gone: EPM sees a request without its cookie
                td.http._jar.clear()
                now["fired"] = True
            rec = {"gen": gen, "t": time.monotonic(), "name": name}
            now["requests"].append(rec)
            try:
                r = real_request(method, url, headers=headers, body=body, timeout=timeout)
            except epmhttp.Refused as e:
                now["counts"]["refused"] += 1
                now["refusals"].append(refused_kind(str(e), url, body))
                raise
            except OSError as e:
                now["counts"][f"req {name} lost ({type(e).__name__})"] += 1
                raise
            rec["status"], rec["t_end"] = r.status, time.monotonic()
            now["counts"][f"req {name} {r.status}"] += 1
            if name == "getData" and body:
                now["last_body"] = body         # EPM's own query, kept in memory for the probe's old-key request
            return r

        async def renew(used):
            now["counts"]["renewals"] += 1
            if td._key == used and now.get("renew_at") is None:
                now["renew_at"] = time.monotonic()
            return await real_renew(used)

        async def take_key():
            key = await real_take()
            now["keys_at"].append(time.monotonic())
            return key
        td.http.request, td._renew, td._take_key = request, renew, take_key

    async def drill(self, i: int) -> dict:
        """Drill cell i (1-based) as fin_drill does; the entry the comparison needs."""
        keep = {k: self.now.get(k) for k in ("trap", "last_body")}
        self.now.update(counts=Counter(), view=None, refusals=[], offsets=Counter(), requests=[], renew_at=None,
                        keys_at=[], fired=False, **keep)
        taken, t0 = self.td._taken if self.td else 0, time.monotonic()
        try:
            res = await asyncio.wait_for(self.drillstep.drill(dict(CELLS[i - 1]), say=lambda text: None,
                                                              split=self.split), CELL_S)
            got = {"ok": True, "rows": row_keys(res.rows), "types": column_types(res.rows),
                   "columns": list(res.columns), "status": res.reconciliation.get("status"),
                   "cell_value": res.cell_value, "secs": res.secs, "pieces": len(res.pieces) or 1,
                   "warnings": len(res.warnings)}
        except Exception as e:   # noqa: BLE001 - reported by its kind; a sign-in failure stops the run
            got = {"ok": False, "error": error_kind(e), "blocked": _blocked(e)}
        got.update(wall=round(time.monotonic() - t0, 1), counts=dict(self.now["counts"]),
                   view=None if self.now["view"] is None else view_digest(self.now["view"]),
                   refusals=self.now["refusals"], offsets=dict(self.now["offsets"]),
                   key_requests=(self.td._taken - taken) if self.td else None)
        return got

    async def old_key_request(self, key: str) -> int:
        """One getData read (EPM's own query, one line) sent with a key that was found gone: its HTTP status."""
        from smartass.epm import drill_token
        try:
            r = await asyncio.to_thread(self.td.http.request, "POST", f"{drill_token.GET_DATA}?limit=1&offset=0",
                                        headers={**drill_token._HEADERS, "x-appbuilder-client-id": key},
                                        body=self.now["last_body"], timeout=120.0)
            return r.status
        except OSError:
            return 0


async def _cells(route: str, picks: list[int], split: str) -> dict:
    r = _Route(route, split)
    cells = {}
    for i in picks:
        cells[str(i)] = got = await r.drill(i)
        if got.get("blocked"):
            break
    return cells


async def _renewal_probe(split: str) -> dict:
    """Cell 1, cell 2, cell 2 with the session lost under its getData requests, an old-key getData on the new
    session, cell 1 again: counts, booleans and timings of the old key only."""
    r = _Route("token", split)
    td = r.td
    first = await r.drill(1)
    base = await r.drill(2)
    p = {"cell1": first, "base": base}
    if not (first["ok"] and base["ok"]):
        return {**p, "ok": False, "blocked": first.get("blocked") or base.get("blocked")}
    old = td._key
    r.now["trap"] = {"n": 0, "at": TRAP_AT}
    hit = await r.drill(2)
    renew_at, keys_at, fired = r.now["renew_at"], r.now["keys_at"], r.now["fired"]
    new_key_at = next((t for t in keys_at if renew_at is not None and t >= renew_at), None)
    timing = old_key_timing(r.now["requests"], r.gens.get(old), renew_at, new_key_at)
    late = sum(1 for q in r.now["requests"] if q["gen"] == r.gens.get(old) and new_key_at is not None
               and q.get("t_end", 0) >= new_key_at)
    statuses = Counter(f"{q['name']} {q.get('status', 'lost')}" for q in r.now["requests"])
    old_status = await r.old_key_request(old)
    del old
    later = await r.drill(1)
    c, cl = compare(base, hit), compare(first, later)
    return {**p, "hit": hit, "later": later, "fired": fired, "ok": hit["ok"] and fired,
            "renewals": hit["counts"].get("renewals", 0), "key_requests": hit["key_requests"],
            "same_rows": c["same_lines"] and not c["type_differs"], "tie": hit.get("status"),
            "cube_equal": c["cube_equal"], "statuses": dict(statuses), "old_before": timing["before"],
            "old_between": timing["between"], "old_after_new_key": timing["after"], "old_answered_after_new_key": late,
            "old_key_status": old_status,
            "later_ok": later["ok"], "later_same_rows": cl["same_lines"] and not cl["type_differs"],
            "later_renewals": later["counts"].get("renewals", 0), "later_key_requests": later["key_requests"]}


def run_route(route: str, picks: list[int], out: Path, split: str = "auto", probe: str | None = None) -> int:
    """A route's process: its cells (or the probe) into out (JSON), then its own hidden browser closed. Never prints
    a result."""
    from smartass.epm import drill_browser
    data: dict = {"route": route, "cells": {}}
    try:
        if probe == "renewal":
            data["probe"] = asyncio.run(_renewal_probe(split))
        else:
            data["cells"] = asyncio.run(_cells(route, picks, split))
    except Exception as e:   # noqa: BLE001 - the parent reports it by its kind
        data["error"], data["blocked"] = error_kind(e), _blocked(e)
    finally:
        data["browser_started"] = drill_browser._shared is not None
        drill_browser.shutdown_shared()
        out.write_text(json.dumps(data), encoding="utf-8")
    return 0


# ---------------------------------------------------------------- the run
def _last_error_type(stderr: str) -> str:
    names = re.findall(r"(?m)^([A-Za-z_][\w.]*(?:Error|Exception|Interrupt|Exit))\b", stderr or "")
    return names[-1] if names else "none"


def _blocked_message(name: str, data: dict) -> str | None:
    if data.get("blocked"):
        return f"{name} route: {data['error']}"
    for i, s in data.get("cells", {}).items():
        if s.get("blocked"):
            return f"{name} route, cell {i}: {s['error']}"
    if (data.get("probe") or {}).get("blocked"):
        return f"{name} route: the renewal probe's first drills"
    return None


def report(picks: list[int], got: list[dict], names: tuple[str, str] = ROUTES) -> int:
    """Step 1 and Step 3 for two routes' results; 0 when the view's columns and every cell are the same."""
    a, b = got
    first = str(picks[0])
    va, vb = a["cells"].get(first, {}).get("view"), b["cells"].get(first, {}).get("view")
    print(f"Step 1 - the view definition ({names[0]}/{names[1]}):")
    if va is not None and vb is not None:
        lines, ok = view_lines(va, vb, CELLS[picks[0] - 1])
        for text in lines:
            print("  " + text)
    else:
        print(f"  no view to compare ({names[0]}: {va is not None}, {names[1]}: {vb is not None})")
        ok = False
    print("Step 3 - the lines:")
    for i in picks:
        c = compare(a["cells"].get(str(i), {"ok": False}), b["cells"].get(str(i), {"ok": False}))
        print("  " + line(CELLS[i - 1], c, names))
        for name, side in zip(names, got):
            if str(i) in side["cells"]:
                print("  " + detail(name, side["cells"][str(i)]))
        rd = rows_detail(c, names)
        if rd:
            print(rd)
        ok &= c["ok"]
    for name, side in zip(names, got):
        if side.get("error"):
            print(f"  the {name} route's process failed: {side['error']}")
            ok = False
    for side in got:
        if side.get("route") == "token":
            print(f"  the token route started a browser: {side['browser_started']}")
            ok &= not side["browser_started"]
    print("PASS: the routes give the same lines" if ok else "FAIL: see the lines above")
    return 0 if ok else 1


def report_probe(data: dict) -> int:
    p = data.get("probe") or {}
    print(f"Renewal probe (token route, one process; cell 2's session cookies dropped before its getData "
          f"#{TRAP_AT}):")
    if data.get("error"):
        print(f"  the process failed: {data['error']}")
        return 1

    def short(s):
        if not s.get("ok"):
            return f"failed: {s.get('error')}"
        return (f"lines {len(s['rows'])}, tie {s['status']}, key requests {s['key_requests']}, "
                f"renewals {s['counts'].get('renewals', 0)}, wall {s['wall']:.1f} s")
    for key, what in (("cell1", f"cell 1 ({label(CELLS[0])})"), ("base", f"cell 2 ({label(CELLS[1])}), baseline")):
        if key in p:
            print(f"  {what}: {short(p[key])}")
    if "hit" not in p:
        return 1
    st = ", ".join(f"{k} x{n}" for k, n in sorted(p["statuses"].items()))
    print(f"  cell 2 with the session lost (cookies dropped: {p['fired']}): {short(p['hit'])}; same rows as the "
          f"baseline: {p['same_rows']}, "
          f"cube value equal: {p['cube_equal']}")
    print(f"    requests: {st}")
    print(f"    old-key requests: sent before the renewal began {p['old_before']}, sent between its start and the "
          f"new key {p['old_between']}, sent after the new key {p['old_after_new_key']}; answered after the new key "
          f"{p['old_answered_after_new_key']}")
    print(f"  an old-key getData sent on the new session: HTTP {p['old_key_status']}")
    print(f"  cell 1 again: {short(p['later'])}; same rows as before: {p['later_same_rows']} -> the old-key request "
          f"{'signed the new session out' if p['later_renewals'] else 'did not sign the new session out'}")
    print(f"  the token route started a browser: {data['browser_started']}")
    ok = renewal_passed(p) and not data["browser_started"]
    print("PASS: the drill renewed its key under load and kept working" if ok else "FAIL: see the lines above")
    return 0 if ok else 1


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--cell", type=int, action="append", choices=range(1, len(CELLS) + 1),
                    help=f"a cell to compare (1-{len(CELLS)}; default all)")
    ap.add_argument("--split", default="auto", help="drillstep's split: 'auto' (default) or e.g. 'off' (page)")
    ap.add_argument("--routes", default=",".join(ROUTES), help="the two routes to compare (default browser,token)")
    ap.add_argument("--probe", choices=("renewal",), help="the forced-renewal probe on the token route")
    ap.add_argument("--route", choices=ROUTES, help=argparse.SUPPRESS)        # a route's own process
    ap.add_argument("--out", type=Path, help=argparse.SUPPRESS)
    args = ap.parse_args(argv)
    picks = sorted(set(args.cell or range(1, len(CELLS) + 1)))
    if args.route:
        return run_route(args.route, picks, args.out, args.split, args.probe)
    names = ("token",) if args.probe else tuple(args.routes.split(","))
    if args.probe is None and (len(names) != 2 or not set(names) <= set(ROUTES)):
        ap.error("--routes takes two of: " + ", ".join(ROUTES))

    from smartass import __version__
    from smartass.core import config
    from smartass.epm import oauth
    print(f"compare_drill_routes (smartass {__version__}): "
          + (f"probe {args.probe}" if args.probe else f"cells {picks}, split {args.split!r}; routes "
             + " then ".join(names) + ", one process each"), flush=True)
    try:
        oauth.access_token(config.EPM_BASE)
    except oauth.OAuthError as e:
        print(f"BLOCKED: the token route has no sign-in ({error_kind(e)}); run 'smartass login' first")
        return BLOCKED
    got = []
    with tempfile.TemporaryDirectory(prefix="smartass-compare-") as tmp:
        for n, route in enumerate(names):
            out = Path(tmp) / f"side{n}.json"
            t0 = time.monotonic()
            argv_child = [sys.executable, str(Path(__file__).resolve()), "--route", route, "--out", str(out),
                          "--split", args.split, *[f"--cell={i}" for i in picks]]
            if args.probe:
                argv_child += ["--probe", args.probe]
            try:
                p = subprocess.run(argv_child, env={**os.environ, "SV_DRILL_ROUTE": route}, capture_output=True,
                                   text=True, encoding="utf-8", errors="replace",
                                   timeout=(5 if args.probe else len(picks)) * CELL_S + 600)
            except subprocess.TimeoutExpired:
                print(f"{route} route: no end after {time.monotonic() - t0:.0f} s; stopped")
                return 1
            print(f"{route} route: its process ended (exit code {p.returncode}) after {time.monotonic() - t0:.0f} s",
                  flush=True)
            if not out.exists():
                print(f"  no result (its last error: {_last_error_type(p.stderr)})")
                return 1
            got.append(json.loads(out.read_text(encoding="utf-8")))
            why = _blocked_message(route, got[-1])
            if why:
                print(f"BLOCKED: {why} - an interactive sign-in is needed; nothing was retried")
                return BLOCKED
        if args.probe:
            return report_probe(got[0])
        return report(picks, got, names)


if __name__ == "__main__":
    sys.exit(main())
