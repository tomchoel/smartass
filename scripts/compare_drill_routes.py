r"""Live gate of the drill over the token (plan 2026-11-19, Task 7): the same FinLite cells drilled on the browser
route (SV_DRILL_ROUTE=browser: drill_direct's warm tab, the page route behind it) and on the token route
(SV_DRILL_ROUTE=token: drill_token over epmhttp), one fresh Python process per route because config reads the route
at import, and compared line for line. The browser route is deleted (Task 9) only after this run passes.

Usage (from the repo root, signed in on plan1; no Excel is used: drillstep's Excel fallbacks raise here):
    engine\.venv\Scripts\python.exe scripts\compare_drill_routes.py [--cell 1 ...]

Per cell one line: "<cell>: lines N/N, same lines: True, tie browser/token: exact/exact, cube value equal: True,
seconds B/T" ("same lines": the rows as sorted lists of json.dumps(row, sort_keys=True) and the columns are equal;
the cube values agree within 0.005), then per route what it took (pieces, too-wide refusals, derived siblings, the
token route's key requests, renewals and requests by path and status, anything epmhttp refused - by its kind). Before
them, the view definition of the first cell on both routes, diffed key by key (key paths and value types only).
Exit code 0: every cell the same; 1: a cell differs, a route failed or the token route started a browser; 2: a route
could not sign in (BLOCKED: nothing is retried, no sign-in is attempted).

Prints and keeps no figure and no business text: counts, booleans, seconds, key paths and value types. The rows
cross from a route's process to this one as hashes in a temporary folder that is deleted at the end. Read-only
towards EPM: every request goes through the engine's own routes. Each route's process closes the hidden browser it
started (drill_browser.shutdown_shared) and only that one; the PC's remembered too-wide cells are not read or
written (each route meets EPM's refusal itself)."""
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
)
ROUTES = ("browser", "token")
CELL_S = 1500                       # one cell's budget inside a route's process (it closes its browser after)
BLOCKED = 2
_PRINTABLE = re.compile(r"[\x21-\x7e]*")


# ---------------------------------------------------------------- the judge (pure: tests/test_compare_drill_routes.py)
def row_keys(rows: list[dict]) -> list[str]:
    """The rows as a sorted list of json.dumps(row, sort_keys=True), each hashed so no line leaves its process."""
    return sorted(hashlib.sha256(json.dumps(r, sort_keys=True, default=str, ensure_ascii=False).encode("utf-8"))
                  .hexdigest() for r in rows)


def _close(a, b) -> bool:
    if a is None or b is None:
        return a is None and b is None
    return abs(float(a) - float(b)) < 0.005


def compare(b: dict, t: dict) -> dict:
    """One cell's browser and token results: same lines (rows and columns), the same tie, the same cube value."""
    both = bool(b.get("ok") and t.get("ok"))
    same_rows = both and b["rows"] == t["rows"]
    same_columns = both and b["columns"] == t["columns"]
    same = bool(same_rows and same_columns)
    tie = both and b["status"] == t["status"]
    cube = both and _close(b["cell_value"], t["cell_value"])
    return {"browser": b, "token": t, "same_lines": same, "same_rows": bool(same_rows),
            "same_columns": bool(same_columns), "tie_equal": bool(tie), "cube_equal": bool(cube),
            "ok": bool(same and tie and cube)}


def label(cell: dict) -> str:
    return " / ".join(cell.values())


def line(cell: dict, c: dict) -> str:
    """The brief's line: counts, booleans, the tie and seconds only; a failed route adds its error kind."""
    b, t = c["browser"], c["token"]

    def n(s):
        return str(len(s["rows"])) if s.get("ok") else "-"

    def secs(s):
        return f"{s['secs']:.1f}" if s.get("ok") else "-"
    text = (f"{label(cell)}: lines {n(b)}/{n(t)}, same lines: {c['same_lines']}, tie browser/token: "
            f"{b.get('status', '-')}/{t.get('status', '-')}, cube value equal: {c['cube_equal']}, "
            f"seconds {secs(b)}/{secs(t)}")
    for route, s in (("browser", b), ("token", t)):
        if not s.get("ok"):
            text += f"; {route} failed: {s.get('error', 'not run')}"
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


def view_diff(a, b) -> dict:
    """Two view definitions key by key: paths only in one, paths whose type differs, leaf paths whose value differs.
    Paths and type names only - never a value."""
    wa, wb = _walk(a), _walk(b)
    both = wa.keys() & wb.keys()
    return {"only_browser": sorted(wa.keys() - wb.keys()), "only_token": sorted(wb.keys() - wa.keys()),
            "type_differs": sorted(f"{p} ({wa[p][0]}/{wb[p][0]})" for p in both if wa[p][0] != wb[p][0]),
            "value_differs": sorted(p for p in both if wa[p][0] == wb[p][0] and wa[p][1] != wb[p][1])}


def _columns(view) -> list | None:
    from smartass.epm import drill_lines
    try:
        return drill_lines.columns_from_view(view)
    except Exception:   # noqa: BLE001 - not a drill view
        return None


def view_lines(a, b, cell: dict | None = None) -> list[str]:
    """Step 1's report: the diff's size, the columns the drill takes from each view (count, same order), then the
    differing paths with list indices folded ([*]) and counted."""
    d = view_diff(a, b)
    ca, cb = _columns(a), _columns(b)
    out = [f"view of {label(cell) if cell else 'the cell'}: {len(_walk(a))} paths on the browser route, "
           f"{len(_walk(b))} on the token route; same documents: {not any(d.values())}",
           f"  columns the drill takes from it: {len(ca) if ca is not None else '-'}/"
           f"{len(cb) if cb is not None else '-'}, same labels and fields in the same order: {ca == cb}"]
    for kind, paths in d.items():
        if paths:
            folded = Counter(re.sub(r"\[\d+\]", "[*]", p) for p in paths)
            shown = ", ".join(f"{p} x{n}" if n > 1 else p for p, n in sorted(folded.items()))
            out.append(f"  {kind.replace('_', ' ')} ({len(paths)}): {shown[:1500]}")
    return out


def _items(query: str) -> list[tuple[str, str, str]]:
    """(decoded name, raw name, raw value) of every query item."""
    out = []
    for item in (p for p in query.split("&") if p):
        k, _, v = item.partition("=")
        out.append((unquote_plus(k), k, v))
    return out


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
    """An error as its type and a short text: no URL query, no amount, EPM's refusals by their EPMFDM code only."""
    text = " ".join(redact(str(e)).split())
    codes = re.findall(r"EPMFDM-\d+", text)
    if codes and ("EPM refused" in text or type(e).__name__ == "CellTooWide"):
        text = "EPM refused the drill: " + ", ".join(dict.fromkeys(codes))
    text = re.sub(r"(?<![\w.])-?\d+(?:[ ,. ]\d+)+|(?<![\w.])\d{4,}", "#", text)
    return f"{type(e).__name__}: {text[:160]}"


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
        parts += [f"key requests {s.get('key_requests', '-')}", f"renewals {k.get('renewals', 0)}",
                  f"requests: {reqs or '-'}", f"refused {k.get('refused', 0)}"]
    text = f"  {route}: " + ", ".join(parts)
    for r in dict.fromkeys(s.get("refusals") or []):
        text += f"\n    refused: {r}"
    return text


# ---------------------------------------------------------------- one route's process
def _blocked(e: BaseException) -> bool:
    from smartass.epm import oauth
    from smartass.epm.drillcore import DrillBusy
    text = str(e)
    return (isinstance(e, (oauth.OAuthError, DrillBusy)) or "Sign-in to EPM did not complete" in text
            or "another fin_drill is running" in text)


async def _route(route: str, picks: list[int]) -> dict:
    """Drill the picked cells on this process's route (drillstep.drill, as fin_drill does) and keep what the
    comparison needs: row hashes, columns, the tie, the cube value, seconds, and counts of what it took."""
    import logging

    from smartass.core import config
    from smartass.epm import drill_split, drill_token, drillstep
    from smartass.epm.drillcore import CellTooWide, TooManyLines
    if config.DRILL_ROUTE != route:
        raise RuntimeError(f"SV_DRILL_ROUTE is {config.DRILL_ROUTE!r}, not {route!r}")
    now: dict = {"counts": Counter(), "view": None, "refusals": []}

    class _Count(logging.Handler):              # every warning counted; none shown (they may name a cell's text)
        def emit(self, record):
            now["counts"]["log_warnings"] += 1
    logging.getLogger().addHandler(_Count(logging.WARNING))

    async def no_excel(fn, what="call"):
        now["counts"]["excel"] += 1
        raise RuntimeError(f"this script never uses Excel ({what})")
    drillstep._call = no_excel

    too_wide: set = set()                       # this process only: each route meets EPM's refusal itself
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

    td = drill_token.shared() if route == "token" else None
    if td is not None:
        real_request, real_renew = td.http.request, td._renew

        def request(method, url, headers=None, body=None, timeout=120.0):
            name = urlsplit(url).path.rsplit("/", 1)[-1]
            try:
                r = real_request(method, url, headers=headers, body=body, timeout=timeout)
            except epmhttp.Refused as e:
                now["counts"]["refused"] += 1
                now["refusals"].append(refused_kind(str(e), url, body))
                raise
            except OSError as e:
                now["counts"][f"req {name} lost ({type(e).__name__})"] += 1
                raise
            now["counts"][f"req {name} {r.status}"] += 1
            return r

        async def renew(used):
            now["counts"]["renewals"] += 1
            return await real_renew(used)
        td.http.request, td._renew = request, renew

    cells = {}
    for i in picks:
        now.update(counts=Counter(), view=None, refusals=[])
        taken, t0 = td._taken if td else 0, time.monotonic()
        try:
            res = await asyncio.wait_for(drillstep.drill(dict(CELLS[i - 1]), say=lambda text: None), CELL_S)
            got = {"ok": True, "rows": row_keys(res.rows), "columns": list(res.columns),
                   "status": res.reconciliation.get("status"), "cell_value": res.cell_value, "secs": res.secs,
                   "pieces": len(res.pieces) or 1, "warnings": len(res.warnings)}
        except Exception as e:   # noqa: BLE001 - reported by its kind; a sign-in failure stops the run
            got = {"ok": False, "error": error_kind(e), "blocked": _blocked(e)}
        got.update(wall=round(time.monotonic() - t0, 1), counts=dict(now["counts"]), view=now["view"],
                   refusals=now["refusals"], key_requests=(td._taken - taken) if td else None)
        cells[str(i)] = got
        if got.get("blocked"):
            break
    return cells


def run_route(route: str, picks: list[int], out: Path) -> int:
    """A route's process: its cells into out (JSON), then its own hidden browser closed. Never prints a result."""
    from smartass.epm import drill_browser
    data: dict = {"route": route, "cells": {}}
    try:
        data["cells"] = asyncio.run(_route(route, picks))
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


def _blocked_message(route: str, data: dict) -> str | None:
    if data.get("blocked"):
        return f"{route} route: {data['error']}"
    for i, s in data["cells"].items():
        if s.get("blocked"):
            return f"{route} route, cell {i}: {s['error']}"
    return None


def report(picks: list[int], got: dict) -> int:
    b, t = got["browser"], got["token"]
    first = str(picks[0])
    vb, vt = b["cells"].get(first, {}).get("view"), t["cells"].get(first, {}).get("view")
    print("Step 1 - the view definition:")
    if vb is not None and vt is not None:
        for text in view_lines(vb, vt, CELLS[picks[0] - 1]):
            print("  " + text)
    else:
        print(f"  no view to compare (browser: {vb is not None}, token: {vt is not None})")
    print("Step 3 - the lines:")
    ok = True
    for i in picks:
        c = compare(b["cells"].get(str(i), {"ok": False}), t["cells"].get(str(i), {"ok": False}))
        print("  " + line(CELLS[i - 1], c))
        for route in ROUTES:
            if str(i) in got[route]["cells"]:
                print("  " + detail(route, got[route]["cells"][str(i)]))
        ok &= c["ok"]
    for route in ROUTES:
        if got[route].get("error"):
            print(f"  the {route} route's process failed: {got[route]['error']}")
            ok = False
    print(f"  the token route started a browser: {t['browser_started']}")
    ok &= not t["browser_started"]
    print("PASS: the token route gives the browser route's lines" if ok else "FAIL: see the lines above")
    return 0 if ok else 1


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--cell", type=int, action="append", choices=range(1, len(CELLS) + 1),
                    help="a cell to compare (1-3; default all)")
    ap.add_argument("--route", choices=ROUTES, help=argparse.SUPPRESS)        # a route's own process
    ap.add_argument("--out", type=Path, help=argparse.SUPPRESS)
    args = ap.parse_args(argv)
    picks = sorted(set(args.cell or range(1, len(CELLS) + 1)))
    if args.route:
        return run_route(args.route, picks, args.out)

    from smartass import __version__
    from smartass.core import config
    from smartass.epm import oauth
    print(f"compare_drill_routes (smartass {__version__}): cells {picks}; the browser route, then the token route, "
          "one process each", flush=True)
    try:
        oauth.access_token(config.EPM_BASE)
    except oauth.OAuthError as e:
        print(f"BLOCKED: the token route has no sign-in ({error_kind(e)}); run 'smartass login' first")
        return BLOCKED
    got = {}
    with tempfile.TemporaryDirectory(prefix="smartass-compare-") as tmp:
        for route in ROUTES:
            out = Path(tmp) / f"{route}.json"
            t0 = time.monotonic()
            try:
                p = subprocess.run([sys.executable, str(Path(__file__).resolve()), "--route", route, "--out", str(out),
                                    *[f"--cell={i}" for i in picks]],
                                   env={**os.environ, "SV_DRILL_ROUTE": route}, capture_output=True, text=True,
                                   encoding="utf-8", errors="replace", timeout=len(picks) * CELL_S + 600)
            except subprocess.TimeoutExpired:
                print(f"{route} route: no end after {time.monotonic() - t0:.0f} s; stopped")
                return 1
            print(f"{route} route: its process ended (exit code {p.returncode}) after {time.monotonic() - t0:.0f} s",
                  flush=True)
            if not out.exists():
                print(f"  no result (its last error: {_last_error_type(p.stderr)})")
                return 1
            got[route] = json.loads(out.read_text(encoding="utf-8"))
            why = _blocked_message(route, got[route])
            if why:
                print(f"BLOCKED: {why} - an interactive sign-in is needed; nothing was retried")
                return BLOCKED
        return report(picks, got)


if __name__ == "__main__":
    sys.exit(main())
