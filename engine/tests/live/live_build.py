"""Live: sv_ws_build (one call, short spec with repeat/sum shorthands, drafted key points) + an sv_ws_answer
spec_patch follow-up, over real stdio MCP through the relay, into a NEW workbook. Prints timings and asserts the
known BUS numbers. Usage: python tests/live_build.py"""
import json
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path

MCP = Path(__file__).resolve().parents[2]
p = subprocess.Popen([sys.executable, "-B", str(MCP / "run_server.py")], stdin=subprocess.PIPE, stdout=subprocess.PIPE)
inbox: queue.Queue = queue.Queue()
threading.Thread(target=lambda: [inbox.put(json.loads(x)) for x in p.stdout if x.strip()], daemon=True).start()
fails = []


def request(i, method, params=None, timeout=600):
    p.stdin.write((json.dumps({"jsonrpc": "2.0", "id": i, "method": method, "params": params or {}}) + "\n").encode())
    p.stdin.flush()
    t0 = time.monotonic()
    while time.monotonic() - t0 < timeout:
        try:
            m = inbox.get(timeout=1)
        except queue.Empty:
            continue
        if m.get("id") == i:
            return m
    raise TimeoutError(method)


def call(i, name, args):
    t0 = time.monotonic()
    r = request(i, "tools/call", {"name": name, "arguments": args})
    secs = time.monotonic() - t0
    if "error" in r or r["result"].get("isError"):
        txt = r.get("error") or r["result"]["content"][0]["text"]
        print(f"{name} ERROR after {secs:.1f}s: {txt}")
        fails.append(name)
        return None, secs
    return json.loads(r["result"]["content"][0]["text"]), secs


def check(cond, what):
    print(("OK   " if cond else "FAIL ") + what)
    if not cond:
        fails.append(what)


request(1, "initialize", {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "t", "version": "1"}})
p.stdin.write(b'{"jsonrpc":"2.0","method":"notifications/initialized"}\n')
p.stdin.flush()

E = ["ac.EBITDA"]
ACC = ["ac.C30", "ac.C40S", "ac.C70", "ac.C73", "ac.EBITDA"]
POV = {"All Entities": "un.SY.2BUS", "All Scenarios": "Actual", "All Currencies": "NOK"}
months = [f"{m} 25" for m in "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()] + \
         [f"{m} 26" for m in "Jan Feb Mar Apr May Jun Jul Aug Sep".split()]
pulls = [{"id": "m", "connection": "FinMax", "rows": [{"dimension": "All Accounts", "members": ACC}],
          "columns": [{"dimension": "All Time Periods", "members": months}],
          "pov": {**POV, "All Balance Types": "Periodic"}},
         {"id": "ytd", "connection": "FinMax", "rows": [{"dimension": "All Accounts", "members": ACC}],
          "columns": [{"dimension": "All Time Periods", "members": ["Sep 26", "Sep 25", "Dec 25"]}],
          "pov": {**POV, "All Balance Types": "Accumulated"}, "title": "Q1 FinMax YTD"}]
spec = {
    "title": "Business Norway — EBITDA per month 2026 vs 2025", "subtitle": "NOK · FinMax · Actual · un.SY.2BUS",
    "key_points": ["June carries the year (one-off low opex)."],
    "columns": [{"header": "Actual 2026"}, {"header": "Actual 2025"}, {"header": "Variance"},
                {"header": "Variance %", "kind": "prosent1"}],
    "sections": [
        {"title": "EBITDA per month (Periodic)",
         "repeat": {"over": "months:Jan..Dec", "line": {"key": "m{i}", "label": "{v}", "cells": [
             {"pull": "m", "row": E, "col": ["{v} 26"]}, {"pull": "m", "row": E, "col": ["{v} 25"]},
             {"var": [0, 1]}, {"varpct": [0, 1]}]}},
         "lines": [{"key": "ytd", "label": "YTD Jan–Sep", "kind": "grand",
                    "cells": [{"sum": "m[0-8]"}, {"sum": "m[0-8]"}, {"var": [0, 1]}, {"varpct": [0, 1]}]},
                   {"key": "fy", "label": "Full year 2025", "kind": "sum", "cells": [None, {"sum": "m*"}, None, None]}]},
        {"title": "Reconciliation (Accumulated)", "lines": [
            {"key": "eacc", "label": "EBITDA YTD Sep", "cells": [{"pull": "ytd", "row": E, "col": ["Sep 26"]},
                                                                {"pull": "ytd", "row": E, "col": ["Sep 25"]}, None, None]},
            {"key": "fyacc", "label": "EBITDA Dec 25", "cells": [None, {"pull": "ytd", "row": E, "col": ["Dec 25"]},
                                                                 None, None]}]}],
    "checks": [{"label": "Sum of months = Accumulated YTD", "expr": "ytd-eacc", "columns": [0, 1]},
               {"label": "Sum of 2025 months = Accumulated Dec 25", "expr": "fy-fyacc", "columns": [1]}],
    "notes": ["Live build test."], "source": "Source: Oracle EPM Cloud (FinMax) via Smart View.",
    "chart": {"type": "column", "lines": "m*", "columns": [0, 1], "title": "EBITDA per month"}}

print("spec size:", len(json.dumps(spec)), "chars")
out, secs = call(2, "sv_ws_build", {"question": "LIVE BUILD BUS EBITDA monthly", "pulls": pulls, "spec": spec,
                                    "target": "new"})
if out:
    print(f"sv_ws_build {secs:.1f}s wall | timings: {out['timings']}")
    print("key points:", *out["key_points"], sep="\n  - ")
    check(all(v == "OK" for v in out["checks"].values()), f"checks {out['checks']}")
    check("3 495 317 868" in out["table"] and "88 601 304" in out["table"], "BUS YTD EBITDA + variance in table")
    check(out["key_points"][0].startswith("YTD Jan–Sep: 3 495 317 868"), "drafted grand-total point first")
    check(out["key_points"][-1].startswith("June carries"), "own interpretation kept after drafted points")
    check("| Oct |  | 404 657 561 |  |  |" in out["table"], "Oct 26 empty (not pulled), Oct 25 shown, no variance")
    wb = out["workbook"]
    out2, secs2 = call(3, "sv_ws_answer", {"q": 1, "target": wb, "spec_patch": {"notes": ["Patched follow-up."]}})
    if out2:
        print(f"spec_patch follow-up {secs2:.1f}s wall | timings: {out2['timings']}")
        check(all(v == "OK" for v in out2["checks"].values()), "checks OK after patch")
        check("3 495 317 868" in out2["table"], "same numbers after patch")
p.stdin.close()
p.wait(120)
print("\nALL OK" if not fails else "\nFAILURES: " + "; ".join(fails))
