"""Relay (run_server.py) over real stdio JSON-RPC: handshake, tools, reload on a code change (+ tools/list_changed),
recovery after a child crash, clean shutdown. Uses only sv_status (does not start Excel).
Usage: python tests/live_relay.py"""
import json
import os
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path

import psutil

MCP = Path(__file__).resolve().parents[2]
p = subprocess.Popen([sys.executable, "-B", str(MCP / "run_server.py")], stdin=subprocess.PIPE, stdout=subprocess.PIPE)
inbox: queue.Queue = queue.Queue()
threading.Thread(target=lambda: [inbox.put(json.loads(l)) for l in p.stdout if l.strip()], daemon=True).start()
fails = []


def send(msg):
    p.stdin.write((json.dumps(msg) + "\n").encode())
    p.stdin.flush()


def request(i, method, params=None, timeout=90):
    send({"jsonrpc": "2.0", "id": i, "method": method, "params": params or {}})
    notes, t0 = [], time.monotonic()
    while time.monotonic() - t0 < timeout:
        try:
            m = inbox.get(timeout=1)
        except queue.Empty:
            continue
        if m.get("id") == i:
            return m, notes
        notes.append(m.get("method"))
    raise TimeoutError(method)


def status(i):
    r, notes = request(i, "tools/call", {"name": "sv_status", "arguments": {}})
    body = json.loads(r["result"]["content"][0]["text"]) if "result" in r else {"error": r.get("error")}
    return body, notes


def check(cond, what):
    print(("OK   " if cond else "FAIL ") + what)
    if not cond:
        fails.append(what)


seen_children: set[int] = set()


def children():
    out = [c for c in psutil.Process(p.pid).children(recursive=True) if "child_server.py" in " ".join(c.cmdline())]
    seen_children.update(c.pid for c in out)
    return out


r, _ = request(1, "initialize", {"protocolVersion": "2025-06-18", "capabilities": {},
                                 "clientInfo": {"name": "relay-test", "version": "1"}})
check("result" in r and r["result"]["serverInfo"]["name"] == "smartview", "initialize answered by the child")
check(r["result"]["capabilities"].get("tools", {}).get("listChanged") is True, "relay announces tools.listChanged")
send({"jsonrpc": "2.0", "method": "notifications/initialized"})
r, _ = request(2, "tools/list")
names = [t["name"] for t in r["result"]["tools"]]
check({"sv_analyse", "sv_ws_pull", "sv_ws_answer", "sv_atlas_changes"} <= set(names), f"{len(names)} tools incl. analysis tools")
old_excel = set()
if "excel" in sys.argv:                              # real Smart View pull through the relay
    r, _ = request(30, "tools/call", {"name": "sv_query", "arguments": {
        "connection": "FinMax", "rows": [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
        "columns": [{"dimension": "All Time Periods", "members": ["Sep 26"]}],
        "pov": {"All Entities": "un.SY.2WHS", "All Scenarios": "Actual", "All Currencies": "NOK",
                "All Balance Types": "Accumulated"}}}, timeout=400)
    txt = json.dumps(r, ensure_ascii=False)
    flat = txt.replace(",", "").replace(" ", "").replace(" ", "")
    check("1858190823" in flat, "sv_query via relay: Wholesale EBITDA YTD Sep 26 = 1 858 190 823")
    if "1858190823" not in flat:
        print(txt[:1500])
    old_excel = {x.pid for x in psutil.process_iter(["name", "cmdline"]) if x.info["name"] == "EXCEL.EXE"
                 and "-Embedding" in " ".join(x.info["cmdline"] or [])}
    print("server Excel:", old_excel)
children()
s1, _ = status(3)
check(s1.get("behind_reload_proxy") is True, f"sv_status via relay (started {s1.get('server_started')})")
s, notes = status(4)
check(s.get("server_started") == s1.get("server_started") and not notes, "no reload without a code change")

time.sleep(1.1)
f = MCP / "smartview_mcp" / "qlog.py"
os.utime(f, None)                                   # simulate a code edit
s2, notes = status(5)
check("notifications/tools/list_changed" in notes, "tools/list_changed sent after the code change")
check(s2.get("server_started") not in (None, s1.get("server_started")), f"new child after code change ({s2.get('server_started')})")
if old_excel:
    check(not any(psutil.pid_exists(x) for x in old_excel), "old server's Excel closed by the graceful reload")

time.sleep(1.1)
for c in children():                                 # simulate a crash (venv launcher + real interpreter)
    try:
        c.kill()
    except psutil.NoSuchProcess:
        pass
time.sleep(1)
s3, notes = status(6)
children()
check(s3.get("server_started") not in (None, s2.get("server_started")), "recovered after a child crash")

p.stdin.close()
try:
    p.wait(90)
    check(True, f"relay exited rc={p.returncode}")
except subprocess.TimeoutExpired:
    check(False, "relay did not exit after stdin closed")
    p.kill()
time.sleep(1)
left = [x for x in seen_children if psutil.pid_exists(x)]
check(not left, f"no child_server processes of this relay left ({len(seen_children)} seen)")
print("\nALL OK" if not fails else "\nFAILURES: " + "; ".join(fails))
