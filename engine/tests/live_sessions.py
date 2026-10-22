"""Live: two client sessions in two projects at the same time - each its own relay+server (as Claude Code / Copilot
start them), each answering two questions. Expect: one file per session in analyser\\<project>\\, both questions in it
(Q1, Q2), the files never shared, a closed session file reopened. Usage: python tests/live_sessions.py <scratch dir>"""
import json
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path

MCP = Path(__file__).resolve().parents[1]
SCRATCH = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
results, fails = {}, []


class Session:
    def __init__(self, project: str):
        self.cwd = SCRATCH / project
        self.cwd.mkdir(parents=True, exist_ok=True)
        self.p = subprocess.Popen([sys.executable, "-B", str(MCP / "run_server.py")], cwd=str(self.cwd),
                                  stdin=subprocess.PIPE, stdout=subprocess.PIPE)
        self.inbox: queue.Queue = queue.Queue()
        threading.Thread(target=lambda: [self.inbox.put(json.loads(x)) for x in self.p.stdout if x.strip()],
                         daemon=True).start()
        self.n = 0
        self.req("initialize", {"protocolVersion": "2025-06-18", "capabilities": {},
                                "clientInfo": {"name": project, "version": "1"}})
        self.p.stdin.write(b'{"jsonrpc":"2.0","method":"notifications/initialized"}\n')
        self.p.stdin.flush()

    def req(self, method, params, timeout=600):
        self.n += 1
        i = self.n
        self.p.stdin.write((json.dumps({"jsonrpc": "2.0", "id": i, "method": method, "params": params}) + "\n").encode())
        self.p.stdin.flush()
        t0 = time.monotonic()
        while time.monotonic() - t0 < timeout:
            try:
                m = self.inbox.get(timeout=1)
            except queue.Empty:
                continue
            if m.get("id") == i:
                return m
        raise TimeoutError(method)

    def build(self, question, entity):
        r = self.req("tools/call", {"name": "sv_ws_build", "arguments": {
            "question": question,
            "pulls": [{"id": "y", "connection": "FinMax",
                       "rows": [{"dimension": "All Accounts", "members": ["ac.C30", "ac.EBITDA"]}],
                       "columns": [{"dimension": "All Time Periods", "members": ["Sep 26", "Sep 25"]}],
                       "pov": {"All Entities": entity, "All Scenarios": "Actual", "All Currencies": "NOK",
                               "All Balance Types": "Accumulated"}}],
            "spec": {"title": question, "columns": [{"header": "YTD Sep 26"}, {"header": "YTD Sep 25"},
                                                    {"header": "Variance"}],
                     "sections": [{"title": "P&L", "lines": [
                         {"key": "e", "label": "EBITDA", "kind": "grand",
                          "cells": [{"pull": "y", "row": ["ac.EBITDA"], "col": ["Sep 26"]},
                                    {"pull": "y", "row": ["ac.EBITDA"], "col": ["Sep 25"]}, {"var": [0, 1]}]}]}]}}})
        if "error" in r or r["result"].get("isError"):
            return {"error": str(r.get("error") or r["result"]["content"][0]["text"])[:400]}
        return json.loads(r["result"]["content"][0]["text"])

    def close(self):
        self.p.stdin.close()
        self.p.wait(120)


def close_in_excel(path: str):
    sys.path.insert(0, str(MCP))
    import pythoncom
    from smartview_mcp import userexcel as u
    pythoncom.CoInitialize()
    for pid, h, app in u.interactive_excels():
        for wb in list(app.Workbooks):
            if wb.Name.lower() == Path(path).name.lower():
                wb.Close(False)
                return True
    return False


def run(project, entity, close_between):
    s = Session(project)
    t0 = time.monotonic()
    a = s.build(f"{project} question 1", entity)
    if close_between and "workbook" in a:
        results[project + " closed"] = close_in_excel(a["workbook"])
    b = s.build(f"{project} question 2", entity)
    results[project] = {"q1": a, "q2": b, "secs": round(time.monotonic() - t0, 1)}
    s.close()


threads = [threading.Thread(target=run, args=("Project Alpha", "un.SY.2WHS", True)),
           threading.Thread(target=run, args=("Project Beta", "un.SY.2BUS", False))]
for th in threads:
    th.start()
for th in threads:
    th.join()


def check(cond, what):
    print(("OK   " if cond else "FAIL ") + what)
    if not cond:
        fails.append(what)


for proj in ("Project Alpha", "Project Beta"):
    r = results.get(proj, {})
    q1, q2 = r.get("q1", {}), r.get("q2", {})
    print(f"\n{proj}: {r.get('secs')} s | q1 {q1.get('error') or (q1.get('question'), q1.get('workbook'))}"
          f" | q2 {q2.get('error') or (q2.get('question'), q2.get('workbook'))}")
    print("   q1 timings:", q1.get("timings"))
    check("workbook" in q1 and f"\\analyser\\{proj}\\" in q1["workbook"], f"{proj}: file in analyser\\{proj}\\")
    check(q1.get("workbook") == q2.get("workbook"), f"{proj}: both questions in the SAME session file")
    check(q1.get("question") == "Q1" and q2.get("question") == "Q2", f"{proj}: numbered Q1, Q2")
check(results.get("Project Alpha", {}).get("q1", {}).get("workbook") !=
      results.get("Project Beta", {}).get("q1", {}).get("workbook"), "the two sessions never share a file")
check(results.get("Project Alpha closed") is True, "Alpha's file was closed between its questions (and reopened)")
print("\nALL OK" if not fails else "\nFAILURES: " + "; ".join(fails))
