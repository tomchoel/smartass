"""End-to-end: run the analysis templates through the MCP client into a NEW analysis workbook (visible Excel).
Asserts: no tool error, all checks OK (except the documented FinPlan/FinMax info line), no formula errors, and the
known Technology YTD Sep 26 numbers. Usage: python tests/live_analysis.py [all|pnl|plan|trend|cost]"""
import asyncio
import json
import sys
import time

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from mcp import Client  # noqa: E402

from smartview_mcp import server  # noqa: E402

RUNS = {
    "pnl": ("Hva er inntekter, varekost, driftskostnader og EBITDA for Technology hittil i år, og årsverk?",
            "division_pnl", {"entity": "un.SY.2TEC", "period": "Sep 26"}),
    "trend": ("Månedlig utvikling for Technology i år", "monthly_trend", {"entity": "un.SY.2TEC", "through": "Sep 26"}),
    "plan": ("Technology faktisk mot budsjett og prognose hittil i år", "actual_vs_plan",
             {"entity": "2TEC Technology", "month": "Sep", "year": "FY26"}),
    "cost": ("Kostnadsbase per koststed i Technology", "cost_base", {"parent": "un.SY.2TEC", "period": "Sep 26", "top": 15}),
}


def body(r):
    t = r.content[0].text if r.content else ""
    try:
        return json.loads(t)
    except ValueError:
        return {"_text": t}


async def freeform(c, target, failures):
    """sv_ws_pull + sv_ws_answer, then a simulated server restart (session dropped) and a follow-up rebuild."""
    r = await c.call_tool("sv_ws_pull", {
        "question": "Wholesale revenue and EBITDA YTD Sep 26 vs Sep 25", "connection": "FinMax", "target": target,
        "rows": [{"dimension": "Accounts", "members": ["ac.C30", "ac.EBITDA"]}],
        "columns": [{"dimension": "All Time Periods", "members": ["Sep 26", "Sep 25"]}],
        "pov": {"All Entities": "un.SY.2WHS", "All Scenarios": "Actual", "All Currencies": "NOK",
                "All Balance Types": "Accumulated"}})
    b = body(r)
    print("\n===== freeform pull error=", r.is_error, b.get("_text", "")[:300] or (b["sheet"], b["q"], b["warnings"]))
    if r.is_error:
        failures.append("freeform: pull error")
        return
    sheet, q, wbk = b["sheet"], b["q"], b["workbook"]
    spec = {"title": "Wholesale — revenue and EBITDA YTD Sep 26", "subtitle": "NOK · FinMax · Actual · un.SY.2WHS",
            "key_points": ["Free-form test."],
            "columns": [{"header": "YTD Sep 26"}, {"header": "YTD Sep 25"}, {"header": "Variance"},
                        {"header": "Variance %", "kind": "prosent1"}],
            "sections": [{"title": "P&L", "lines": [
                {"key": "rev", "label": "Revenue", "cells": [{"pull": sheet, "row": ["ac.C30"], "col": ["Sep 26"]},
                                                             {"pull": sheet, "row": ["ac.C30"], "col": ["Sep 25"]},
                                                             {"var": [0, 1]}, {"varpct": [0, 1]}]},
                {"key": "e", "label": "EBITDA", "kind": "grand",
                 "cells": [{"pull": sheet, "row": ["ac.EBITDA"], "col": ["Sep 26"]},
                           {"pull": sheet, "row": ["ac.EBITDA"], "col": ["Sep 25"]}, {"var": [0, 1]}, {"varpct": [0, 1]}]}]}],
            "checks": [{"label": "EBITDA below revenue", "expr": "rev-e-rev+e", "columns": [0]}]}
    r = await c.call_tool("sv_ws_answer", {"q": q, "spec": spec, "target": wbk})
    b = body(r)
    print("freeform answer error=", r.is_error, b.get("_text", "")[:300] or b["table"])
    if r.is_error or "1 994 151 358" not in b.get("table", ""):
        failures.append("freeform: answer wrong/error")
    server._sessions.clear()                                         # simulate a server restart
    spec["key_points"] = ["Rebuilt after restart."]
    r = await c.call_tool("sv_ws_answer", {"q": q, "spec": spec, "target": wbk})
    b = body(r)
    print("freeform rebuild after restart error=", r.is_error, (b.get("_text") or "")[:300])
    if r.is_error or "1 994 151 358" not in b.get("table", ""):
        failures.append("freeform: rebuild after restart failed")


async def main(which):
    names = list(RUNS) if which in ("all", "full") else ([] if which == "ws" else [which])
    failures = []
    target = "new"
    async with Client(server.mcp) as c:
        if which in ("ws", "full"):
            await freeform(c, target, failures)
        for n in names:
            question, template, params = RUNS[n]
            t0 = time.monotonic()
            r = await c.call_tool("sv_analyse", {"question": question, "template": template, "params": params,
                                                 "target": target})
            b = body(r)
            print(f"\n===== {n} ({time.monotonic() - t0:.1f}s) error={r.is_error}")
            if r.is_error:
                print(b.get("_text"))
                failures.append(f"{n}: tool error")
                continue
            target = b["workbook"]                  # later runs go into the same new workbook
            print("workbook:", b["workbook"], "| sheets:", b["answer_sheet"], b["data_sheets"])
            print("key points:", *b["key_points"], sep="\n  - ")
            print(b["table"])
            print("checks:", b["checks"])
            print("warnings:", b["warnings"], "| lint:", b["style_lint"][:5])
            bad = [k for k, v in b["checks"].items() if v != "OK"]
            if bad:
                failures.append(f"{n}: checks not OK: {bad}")
            if any("formula error" in w for w in b["warnings"]):
                failures.append(f"{n}: formula errors")
            if n == "pnl" and "(1 112 834 010)" not in b["table"]:
                failures.append("pnl: EBITDA YTD Sep 26 value not found in table")
    print("\nFAILURES:" if failures else "\nALL OK", *failures, sep="\n  ")


asyncio.run(main(sys.argv[1] if len(sys.argv) > 1 else "all"))
