"""Live acceptance for sv_drill (needs Smart View + SSO; opens NO visible window). Not part of the unit suite.
Usage: PY tests/live_drill.py [link|fetch|tool|split|all]"""
import asyncio
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

if len(sys.argv) > 1 and sys.argv[1] == "split":
    import os
    os.environ["SV_DRILL_SPLIT_AT"] = "600"          # must be set before config is imported

from smartview_mcp import ops  # noqa: E402
from smartview_mcp.host import ExcelHost  # noqa: E402

CELLS = [  # (cell, expected lines, expected sum) - reconciled in the 15.10.2026 spike
    ({"All Accounts": "ac.40100", "All Entities": "un.SY.E400", "All Time Periods": "Sep 26"}, 81, -1140790.92),
    ({"All Accounts": "ac.C40S", "All Entities": "un.SY.2WHS", "All Time Periods": "Sep 26"}, 418, -8025554.03),
    ({"All Accounts": "ac.C40S", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"}, 120, -35008728.86),
    ({"All Accounts": "ac.C40S", "All Entities": "un.SY.2WHS", "All Time Periods": "Q3 26"}, 1231, -24706615.55),
]


async def link():
    h = ExcelHost(visible=False)
    try:
        for cell, _, total in CELLS:
            r = await h.call(lambda hh, c=cell: ops.drill_link(hh, "FinLite", c), timeout=300, what="drill_link")
            print(cell, r["report"], round(r["cube_value"], 2), (r["url"] or "")[:70])
            assert r["url"] and r["url"].startswith("aif/") and round(r["cube_value"], 2) == total
    finally:
        h.shutdown()


def visible_windows() -> set[str]:
    import win32gui
    out = set()

    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            out.add(f"{hwnd}:{win32gui.GetWindowText(hwnd)}")
    win32gui.EnumWindows(cb, None)
    return out


def our_chrome() -> list[int]:
    """chrome.exe processes started with OUR profile dir (listing only, kills nothing)."""
    import psutil
    from smartview_mcp import config
    from smartview_mcp.drill_browser import is_our_chrome
    return [p.pid for p in psutil.process_iter(["name", "cmdline"])
            if is_our_chrome(p.info["name"], p.info["cmdline"], config.DRILL_PROFILE)]


async def fetch():
    from smartview_mcp import drill_lines as dl
    from smartview_mcp import config
    from smartview_mcp.drill_browser import DrillBrowser
    h, b = ExcelHost(visible=False), DrillBrowser()
    before = visible_windows()
    try:
        for cell, n, total in CELLS:
            link = await h.call(lambda hh, c=cell: ops.drill_link(hh, "FinLite", c), timeout=300, what="drill_link")
            res = b.fetch(link["url"]).result(timeout=600)
            cols = dl.columns_from_view(res["view"])
            rows, stats = dl.normalize(res["items"], cols,
                                       dl.load_text_fixes(config.TEXT_FIXES, config.USER_TEXT_FIXES))
            rec = dl.reconcile(link["cube_value"], rows)
            print(cell["All Entities"], cell["All Time Periods"], rec["lines"], rec["lines_sum"], rec["status"],
                  f"{res['secs']}s", stats["unfixed"], f"our chrome.exe: {len(our_chrome())}", flush=True)
            assert rec["lines"] == n and rec["lines_sum"] == total and rec["status"] == "exact"
    finally:
        b.shutdown()
        h.shutdown()
    new = {w for w in visible_windows() - before if "chrome" in w.lower() or "edge" in w.lower()}
    assert not new, f"visible browser windows appeared: {new}"
    left = our_chrome()
    assert not left, f"our chrome.exe still running after shutdown: {left}"


async def _call(c, name: str, args: dict) -> dict:
    r = await c.call_tool(name, args)
    if r.is_error:
        raise AssertionError(f"{name} failed: {r.content}")
    body = r.structured_content or json.loads(r.content[0].text)   # a plain -> dict tool has no output schema
    return body.get("result", body)


async def tool(cell=CELLS[3][0], expect=CELLS[3][1:], redrill=CELLS[1]):
    """The MCP tool end to end (in-process server): a NEW analysis workbook in the user's Excel with 'Q1 FinLite',
    'Q1 Linjer' (table Linjer_Q1) and 'Kilder'; exact reconciliation; then an answer with drill cells, and a
    re-drill (a smaller cell) into the SAME workbook with q=1: the table is refreshed in place and the answer's
    formulas follow it. No visible browser; our Chrome closed after."""
    from mcp import Client
    from smartview_mcp import server
    before = visible_windows()
    async with Client(server.mcp) as c:
        body = await _call(c, "sv_drill", {"question": "Live test: Wholesale COGS Q3 26", "cell": cell,
                                           "target": "new"})
        rec = body["reconciliation"]
        print(body.get("workbook"), body.get("lines_sheet"), body.get("table"), body.get("cube_pull", {}).get("sheet"),
              rec, body["timings"], body.get("warnings"), flush=True)
        assert (rec["lines"], rec["lines_sum"], rec["status"]) == (expect[0], expect[1], "exact")
        assert body["text_repairs"]["unfixed"] == []
        wb, cp = body["workbook"], body["cube_pull"]
        spec = {"title": "Live test drill answer", "columns": [{"header": "NOK"}], "auto_points": False,
                "sections": [{"title": "Lines", "lines": [
                    {"key": "all", "label": "All lines", "cells": [{"drill": "d", "where": {}}]},
                    {"key": "pay", "label": "Payables", "cells": [{"drill": "d", "where": {"JE_SOURCE": "Payables"}}]},
                    {"key": "n", "label": "Line count", "cells": [{"drill": "d", "where": {}, "agg": "count"}]},
                    {"key": "cube", "label": "Cube", "cells": [{"pull": cp["sheet"], "row": cp["row_keys"][0],
                                                                 "col": cp["col_keys"][0]}]}]}]}
        ans = await _call(c, "sv_ws_answer", {"q": body["q"], "spec": spec, "target": wb})
        print("answer:", ans["answer_sheet"], ans["warnings"], ans.get("formula_style"), flush=True)
        print(ans["table"], flush=True)
        r2 = await _call(c, "sv_drill", {"question": "Live test: re-drill Wholesale COGS Sep 26", "cell": redrill[0],
                                         "target": wb, "q": 1})
        rec2 = r2["reconciliation"]
        print("re-drill:", r2.get("workbook"), r2.get("lines_sheet"), r2.get("table"), r2.get("notes"), rec2,
              r2["timings"], flush=True)
        assert r2["workbook"] == wb and r2["q"] == 1 and r2["lines_sheet"] == body["lines_sheet"]
        assert r2["table"] == "Linjer_Q1" and any("table Linjer_Q1 refreshed in place" in n for n in r2["notes"])
        assert any("grid" in n and "rebuilt" in n for n in r2["notes"]) and r2["cube_pull"]["sheet"] == cp["sheet"]
        assert (rec2["lines"], rec2["lines_sum"], rec2["status"]) == (redrill[1], redrill[2], "exact")
        r3 = await _call(c, "sv_drill", {"question": "Live test: same cell again", "cell": redrill[0],
                                         "target": wb, "q": 1})
        print("same cell again:", r3.get("notes"), r3["reconciliation"]["status"], r3["timings"]["_total"], flush=True)
        assert any("refreshed in place" in n and "grid" in n for n in r3["notes"])
        assert r3["cube_pull"]["sheet"] == cp["sheet"] and r3["reconciliation"]["lines"] == redrill[1]
    new = {w for w in visible_windows() - before if "chrome" in w.lower() or "edge" in w.lower()}
    assert not new, new
    left = our_chrome()
    assert not left, f"our chrome.exe still running after the server stopped: {left}"


async def split(cell=CELLS[3][0], expect=CELLS[3][1:]):
    """Forced split: SV_DRILL_SPLIT_AT=600 (set at the top of this file in 'split' mode, before config is imported)
    makes Q3 26 (1 231 lines) too big, so sv_drill drills its months instead. Cache only (target 'none')."""
    from mcp import Client
    from smartview_mcp import config, server
    assert config.DRILL_SPLIT_AT == 600, config.DRILL_SPLIT_AT
    before = visible_windows()
    async with Client(server.mcp) as c:
        body = await _call(c, "sv_drill", {"question": "Live test: split", "cell": cell, "target": "none"})
        got = [(p["cell"]["All Time Periods"], p["lines"], p["status"]) for p in body["pieces"]]
        print(got, body["reconciliation"], body["timings"], flush=True)
        assert [p[0] for p in got] == ["Jul 26", "Aug 26", "Sep 26"], got
        assert all(p[2] == "exact" for p in got), got
        rec = body["reconciliation"]
        assert (rec["lines"], rec["lines_sum"], rec["status"]) == (expect[0], expect[1], "exact"), rec
        assert rec["cube_value"] == expect[1] and body.get("split_check") == 0.0, (rec, body.get("split_check"))
        assert "workbook" not in body, "target 'none' must not write a workbook"
    new = {w for w in visible_windows() - before if "chrome" in w.lower() or "edge" in w.lower()}
    assert not new, new
    left = our_chrome()
    assert not left, f"our chrome.exe still running after the server stopped: {left}"


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "all"
    t0 = time.monotonic()
    if mode == "split":                                 # not in 'all': it needs SV_DRILL_SPLIT_AT set at import
        asyncio.run(split())
    if mode in ("link", "all"):
        asyncio.run(link())
    if mode in ("fetch", "all"):
        asyncio.run(fetch())
    if mode in ("tool", "all"):
        asyncio.run(tool())
    print(f"done in {time.monotonic() - t0:.0f}s")
