"""Live end-to-end smoke test through the MCP client (needs Smart View + SSO). Not part of the unit suite.
Usage: python tests/live_smoke.py [quick|atlas|hang]"""
import asyncio
import datetime as dt
import json
import sys
import time

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])

from mcp import Client  # noqa: E402

from smartview_mcp import server  # noqa: E402

WHS = dict(connection="FinMax",
           rows=[{"dimension": "All Accounts", "members": ["ac.C30", "ac.C40S", "ac.C70", "ac.EBITDA", "ac.C75"]}],
           columns=[{"dimension": "All Time Periods", "members": ["Aug 26", "Sep 26"]}],
           pov={"All Entities": "un.SY.2WHS", "All Scenarios": "Actual", "All Currencies": "NOK",
                "All Balance Types": "Accumulated"})


def show(name, r, t0):
    print(f"\n=== {name} ({time.monotonic() - t0:.1f}s) error={r.is_error}")
    body = r.structured_content
    if body is None:
        texts = [c.text for c in r.content if hasattr(c, "text")]
        try:
            body = json.loads(texts[0]) if len(texts) == 1 else texts
        except ValueError:
            body = texts
    if isinstance(body, dict) and set(body) == {"result"}:
        body = body["result"]
    s = json.dumps(body, ensure_ascii=False, default=str, indent=1)
    print(s[:2500] + ("\n..." if len(s) > 2500 else ""))
    return body


async def main(mode):
    async with Client(server.mcp) as c:
        tools = await c.list_tools()
        print("tools:", [t.name for t in tools.tools])
        t0 = time.monotonic(); show("sv_status", await c.call_tool("sv_status", {}), t0)
        if mode == "hang":
            t0 = time.monotonic(); show("sv_dimensions FinMax", await c.call_tool("sv_dimensions", {"connection": "FinMax"}), t0)
            h = server.host()
            old_pid = h.pid
            t0 = time.monotonic()
            try:
                await h.call(lambda hh: hh._run("B_TestBusy", 60), timeout=5, what="forced hang")
                print("!! forced hang returned normally")
            except Exception as e:   # noqa: BLE001
                print(f"\n=== forced hang -> {type(e).__name__}: {e} ({time.monotonic() - t0:.1f}s)")
            t0 = time.monotonic(); b = show("sv_query after hang (auto-restart)", await c.call_tool("sv_query", WHS), t0)
            assert "1,994,151,357.55" in b["table"]
            t0 = time.monotonic(); st = show("sv_status after hang", await c.call_tool("sv_status", {}), t0)
            assert st["restarts"] >= 1 and st["pid"] != old_pid, "watchdog did not restart Excel"
            import psutil
            assert not psutil.pid_exists(old_pid), "old Excel still alive"
            print("\nHANG TEST OK")
            return
        t0 = time.monotonic(); show("sv_dimensions FinMax", await c.call_tool("sv_dimensions", {"connection": "FinMax"}), t0)
        t0 = time.monotonic(); b = show("sv_query WHS", await c.call_tool("sv_query", WHS), t0)
        assert "1,994,151,357.55" in b["table"], "WHS regression value missing"
        bad = json.loads(json.dumps(WHS)); bad["rows"][0]["members"].append("ac.NOPE99")
        t0 = time.monotonic(); b = show("sv_query unknown member", await c.call_tool("sv_query", bad), t0)
        assert any("NOPE99" in w for w in b["warnings"])
        t0 = time.monotonic(); show("sv_find_member live", await c.call_tool(
            "sv_find_member", {"text": "un.SY.3WHS_W", "connection": "FinMax", "live": True}), t0)
        sv = dict(WHS, save_xlsx=True, title="smoke WHS")
        t0 = time.monotonic(); b = show("sv_query save_xlsx", await c.call_tool("sv_query", sv), t0)
        assert b["saved_xlsx"] and not b["saved_xlsx"].startswith("NOT"), b["saved_xlsx"]
        t0 = time.monotonic(); show("sv_adhoc_open", await c.call_tool("sv_adhoc_open", {"connection": "FinMax"}), t0)
        t0 = time.monotonic(); show("sv_dimensions FinPlan", await c.call_tool("sv_dimensions", {"connection": "FinPlan"}), t0)
        if mode in ("atlas",):
            t0 = time.monotonic(); show("sv_atlas_build FinMax entities+accounts", await c.call_tool(
                "sv_atlas_build", {"connections": ["FinMax"], "dimensions": ["All Entities", "All Accounts"]}), t0)
            t0 = time.monotonic(); show("sv_find_member WHS", await c.call_tool("sv_find_member", {"text": "WHS"}), t0)
            t0 = time.monotonic(); show("sv_members children 2WHS", await c.call_tool(
                "sv_members", {"connection": "FinMax", "member": "un.SY.2WHS"}), t0)
        if mode in ("hang",):
            h = server.host()
            t0 = time.monotonic()
            try:
                old_pid = h.pid
                await h.call(lambda hh: hh._run("B_TestBusy", 60), timeout=5, what="forced hang")
            except Exception as e:   # noqa: BLE001
                print(f"\n=== forced hang -> {type(e).__name__}: {e} ({time.monotonic() - t0:.1f}s)")
            t0 = time.monotonic(); b = show("sv_query after hang (auto-restart)", await c.call_tool("sv_query", WHS), t0)
            assert "1,994,151,357.55" in b["table"]
            t0 = time.monotonic(); st = show("sv_status after hang", await c.call_tool("sv_status", {}), t0)
            assert st["restarts"] >= 1 and st["pid"] != old_pid, "watchdog did not restart Excel"
            import psutil
            assert not psutil.pid_exists(old_pid), "old Excel still alive"
    print("\nSMOKE OK")


if __name__ == "__main__":
    asyncio.run(main(sys.argv[1] if len(sys.argv) > 1 else "quick"))
