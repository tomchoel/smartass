"""Probe FinPlan member naming: are atlas labels (aliases) usable as members, and what are the real names?"""
import asyncio
import json
import sys

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from mcp import Client  # noqa: E402

from smartview_mcp import server  # noqa: E402


def body(r):
    t = r.content[0].text if r.content else ""
    try:
        return json.loads(t)
    except ValueError:
        return t


async def main():
    async with Client(server.mcp) as c:
        for m in ("4TIP_I Tipo", "4TIP_I", "BANORWAY", "FY26", "Target 2025", "Wholesale"):
            r = await c.call_tool("sv_member_info", {"connection": "FinPlan", "member": m})
            b = body(r)
            info = b.get("info", {}) if isinstance(b, dict) else {}
            print(f"{m!r}: error={r.is_error} found={b.get('found') if isinstance(b, dict) else b} "
                  f"name={info.get('Name')} dim={info.get('Dimension')} parent={info.get('ParentMbrName')} "
                  f"aliases={str(info.get('AliasNames'))[:80]}")
        r = await c.call_tool("sv_query", {
            "connection": "FinPlan",
            "rows": [{"dimension": "Entity", "members": ["BANORWAY", "4TIP_I Tipo", "4TIP_I"]},
                     {"dimension": "Account", "members": ["Account"]}],
            "columns": [{"dimension": "Years", "members": ["FY25", "FY26"]}, {"dimension": "Period", "members": ["TotalYear"]}],
            "pov": {"Scenario": "Target 2025", "Version": "Final"}, "suppress_missing": False})
        b = body(r)
        print("query error:", r.is_error)
        print(b["table"] if isinstance(b, dict) else b)
        if isinstance(b, dict):
            print("warnings:", b["warnings"])


asyncio.run(main())
