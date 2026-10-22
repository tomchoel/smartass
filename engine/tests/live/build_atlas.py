"""Build the full atlas (all cubes, all dimensions) through the MCP client; prints a summary."""
import asyncio
import json
import sys
import time

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from mcp import Client  # noqa: E402

from smartview_mcp import server  # noqa: E402


async def main():
    async with Client(server.mcp) as c:
        t0 = time.monotonic()
        args = {"connections": sys.argv[1:]} if len(sys.argv) > 1 else {}
        r = await c.call_tool("sv_atlas_build", args)
        body = json.loads(r.content[0].text)
        for b in body["built"]:
            print(json.dumps(b, ensure_ascii=False))
        print(f"total {time.monotonic() - t0:.0f}s error={r.is_error}")
        r = await c.call_tool("sv_atlas_status", {})
        print(r.content[0].text[:3000])


asyncio.run(main())
