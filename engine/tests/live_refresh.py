"""Refresh a COPY of an archive Smart View workbook via the MCP and compare numbers with the archived values.
FY2025 actuals should be unchanged, so this is also a reconciliation check. Usage: python live_refresh.py <xlsx>"""
import asyncio
import json
import sys

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
import openpyxl  # noqa: E402
from mcp import Client  # noqa: E402

from smartview_mcp import server  # noqa: E402


def numbers(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    out[(ws.title, c.coordinate)] = float(c.value)
    wb.close()
    return out


async def main(src):
    before = numbers(src)
    async with Client(server.mcp) as c:
        r = await c.call_tool("sv_refresh_workbook", {"path": src, "output_name": "refreshtest"})
        if r.is_error:
            print("TOOL ERROR:", r.content[0].text)
            return
        body = json.loads(r.content[0].text)
        print(json.dumps(body, ensure_ascii=False, indent=1)[:3000])
    if r.is_error or not body.get("saved"):
        print("REFRESH FAILED")
        return
    after = numbers(body["output"])
    common = set(before) & set(after)
    diff = [(k, before[k], after[k]) for k in common if abs(before[k] - after[k]) > 0.5]
    print(f"numeric cells: before={len(before)} after={len(after)} common={len(common)} differing={len(diff)}")
    for k, b, a in sorted(diff)[:15]:
        print(f"  {k}: archive={b:,.2f} refreshed={a:,.2f} delta={a - b:,.2f}")


asyncio.run(main(sys.argv[1]))
