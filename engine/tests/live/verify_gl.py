"""Verify the GL-extract recipe: rebuild an archived FinMax GL sheet with sv_query and compare every cell.
Usage: python verify_gl.py <workbook> [sheet=GL] [period=Dec 25]"""
import asyncio
import collections
import json
import sys

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
import openpyxl  # noqa: E402
from mcp import Client  # noqa: E402

from smartview_mcp import server  # noqa: E402
from smartview_mcp.grid import name_from_label  # noqa: E402

path = sys.argv[1]
sheet = sys.argv[2] if len(sys.argv) > 2 else "GL"
period = sys.argv[3] if len(sys.argv) > 3 else "Dec 25"

ws = openpyxl.load_workbook(path, data_only=True, read_only=True)[sheet]
grid = [list(r) for r in ws.iter_rows(values_only=True)]
hdr_p, hdr_b, hdr_a = grid[1], grid[2], grid[5]
cols = [(j, name_from_label(hdr_a[j])) for j in range(2, len(hdr_a))
        if hdr_p[j] == period and hdr_b[j] == "Accumulated" and name_from_label(hdr_a[j])]
ents = [(i, r[0].strip()) for i, r in enumerate(grid[6:], start=6) if r and isinstance(r[0], str) and r[0].strip().startswith("un.")]
accounts = list(dict.fromkeys(n for _, n in cols))
entities = list(dict.fromkeys(n for _, n in ents))
print(f"archive: {len(entities)} entities x {len(accounts)} accounts ({period}, Accumulated)")


async def main():
    async with Client(server.mcp) as c:
        r = await c.call_tool("sv_query", {
            "connection": "FinMax",
            "rows": [{"dimension": "All Entities", "members": entities}],
            "columns": [{"dimension": "All Accounts", "members": accounts}],
            "pov": {"All Scenarios": "Actual", "All Currencies": "NOK", "All Time Periods": period,
                    "All Balance Types": "Accumulated"},
            "suppress_missing": False, "save_xlsx": True, "title": f"verify GL {period}", "max_preview_rows": 1})
        if r.is_error:
            print("TOOL ERROR:", r.content[0].text)
            return
        body = json.loads(r.content[0].text)
        print("warnings:", body["warnings"][:3], "| retrieve_s:", body["retrieve_seconds"], "| saved:", body["saved_xlsx"])
    out = openpyxl.load_workbook(body["saved_xlsx"], data_only=True, read_only=True)["Data"]
    rows = [list(x) for x in out.iter_rows(values_only=True)]
    head = rows[0]
    live = {}
    for r in rows[1:]:
        for k, acc in enumerate(accounts):
            live[(r[0], acc)] = r[1 + k]
    n = same = 0
    diffs = []
    for i, ent in ents:
        for j, acc in cols:
            a = grid[i][j]
            a = float(a) if isinstance(a, (int, float)) else 0.0
            b = live.get((ent, acc))
            b = float(b) if isinstance(b, (int, float)) else 0.0
            n += 1
            if abs(a - b) <= 0.5:
                same += 1
            else:
                diffs.append((ent, acc, a, b))
    print(f"cells compared: {n}  identical: {same}  differing: {len(diffs)}")
    by_ent = collections.Counter(d[0] for d in diffs)
    for ent, k in by_ent.most_common(15):
        tot_a = sum(d[2] for d in diffs if d[0] == ent)
        tot_b = sum(d[3] for d in diffs if d[0] == ent)
        print(f"  {ent}: {k} cells  archive_sum={tot_a:,.0f} live_sum={tot_b:,.0f}")


asyncio.run(main())
