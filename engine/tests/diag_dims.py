"""Why does HypGetDimensions return 'Accounts' (no 'All ') for FinMax in some sessions?"""
import asyncio
import sys

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from smartview_mcp import ops  # noqa: E402
from smartview_mcp.host import ExcelHost  # noqa: E402


def probe(h, conn, sheet):
    ws = h.sheet(sheet, fresh=True)
    ws.Activate()
    h.connect_sheet(sheet, conn)
    out = {"opt16_sheet_before": h._run("B_GetOption", 16, sheet), "opt16_global": h._run("B_GetOption", 16, "")}
    h.run("B_Retrieve", sheet)
    out["grid"] = [list(r) for r in ops._used_values(ws)]
    out["dims"] = h._run("B_GetDimensions", sheet)
    for k, v in {5: 0, 13: "#Missing"}.items():
        h.run("B_SetOption", k, v, sheet)
    h.run("B_Retrieve", sheet)
    out["grid_after_opts"] = [list(r) for r in ops._used_values(ws)]
    out["dims_after_opts"] = h._run("B_GetDimensions", sheet)
    out["sheetinfo"] = h._run("B_SheetInfo", sheet)
    return out


async def main():
    h = ExcelHost()
    for conn in ("FinMax", "FinLite"):
        r = await h.call(lambda hh: probe(hh, conn, f"D_{conn}"), what="probe")
        print(f"\n##### {conn}")
        for k, v in r.items():
            print(k, "=", str(v)[:700])
    h.shutdown()


asyncio.run(main())
