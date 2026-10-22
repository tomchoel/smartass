"""Can FinMax be brought back to 'Adhoc (Standard)' with the Default alias table and an in-sheet POV row?"""
import asyncio
import sys

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from smartview_mcp import ops  # noqa: E402
from smartview_mcp.host import ExcelHost  # noqa: E402


def snap(h, sheet, tag):
    ws = h.wb.Worksheets(sheet)
    info = h._run("B_SheetInfo", sheet)
    meta = dict(zip(info[1], info[2])) if info and info[0] == "OK" else info
    print(f"[{tag}] type={meta.get('Sheet Type')} alias={meta.get('Alias Table')} grid={[list(r) for r in ops._used_values(ws)][:3]}",
          flush=True)


def probe(h):
    s = "D1"
    ws = h.sheet(s, fresh=True)
    ws.Activate()
    for attempt in range(3):
        try:
            h.connect_sheet(s, "FinMax")
            break
        except Exception as e:   # noqa: BLE001
            print("connect attempt", attempt, "->", e, flush=True)
            import time
            time.sleep(5)
    h.run("B_Retrieve", s)
    snap(h, s, "as connected")
    print("conn alias ->", h._run("B_SetConnAliasTable", "FinMax", "Default"))
    h.run("B_Retrieve", s)
    snap(h, s, "after conn alias Default")
    print("sheet alias ->", h._run("B_SetAliasTable", s, "Default"))
    print("opt16=2 ->", h._run("B_SetOption", 16, 2, s))
    h.run("B_Retrieve", s)
    snap(h, s, "after sheet alias + 16=2")
    s2 = "D2"
    ws2 = h.sheet(s2, fresh=True)
    ws2.Activate()
    h.connect_sheet(s2, "FinMax")
    h.run("B_Retrieve", s2)
    snap(h, s2, "new sheet afterwards")


async def main():
    h = ExcelHost()
    await h.call(probe, what="probe")
    h.shutdown()


asyncio.run(main())
