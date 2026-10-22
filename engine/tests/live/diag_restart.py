"""Diagnose what blocks the Excel restart after a watchdog kill: dump top-level windows of new Excel processes."""
import asyncio
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
import psutil  # noqa: E402

from smartview_mcp import ops  # noqa: E402
from smartview_mcp.host import ExcelHost  # noqa: E402

DUMP = str(Path(__file__).resolve().parents[2] / "scripts" / "uia_dump.ps1")


def dump(tag):
    out = subprocess.run(["powershell", "-NoProfile", "-File", DUMP, "-MaxDepth", "12"], capture_output=True,
                         text=True, timeout=90).stdout
    print(f"--- UIA dump {tag} ---\n{out[:4000]}", flush=True)


async def main():
    h = ExcelHost()
    await h.call(lambda hh: ops.dimensions(hh, "FinMax"), what="dims")
    print("started pid", h.pid, flush=True)
    try:
        await h.call(lambda hh: hh._run("B_TestBusy", 60), timeout=5, what="forced hang")
    except Exception as e:   # noqa: BLE001
        print("hang ->", type(e).__name__, flush=True)
    task = asyncio.create_task(h.call(lambda hh: ops.dimensions(hh, "FinMax"), timeout=120, what="restart"))
    for i in range(4):
        await asyncio.sleep(15)
        procs = [(p.pid, " ".join(p.cmdline())[-40:]) for p in psutil.process_iter(["name"])
                 if (p.info["name"] or "").lower() == "excel.exe"]
        print(f"t+{15 * (i + 1)}s excel procs: {procs} | host pid {h.pid} | done={task.done()}", flush=True)
        if task.done():
            break
        await asyncio.to_thread(dump, f"t+{15 * (i + 1)}s")
    try:
        print("restart result:", str(await task)[:200], flush=True)
    except Exception as e:   # noqa: BLE001
        print("restart failed:", e, flush=True)
    h.shutdown()


asyncio.run(main())
