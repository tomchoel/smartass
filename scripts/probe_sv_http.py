"""Stage 0 probe: read one grid from the Smart View provider over HTTP, without Excel.

Signs in through the drill's hidden Edge (Telenor SSO), then posts Smart View provider XML from inside that page:
ConnectToProvider -> OpenApplication -> OpenCube -> ExecuteGrid (free-form retrieve). Read-only: any request that
could write (submit, WriteBack, SetVals, business rules, calculate ...) is refused before it is sent.

Usage (from engine\\):
  .venv\\Scripts\\python.exe ..\\scripts\\probe_sv_http.py --entity un.SY.2WHS --account ac.EBITDA --periods "Aug 26,Sep 26"
Prints the returned member names and values. Never commit the output (real figures)."""
from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "engine"))
from smartview_mcp.drill_browser import DrillBrowser  # noqa: E402

WRITES = ("submit", "writeback", "setvals", "settxts", "setcell", "setsubvar", "launch", "calculate", "consolidate",
          "copydata", "addmembers", "delete")
DIMS = ["Accounts", "Counterparts", "Currencies", "Entities", "Functions", "Products", "Projects", "Reserves",
        "Scenarios", "Systems", "GLSources", "LegalEntities"]           # FinMax POV dimensions, in grid order


def execute_grid(cube: str, pov: dict[str, str], periods: list[str], balance: str) -> str:
    """Free-form grid: row 0 = POV members, row 1 = periods, row 2 = balance type + data cells."""
    cols = 1 + len(DIMS)
    rows = [[""] + [pov.get(d, d) for d in DIMS],
            [""] + periods + [""] * (cols - 1 - len(periods)),
            [balance] + [""] * (cols - 1)]
    types = ["7"] + ["0"] * len(DIMS) + ["7"] + ["0"] * len(periods) + ["7"] * (cols - 1 - len(periods)) + \
            ["0"] + ["2"] * len(periods) + ["7"] * (cols - 1 - len(periods))
    vals = [v for r in rows for v in r]
    return (f'<req_ExecuteGrid><sID>{{sID}}</sID><action>retrieve</action><ords></ords><preferences/>'
            f'<grid><cube>{cube}</cube><gridType>1</gridType><dims/><slices><slice rows="3" cols="{cols}"><data>'
            f'<range start="0" end="{len(vals) - 1}"><vals enc="1">{"|".join(vals)}</vals>'
            f'<types enc="0">{"|".join(types)}</types></range></data></slice></slices></grid></req_ExecuteGrid>')


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--app", default="FinMax")
    ap.add_argument("--cube", default="FinMax")
    ap.add_argument("--entity", required=True)
    ap.add_argument("--account", required=True)
    ap.add_argument("--periods", required=True, help="comma-separated, e.g. 'Aug 26,Sep 26'")
    ap.add_argument("--scenario", default="Actual")
    ap.add_argument("--currency", default="NOK")
    ap.add_argument("--balance", default="Periodic")
    a = ap.parse_args()
    pov = {"Accounts": a.account, "Entities": a.entity, "Scenarios": a.scenario, "Currencies": a.currency}
    steps = [
        '<req_ConnectToProvider><ClientXMLVersion>4.2.5.7.3</ClientXMLVersion><lngs enc="0">en_US</lngs>'
        '<usr></usr><pwd></pwd><sso></sso></req_ConnectToProvider>',
        f'<req_OpenApplication><sID>{{sID}}</sID><srv>{a.app}</srv><app>{a.app}</app><type></type><url></url>'
        '<sessionInfo></sessionInfo></req_OpenApplication>',
        f'<req_OpenCube><sID>{{sID}}</sID><srv>{a.app}</srv><app>{a.app}</app><cube>{a.cube}</cube><type></type>'
        '<url></url></req_OpenCube>',
        execute_grid(a.cube, pov, [p.strip() for p in a.periods.split(",")], a.balance),
    ]
    bad = [w for s in steps for w in WRITES if w in s.lower()]
    if bad:
        raise SystemExit(f"refused: request contains a write operation ({bad})")

    b = DrillBrowser()

    def run(s):
        s.deadline = time.monotonic() + 180
        b._ensure(s)
        sid, out = "", []
        for body in steps:
            t0 = time.monotonic()
            r = s.page.evaluate("""async (b) => { const r = await fetch('/HyperionPlanning/SmartView', {method: 'POST',
                credentials: 'include', body: b, headers: {'Content-Type': 'text/xml'}}); return await r.text(); }""",
                                body.replace("{sID}", sid))
            m = re.search(r"<sID>(.*?)</sID>", r)
            sid = m.group(1) if m else sid
            out.append((body[5:body.index(">")], round(time.monotonic() - t0, 2), r))
        return out

    try:
        for name, secs, r in b.submit(run).result(timeout=300):
            err = re.search(r"<desc>(.*?)</desc>", r)
            vals = re.search(r"<vals[^>]*>(.*?)</vals>", r)
            print(f"{name:22} {secs:5.2f}s  " + (f"ERROR {err.group(1)}" if err else (vals.group(1) if vals else "ok")))
    finally:
        b.submit(lambda s: DrillBrowser._close(s)).result(timeout=60)


if __name__ == "__main__":
    main()
