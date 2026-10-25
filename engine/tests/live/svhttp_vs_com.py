"""Live: every grid read over Smart View's HTTP protocol equals the COM route (hidden Excel) cell for cell.
Run from engine\\: .venv\\Scripts\\python.exe tests\\live\\svhttp_vs_com.py  (needs SSO; prints no figures).
Also reports any cell type code other than 0 (member), 2 (data) and 7 (blank) seen in the HTTP answers."""
from __future__ import annotations

import collections
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from smartview_mcp import ops, svhttp, svxml  # noqa: E402
from smartview_mcp.grid import Axis  # noqa: E402
from smartview_mcp.host import ExcelHost  # noqa: E402

PNL = ["ac.C30", "ac.C40S", "ac.C48", "ac.C50S", "ac.C60", "ac.C70", "ac.EBITDA_BOI", "ac.C73", "ac.EBITDA", "ac.C75",
       "ac.STA20Y"]
BA = ["un.SY.2WHS", "un.SY.2BUS", "un.SY.2SSM", "un.SY.2EI", "un.SY.2TEC", "un.SY.2IT", "un.SY.2ADM"]
MONTHS = ["Jan 26", "Feb 26", "Mar 26", "Apr 26", "May 26", "Jun 26", "Jul 26", "Aug 26", "Sep 26"]
ACT = {"All Scenarios": "Actual", "All Currencies": "NOK"}
CASES = {
    "FinMax 2x2 periodic": ("FinMax", {**ACT, "All Entities": "un.SY.2WHS", "All Balance Types": "Periodic"},
                            [Axis("All Accounts", ["ac.C30", "ac.EBITDA"])], [Axis("All Time Periods", MONTHS[-2:])]),
    "FinMax BA x P&L x 9 months": ("FinMax", {**ACT, "All Balance Types": "Periodic"},
                                   [Axis("All Entities", BA), Axis("All Accounts", PNL)],
                                   [Axis("All Time Periods", MONTHS)]),
    "FinMax YTD": ("FinMax", {**ACT, "All Entities": "un.SY.2WHS", "All Balance Types": "Accumulated"},
                   [Axis("All Accounts", PNL)], [Axis("All Time Periods", ["Sep 26"])]),
    "FinLite cost base": ("FinLite", {**ACT, "All Balance Types": "Periodic"},
                          [Axis("All Entities", BA)], [Axis("All Time Periods", MONTHS[-3:])]),
    "FinPlan actual vs target": ("FinPlan", {"Entity": "2WHS Wholesale", "Years": "FY26", "Period": "Sep",
                                             "All Balance Types": "Accumulated", "Version": "Final"},
                                 [Axis("Account", ["Revenue", "COGS", "OPEX", "EBITDA"])],
                                 [Axis("Scenario", ["Target"])]),
}

seen_types: collections.Counter = collections.Counter()
_cell = svxml._cell


def _counting_cell(v: str, t: str):
    seen_types[t] += 1
    return _cell(v, t)


svxml._cell = _counting_cell


def same(a, b) -> bool:
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) <= 0.005
    return a == b


def main() -> int:
    http, host, bad = svhttp.shared(), ExcelHost().ensure(), 0
    try:
        for name, (conn, pov, rows, cols) in CASES.items():
            t0 = time.monotonic()
            _, ph, _ = http.retrieve(conn, pov, rows, cols, False, False)
            th = time.monotonic() - t0
            t0 = time.monotonic()
            _, pc, _ = host.on_host(lambda: ops.query(host, conn, pov, rows, cols, False, False), timeout=600)
            tc = time.monotonic() - t0
            cells = sum(len(r) for r in pc.rows)
            diff = [(i, j) for i, (rh, rc) in enumerate(zip(ph.rows, pc.rows)) for j, (a, b) in enumerate(zip(rh, rc))
                    if not same(a, b)]
            ok = len(ph.rows) == len(pc.rows) and not diff and ph.row_labels == pc.row_labels
            bad += not ok
            print(f"{'PASS' if ok else 'FAIL'}  {name:30} {cells:6} cells  http {th:5.1f}s  com {tc:5.1f}s"
                  + ("" if ok else f"  first differing cells: {diff[:5]}"
                     + ("" if ph.row_labels == pc.row_labels else "  (row labels differ)")))
    finally:
        host.shutdown()
        from smartview_mcp import drill_browser
        drill_browser.shutdown_shared()
    odd = {t: n for t, n in seen_types.items() if t not in ("0", "2", "7")}
    print("cell type codes seen:", dict(seen_types), "- unexpected:", odd or "none")
    return 1 if bad or odd else 0


if __name__ == "__main__":
    raise SystemExit(main())
