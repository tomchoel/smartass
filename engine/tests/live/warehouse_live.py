"""Live (controller only): the closed month over HTTP, a base fill of a scratch warehouse on the real cubes with its
reconciliation, fin_pnl's engine numbers for Wholesale against direct retrieves, and timings. Read-only towards EPM.
Prints no figures (only PASS/FAIL, counts and seconds).
Run from the worktree's engine dir: PY tests\\live\\warehouse_live.py"""
from __future__ import annotations

import contextlib
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from smartview_mcp import atlas, catalog, config, extract, fin_pnl, refresh, svhttp, templates  # noqa: E402
from smartview_mcp.analytics import periods  # noqa: E402
from smartview_mcp.grid import Axis  # noqa: E402
from smartview_mcp.warehouse import Warehouse  # noqa: E402


def same(a, b) -> bool:
    return a is not None and b is not None and abs(a - b) <= 1.0


def where(problem: str) -> str:
    """A reconciliation problem without its amount."""
    return problem.split(" is off by")[0].split(" differs from")[0]


def main() -> int:
    bad = 0
    http = svhttp.shared()
    t0 = time.monotonic()
    closed = http.closed_month()
    print(f"PASS  closed month over HTTP: {closed} ({time.monotonic() - t0:.1f} s)")

    path = config.HOME / "live" / "warehouse_check.duckdb"
    path.parent.mkdir(parents=True, exist_ok=True)
    for f in path.parent.glob("warehouse_check.duckdb*"):
        f.unlink()
    cat, store, retrieve = catalog.load(), Warehouse(path), extract.route()
    with contextlib.closing(atlas.connect()) as db:
        fills = refresh.refresh(store, retrieve, cat, closed, db, templates.latest_forecast)
        base = fills[0].fill
        ok = base.result == "ok"
        bad += not ok
        print(f"{'PASS' if ok else 'FAIL'}  base fill: {base.detail['grids']} grids, {base.cells} cells, "
              f"{base.seconds:.1f} s, routes {base.detail['routes']}, {base.detail['checks']} checks, "
              f"{len(base.detail['problems'])} problems, {len(base.detail['errors'])} errors")
        for p in base.detail["problems"][:5]:
            print("        problem at", where(p))
        for e in base.detail["errors"][:5]:
            print("        error:", e[:160])
        t0 = time.monotonic()
        again = refresh.refresh(store, retrieve, cat, closed, db, templates.latest_forecast)
        ok = again == []
        bad += not ok
        print(f"{'PASS' if ok else 'FAIL'}  warm refresh check: {len(again)} jobs ({time.monotonic() - t0:.2f} s)")

        whs = catalog.resolve_entity(cat, "Wholesale")
        kids = refresh.sub_units(lambda m: [r["name"] for r in atlas.children(db, "FinMax", m)], whs.finmax)
    fc = templates.latest_forecast(periods.split_period(closed)[1])
    t0 = time.monotonic()
    tops = refresh.top_up(store, retrieve, cat, whs, closed, closed, fc, kids)
    top_s = time.monotonic() - t0
    ok = all(r.ok for r in tops)
    bad += not ok
    print(f"{'PASS' if ok else 'FAIL'}  top-up: {sum(r.fill.cells for r in tops)} cells, {top_s:.1f} s")
    t0 = time.monotonic()
    res = fin_pnl.compute(cat, fin_pnl.Scope(whs, closed, closed, fc, tuple((k, k) for k in kids)), store)
    calc_ms = (time.monotonic() - t0) * 1000
    mon, fy = periods.finplan(closed)
    fm = {"All Entities": whs.finmax, "All Scenarios": "Actual", "All Currencies": "NOK"}
    fp = {"Entity": whs.finplan, "Years": fy, "Period": mon, "All Balance Types": "Accumulated"}
    cases = {
        "EBITDA YTD (FinMax)": ("fm.ebitda.ytd.cur", "FinMax", {**fm, "All Balance Types": "Accumulated"},
                                "All Accounts", "ac.EBITDA", "All Time Periods", closed),
        "EBITDA YTD last year": ("fm.ebitda.ytd.ly", "FinMax", {**fm, "All Balance Types": "Accumulated"},
                                 "All Accounts", "ac.EBITDA", "All Time Periods", res.ly),
        "Revenue, the month": ("fm.rev.m.cur", "FinMax", {**fm, "All Balance Types": "Periodic"},
                               "All Accounts", "ac.C30", "All Time Periods", closed),
        "FTE at month end": ("fm.fte.m.cur", "FinMax", {**fm, "All Balance Types": "Accumulated"},
                             "All Accounts", "ac.STA20Y", "All Time Periods", closed),
        "FinPlan actual EBITDA YTD": ("fp.ebitda.actual.ytd", "FinPlan", {**fp, "Version": "Working"},
                                      "Account", "EBITDA", "Scenario", "Actual"),
        "Target EBITDA YTD": ("fp.ebitda.budget.ytd", "FinPlan", {**fp, "Version": "Final"},
                              "Account", "EBITDA", "Scenario", "Target"),
    }
    if fc:
        cases["Forecast EBITDA YTD"] = ("fp.ebitda.forecast.ytd", "FinPlan", {**fp, "Version": "Final"},
                                        "Account", "EBITDA", "Scenario", fc)
    for name, (fid, conn, pov, rdim, rmem, cdim, cmem) in cases.items():
        _, p, _ = http.retrieve(conn, pov, [Axis(rdim, [rmem])], [Axis(cdim, [cmem])], False, False)
        got = p.rows[0][1] if p.rows else None
        got = 0.0 if got is None else got
        ok = isinstance(got, float) and same(res.facts[fid].value, got)
        bad += not ok
        print(f"{'PASS' if ok else 'FAIL'}  {name}: the engine equals a direct retrieve")
    ok = res.line_bridge.closes() and (res.unit_bridge is None or res.unit_bridge.closes())
    bad += not ok
    print(f"{'PASS' if ok else 'FAIL'}  bridges close ({len(kids)} sub-units)")
    ok = not res.line_bridge.unexplained
    bad += not ok
    print(f"{'PASS' if ok else 'FAIL'}  line bridge has no unexplained residual (> NOK 1)")
    missing = res.missing() if callable(getattr(res, "missing", None)) else []
    print(f"info  missing facts: {len(missing)}; problems: {len(getattr(res, 'problems', []))}")
    print(f"timing: base fill {base.seconds:.1f} s, top-up {top_s:.1f} s, calc engine {calc_ms:.0f} ms")
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
