"""Risk spike R2 (docs/specs/2026-11-02-pipeline-architecture-design.md §7): can the Smart View HTTP provider
fetch everything benchmark B15 needs (Wholesale full P&L, actual vs budget vs last year, YTD) fast, safely in
parallel, and after a long idle?

Run from engine\\:  uv run python tests\\live\\spike_r2_svhttp.py [--idle SECONDS] [--rounds N]
Needs Telenor SSO (hidden Edge). Prints shapes and timings only, never figures. Writes docs\\audit\\<date>-spike-r2.json.
Go: warm B15 fetch < 10 s; parallel = sequential cell for cell; fetch after --idle works (reconnect allowed)."""
from __future__ import annotations

import argparse
import concurrent.futures as cf
import datetime as dt
import json
import statistics
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from smartview_mcp import catalog, drill_browser, svhttp  # noqa: E402
from smartview_mcp.grid import Axis  # noqa: E402

AUDIT = Path(__file__).resolve().parents[3] / "docs" / "audit"
P, LY = "Sep 26", "Sep 25"
CUR = ["Dec 25"] + [f"{m} 26" for m in svhttp._MONTHS[:9]]
PREV = ["Dec 24"] + [f"{m} 25" for m in svhttp._MONTHS[:9]]


def b15_pulls() -> dict[str, tuple]:
    """The pulls fin_pnl makes for Wholesale at Sep 26 (fin_spec.finmax/fte_levels/finplan), as retrieve() args."""
    cat = catalog.load()
    fm, fp = cat.pnl["FinMax"], cat.pnl["FinPlan"]
    base = {"All Entities": "un.SY.2WHS", "All Scenarios": "Actual", "All Currencies": "NOK"}
    stocks = [x.account for x in fm if x.measure == "stock"]
    accts = [x.account for x in fp]
    out = {"pnl": ("FinMax", base, [Axis("All Accounts", [x.account for x in fm])],
                   [Axis("All Balance Types", ["Periodic", "Accumulated"]), Axis("All Time Periods", [P, LY])]),
           "fpa": ("FinPlan", {"Entity": "2WHS Wholesale", "Scenario": "Actual", "Version": "Working"},
                   [Axis("Account", accts)],
                   [Axis("Years", ["FY26", "FY25"]), Axis("All Balance Types", ["Periodic", "Accumulated"]),
                    Axis("Period", ["Sep", "Dec"])]),
           "fpp": ("FinPlan", {"Entity": "2WHS Wholesale", "Years": "FY26", "Version": "Final"},
                   [Axis("Account", accts)],
                   [Axis("Scenario", ["Target"]), Axis("All Balance Types", ["Periodic", "Accumulated"]),
                    Axis("Period", ["Sep", "Dec"])])}
    if stocks:
        out["fte"] = ("FinMax", {**base, "All Balance Types": "Accumulated"}, [Axis("All Accounts", stocks)],
                      [Axis("All Time Periods", CUR + PREV)])
    # not B15: a ~1 600-cell grid (7 business areas x P&L x 21 months) to see how the route scales
    ba = ["un.SY.2WHS", "un.SY.2BUS", "un.SY.2SSM", "un.SY.2EI", "un.SY.2TEC", "un.SY.2IT", "un.SY.2ADM"]
    months = [f"{m} {y}" for y in ("25", "26") for m in svhttp._MONTHS][:21]
    out["stress"] = ("FinMax", {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"},
                     [Axis("All Entities", ba), Axis("All Accounts", [x.account for x in fm if x.measure == "flow"])],
                     [Axis("All Time Periods", months)])
    return out


def fetch(http, args) -> tuple[list, float]:
    conn, pov, rows, cols = args
    t0 = time.monotonic()
    _, parsed, _ = http.retrieve(conn, pov, rows, cols, False, False)
    return parsed.rows, time.monotonic() - t0


def cells(rows) -> int:
    return sum(len(r) for r in rows)


def round_seq(http, pulls) -> tuple[float, dict]:
    t0 = time.monotonic()
    got = {k: fetch(http, a) for k, a in pulls.items()}
    return time.monotonic() - t0, got


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--idle", type=float, default=0, help="seconds to sit idle before the last fetch")
    ap.add_argument("--rounds", type=int, default=3)
    a = ap.parse_args()
    pulls, http = b15_pulls(), svhttp.shared()
    rep: dict = {"spike": "R2", "when": dt.datetime.now().isoformat(timespec="seconds"),
                 "pulls": {k: {"connection": v[0]} for k, v in pulls.items()}}
    try:
        t, ref = round_seq(http, pulls)
        rep["cold_s"] = round(t, 2)
        for k, (rows, s) in ref.items():
            rep["pulls"][k].update(cells=cells(rows), cold_s=round(s, 2))
        print(f"cold (browser + SSO + sessions): {t:.1f}s  cells "
              + ", ".join(f"{k}={cells(r)}" for k, (r, _) in ref.items()))

        warm = []
        for i in range(a.rounds):
            t, got = round_seq(http, pulls)
            same = all(got[k][0] == ref[k][0] for k in pulls)
            warm.append(t)
            print(f"warm round {i + 1}: {t:.2f}s  per pull "
                  + ", ".join(f"{k}={s:.2f}" for k, (_, s) in got.items()) + f"  same={same}")
        rep["warm_s"] = [round(x, 2) for x in warm]
        rep["warm_median_s"] = round(statistics.median(warm), 2)

        t0 = time.monotonic()
        with cf.ThreadPoolExecutor(len(pulls)) as ex:
            par = dict(zip(pulls, ex.map(lambda k: fetch(http, pulls[k]), pulls)))
        rep["parallel_s"] = round(time.monotonic() - t0, 2)
        rep["parallel_same"] = all(par[k][0] == ref[k][0] for k in pulls)
        print(f"parallel ({len(pulls)} threads): {rep['parallel_s']}s  same={rep['parallel_same']}")

        if a.idle:
            print(f"idle {a.idle:.0f}s ...", flush=True)
            time.sleep(a.idle)
            t0 = time.monotonic()
            try:
                t, got = round_seq(http, pulls)
                # new cube sessions in the same browser: if they agree with the reconnected ones, a 'differs' is
                # the cube's data moving (month-end loads), not a stale or wrong reconnect
                fresh_http = svhttp.SvHttp(svhttp.BrowserTransport(lambda: drill_browser.shared_browser()))
                _, fresh = round_seq(fresh_http, pulls)
                rep["after_idle"] = {"idle_s": a.idle, "ok": True, "s": round(t, 2),
                                     "same": all(got[k][0] == ref[k][0] for k in pulls),
                                     "differs": [k for k in pulls if got[k][0] != ref[k][0]],
                                     "matches_fresh_session": all(got[k][0] == fresh[k][0] for k in pulls)}
            except Exception as e:  # noqa: BLE001 - the spike records the failure
                rep["after_idle"] = {"idle_s": a.idle, "ok": False, "s": round(time.monotonic() - t0, 2),
                                     "error": str(e)[:300]}
            print("after idle:", rep["after_idle"])
    finally:
        drill_browser.shutdown_shared()
    total = sum(v["cells"] for v in rep["pulls"].values())
    rep["total_cells"] = total
    rep["go"] = {"warm_lt_10s": rep["warm_median_s"] < 10, "parallel_safe": rep["parallel_same"],
                 "idle_ok": rep.get("after_idle", {}).get("ok")}
    out = AUDIT / f"{dt.date.today().isoformat()}-spike-r2{'-idle' if a.idle else ''}.json"
    out.write_text(json.dumps(rep, indent=1), encoding="utf-8")
    print("go:", rep["go"], "->", out.name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
