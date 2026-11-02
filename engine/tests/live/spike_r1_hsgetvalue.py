"""Risk spike R1 (docs/specs/2026-11-02-pipeline-architecture-design.md §7): can a source sheet be a grid of
Smart View HsGetValue formulas in the user's visible Excel?

Writes a ~500-cell FinMax grid (7 business areas x P&L lines x 7 months) as HsGetValue formulas that reference
their row/column labels, in one block, into a new workbook in the interactive Excel; then Smart View Refresh on
that sheet; then compares every cell with the same grid fetched over HTTP; then saves, closes and reopens the file
to see what a colleague would see without refreshing.

Run from engine\\:  uv run python tests\\live\\spike_r1_hsgetvalue.py [--keep]
Needs Excel + Smart View + SSO. Prints counts and timings only. Writes docs\\audit\\<date>-spike-r1.json.
Go: refresh < 15 s; every cell = HTTP; no dialog or manual connection step."""
from __future__ import annotations

import argparse
import collections
import datetime as dt
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from smartview_mcp import catalog, drill_browser, svhttp, userexcel  # noqa: E402
from smartview_mcp.grid import Axis  # noqa: E402
from smartview_mcp.host import ExcelHost  # noqa: E402
from smartview_mcp.hsgetvalue import FINMAX_DEFAULTS  # noqa: E402
from smartview_mcp.userexcel import com  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from helpers import close_scratch  # noqa: E402

AUDIT = Path(__file__).resolve().parents[3] / "docs" / "audit"
BA =  ["un.SY.2WHS", "un.SY.2BUS", "un.SY.2SSM", "un.SY.2EI", "un.SY.2TEC", "un.SY.2IT", "un.SY.2ADM"]
MONTHS = [f"{m} 26" for m in svhttp._MONTHS[2:9]]
POV = {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"}
TOP = 4                                                     # header row; data from row 5, members in A:B


def col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def formula(r: int, c: int, fixed: dict[str, str]) -> str:
    refs = [f'"All Entities#"&$A{r}', f'"All Accounts#"&$B{r}', f'"All Time Periods#"&{col(c)}${TOP}']
    return "=HsGetValue(" + ",".join(['"FinMax"', *refs, *(f'"{d}#{m}"' for d, m in fixed.items())]) + ")"


def kind(v) -> str:
    if isinstance(v, float):
        return "number"
    if isinstance(v, int):
        return f"excel error {v}"
    return f"text {str(v)[:24]!r}" if v not in (None, "") else "blank"


def same(a, b) -> bool:
    empty = (None, "", "#Missing", "#missing", 0, 0.0)
    if a in empty and b in empty:
        return True
    return isinstance(a, float) and isinstance(b, float) and abs(a - b) <= 0.005


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep", action="store_true", help="leave the scratch workbook")
    ap.add_argument("--connect", choices=("none", "helper", "sheet"), default="helper",
                    help="connect FinMax first: on a helper sheet, on the formula sheet itself, or not at all")
    a = ap.parse_args()
    accts = [x.account for x in catalog.load().pnl["FinMax"] if x.measure == "flow"]
    rep: dict = {"spike": "R1", "when": dt.datetime.now().isoformat(timespec="seconds")}

    http = svhttp.shared()
    t0 = time.monotonic()
    _, parsed, _ = http.retrieve("FinMax", POV, [Axis("All Entities", BA), Axis("All Accounts", accts)],
                                 [Axis("All Time Periods", MONTHS)], False, False)
    rep["http_s"] = round(time.monotonic() - t0, 2)
    dims = http.dimensions("FinMax")
    drill_browser.shutdown_shared()
    fixed = {d: m for d, m in {**FINMAX_DEFAULTS, **POV}.items()}
    rep["dims_not_in_formula"] = [d for d in dims if d not in fixed and d not in
                                  ("All Entities", "All Accounts", "All Time Periods")]
    rep["formula_dims_not_in_cube"] = [d for d in fixed if d not in dims]
    rows = [(e, ac) for e in BA for ac in accts]
    assert len(parsed.row_labels) == len(rows), "HTTP returned a different number of rows"   # labels are aliases
    n = len(rows) * len(MONTHS)
    rep["cells"] = n
    print(f"HTTP: {n} cells in {rep['http_s']}s; dims missing from formula: {rep['dims_not_in_formula']}; "
          f"formula dims not in cube: {rep['formula_dims_not_in_cube']}")

    h = ExcelHost().ensure()
    t = None
    try:
        t0 = time.monotonic()
        t = userexcel.resolve_target(h, "new", title="spike-r1")
        rep["open_workbook_s"] = round(time.monotonic() - t0, 2)
        ws = userexcel.add_sheet(t, "R1 HsGetValue")
        last = 2 + len(MONTHS)
        header = [["Entity", "Account", *MONTHS]]
        body = [[e, ac, *(formula(TOP + 1 + i, 3 + j, fixed) for j in range(len(MONTHS)))]
                for i, (e, ac) in enumerate(rows)]
        end = TOP + len(rows)
        t0 = time.monotonic()
        com(lambda: setattr(ws.Range(f"A{TOP}:{col(last)}{TOP}"), "NumberFormat", "@"))
        com(lambda: setattr(ws.Range(f"A{TOP + 1}:B{end}"), "NumberFormat", "@"))
        com(lambda: setattr(ws.Range("A1"), "Value", "R1 spike: HsGetValue source grid"))
        com(lambda: setattr(ws.Range(f"A{TOP}:{col(last)}{end}"), "Formula2", header + body))
        rep["write_s"] = round(time.monotonic() - t0, 2)
        data = f"C{TOP + 1}:{col(last)}{end}"

        def read():
            return [v for r in com(lambda: ws.Range(data).Value2) for v in r]

        before = read()
        rep["before_refresh"] = dict(collections.Counter(map(kind, before)))
        print(f"wrote {n} formulas in {rep['write_s']}s; before refresh: {rep['before_refresh']}")

        rep["connect_mode"] = a.connect
        if a.connect != "none":
            where = ws.Name
            if a.connect == "helper":
                where = userexcel.add_sheet(t, "_sv").Name
            with userexcel.SmartViewBoxes(t.pid) as boxes:
                t0 = time.monotonic()
                for attempt in range(3):            # as userexcel._pull: EPM sometimes answers a connect with junk
                    st = t.run("B_Connect", where, "FinMax")
                    if st == "OK":
                        break
                    time.sleep(4 * (attempt + 1))
                rep["connect"] = {"sheet": where, "status": st, "attempts": attempt + 1,
                                  "s": round(time.monotonic() - t0, 2), "dialogs": boxes.texts[:5]}
            print("connect:", rep["connect"])

        want = [v for r in parsed.rows for v in r[-len(MONTHS):]]          # rows carry the 2 label columns first
        for i in (1, 2):
            with userexcel.SmartViewBoxes(t.pid) as boxes:
                t0 = time.monotonic()
                st = t.run("B_RefreshSheet", ws.Name)
                userexcel.wait_calculated(t.app)
                secs = time.monotonic() - t0
            got = read()
            diff = sum(not same(g, w) for g, w in zip(got, want))
            rep[f"refresh_{i}"] = {"status": st, "s": round(secs, 2), "dialogs": boxes.texts[:5],
                                   "kinds": dict(collections.Counter(map(kind, got))), "cells_differing": diff}
            print(f"refresh {i}: {st} in {secs:.1f}s; kinds {rep[f'refresh_{i}']['kinds']}; "
                  f"differ from HTTP: {diff}; dialogs: {boxes.texts[:3]}")

        com(t.wb.Save)
        path = t.path
        com(lambda: t.wb.Close(False))
        t0 = time.monotonic()
        wb = com(lambda: t.app.Workbooks.Open(path), limit=120)
        ws = com(lambda: wb.Worksheets("R1 HsGetValue"))
        reopened = read()
        rep["reopen"] = {"s": round(time.monotonic() - t0, 2),
                         "kinds": dict(collections.Counter(map(kind, reopened))),
                         "cells_differing": sum(not same(g, w) for g, w in zip(reopened, want))}
        print(f"reopened without refresh: {rep['reopen']}")
        t.wb = wb
    finally:
        if t is not None and not a.keep:
            try:
                close_scratch(t)
            except Exception as e:  # noqa: BLE001 - cleanup only
                print("cleanup:", e)
        h.shutdown()
    r1, r2 = rep.get("refresh_1", {}), rep.get("refresh_2", {})
    rep["go"] = {"refresh_lt_15s": r1.get("s", 99) < 15, "equals_http": r1.get("cells_differing") == 0,
                 "no_dialogs": not r1.get("dialogs") and not r2.get("dialogs"), "status_ok": r1.get("status") == "OK"}
    out = AUDIT / f"{dt.date.today().isoformat()}-spike-r1-{a.connect}.json"
    out.write_text(json.dumps(rep, indent=1, default=str), encoding="utf-8")
    print("go:", rep["go"], "->", out.name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
