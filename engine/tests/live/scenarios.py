"""Live scenarios S1-S10. Each is fn(ctx) -> dict(notes=..., budget=...) and raises AssertionError/ToolFailed on failure."""
from __future__ import annotations

import json
import re
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path

import helpers

REGISTRY: dict[str, "Scenario"] = {}


@dataclass
class Scenario:
    id: str
    fn: object
    tags: list[str]
    budget: float            # seconds; over it -> SLOW
    limit: str | None = None # set when the failure is an accepted, documented limit
    fresh: bool = False      # only runs with --fresh


@dataclass
class Ctx:
    client: object           # McpClient
    scratch: object          # Path to <analyser>\_matrix
    keep: bool
    state: dict = field(default_factory=dict)   # shared between scenarios (e.g. S2's workbook for S7)
    watch: object = None     # the running helpers.DialogWatch (matrix.py sets it around each scenario)


def scenario(id: str, tags: list[str], budget: float = 60, limit: str | None = None, fresh: bool = False):
    def deco(fn):
        REGISTRY[id] = Scenario(id, fn, tags, budget, limit, fresh)
        return fn
    return deco


@scenario("S0.status", ["smoke"], budget=10)
def s0(ctx):
    b, _ = ctx.client.must("sv_status", {})
    assert b


# ================================================================== shared specs and expected values
WHS_POV = {"All Entities": "un.SY.2WHS", "All Scenarios": "Actual", "All Currencies": "NOK",
           "All Balance Types": "Accumulated"}
EBITDA_WHS_YTD = 1858190823.29            # FinMax ac.EBITDA x un.SY.2WHS x Sep 26, Accumulated
CELL_12881 = {"All Accounts": "ac.12881", "All Entities": "un.SY", "All Time Periods": "Sep 26"}
SUM_12881 = -161233427.84                 # FinLite, Periodic; the drill gives 95 lines
LINES_12881 = 95


def _q(rows, cols, pov, connection="FinMax", **kw) -> dict:
    return {"connection": connection, "rows": rows, "columns": cols, "pov": pov, **kw}


def _ebitda_q(pov=None, rows=None) -> dict:
    return _q(rows or [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
              [{"dimension": "All Time Periods", "members": ["Sep 26"]}], pov or WHS_POV)


def _close(a: float | None, b: float, tol: float = 0.01) -> bool:
    return a is not None and abs(a - b) <= tol


def _need(ctx, key: str, who: str):
    v = ctx.state.get(key)
    if not v:
        raise AssertionError(f"needs {who} (no {key} in this run: run it in the same matrix run, and it must pass)")
    return v


def _all_ok(checks: dict) -> bool:
    """Every check OK (vacuously True for {}: _answer_ok rejects empty checks unless allowed)."""
    return all(v == "OK" for v in checks.values())


ERROR_CODES = ("#NoAccess", "#No Access", "#Invalid", "#VALUE!", "#REF!", "#N/A", "#DIV/0!", "#NAME?")


def _no_formula_errors(out: dict) -> list[str]:
    """Formula errors and error cells (answer lines and data sheets: '#NoAccess', '#VALUE!' ...) in the build."""
    return [w for w in out.get("warnings", [])
            if "formula error" in w or "error cell" in w or any(c in w for c in ERROR_CODES)]


# ================================================================== S1 lookups
FIND = [("12881", "ac.12881"), ("un.SY.2TEC", "un.SY.2TEC"), ("mobil", None), ("årsverk", "ac.HR520")]


@scenario("S1.find", ["lookup"], budget=5)
def s1_find(ctx):
    bad = []
    for text, want in FIND:
        b, _ = ctx.client.must("sv_find_member", {"text": text, "limit": 10})
        hits = b.get("atlas_hits") or []
        top = [h["name"] for h in hits[:3]]
        if not hits:
            bad.append(f"{text}: no hits")
        elif want and want not in top:
            bad.append(f"{text}: {want} not in top 3 {top}")
        elif not want and not any("Products" in h.get("dimension", "") for h in hits[:3]):
            bad.append(f"{text}: no product in top 3 {[(h['dimension'], h['name']) for h in hits[:3]]}")
    assert not bad, "; ".join(bad)


@scenario("S1.query.finmax", ["lookup"], budget=5)
def s1_finmax(ctx):
    b, _ = ctx.client.must("sv_query", _ebitda_q())
    v = helpers.number(b["table"])
    assert _close(v, EBITDA_WHS_YTD), f"EBITDA WHS YTD Sep 26 = {v}, expected {EBITDA_WHS_YTD}"


@scenario("S1.query.finlite", ["lookup"], budget=5)
def s1_finlite(ctx):
    b, _ = ctx.client.must("sv_query", _q([{"dimension": "All Accounts", "members": ["ac.12881"]}],
                                          [{"dimension": "All Time Periods", "members": ["Sep 26"]}],
                                          {"All Entities": "un.SY", "All Scenarios": "Actual",
                                           "All Currencies": "NOK", "All Balance Types": "Periodic"}, "FinLite"))
    v = helpers.number(b["table"])
    assert _close(v, SUM_12881), f"FinLite 12881 un.SY Sep 26 = {v}, expected {SUM_12881}"


@scenario("S1.query.finplan", ["lookup"], budget=8)
def s1_finplan(ctx):
    b, _ = ctx.client.must("sv_query", _q([{"dimension": "Account", "members": ["EBITDA"]}],
                                          [{"dimension": "Scenario", "members": ["Actual"]}],
                                          {"Entity": "BANORWAY", "Years": "FY26", "Period": "Sep",
                                           "Version": "Working", "All Balance Types": "Accumulated"}, "FinPlan"))
    v = helpers.number(b["table"])
    assert v is not None, f"no number: {b['table']} {b['warnings']}"
    real = [w for w in b["warnings"] if not w.startswith("HTTP route failed")]   # a fallback note is not a data warning
    assert not real, real


def _corrected_or_named(ctx, args: dict, name: str):
    """Bad input passes: corrected with a warning naming `name`, or rejected with a message naming it."""
    b, _, err = ctx.client.call("sv_query", args)
    if err:
        assert name in err, f"rejected without naming '{name}': {err[:300]}"
        return {"notes": "rejected, names " + name}
    warned = [w for w in b.get("warnings", []) if name in w]
    assert warned, f"accepted without a warning naming '{name}': warnings={b.get('warnings')}"
    assert _close(helpers.number(b["table"]), EBITDA_WHS_YTD), f"corrected, but wrong value: {b['table']}"
    return {"notes": "corrected with a warning"}


@scenario("S1.bad.balance", ["lookup", "input"], budget=5)
def s1_bad_balance(ctx):
    pov = {k: v for k, v in WHS_POV.items() if k != "All Balance Types"} | {"Balance": "Accumulated"}
    return _corrected_or_named(ctx, _ebitda_q(pov), "All Balance Types")


@scenario("S1.bad.quoted", ["lookup", "input"], budget=5)
def s1_bad_quoted(ctx):
    pov = {k: v for k, v in WHS_POV.items() if k != "All Balance Types"} | {"'All Balance Types'": "Accumulated"}
    return _corrected_or_named(ctx, _ebitda_q(pov), "All Balance Types")


@scenario("S1.bad.dictrows", ["lookup", "input"], budget=5)
def s1_bad_dictrows(ctx):
    args = _ebitda_q()
    args["rows"] = {"All Accounts": ["ac.EBITDA"]}
    b, _ = ctx.client.must("sv_query", args)
    assert _close(helpers.number(b["table"]), EBITDA_WHS_YTD), b["table"]


@scenario("S1.bad.member", ["lookup", "input"], budget=5)
def s1_bad_member(ctx):
    b, _, err = ctx.client.call("sv_query", _ebitda_q(rows=[{"dimension": "All Accounts",
                                                             "members": ["ac.EBITDA", "ac.NOPE"]}]))
    text = err or " ".join(b.get("warnings", []))
    assert "ac.NOPE" in text and "sv_find_member" in text, f"error/warning does not name ac.NOPE + sv_find_member: {text}"


# ================================================================== S2 free-form build (live_build.py's BUS spec)
E = ["ac.EBITDA"]
BUS_ACC = ["ac.C30", "ac.C40S", "ac.C70", "ac.C73", "ac.EBITDA"]
BUS_POV = {"All Entities": "un.SY.2BUS", "All Scenarios": "Actual", "All Currencies": "NOK"}
MONTHS = [f"{m} 25" for m in "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()] + \
         [f"{m} 26" for m in "Jan Feb Mar Apr May Jun Jul Aug Sep".split()]
BUS_YTD, BUS_VAR = "3 496 062 985", "89 346 421"     # sv_query 23.10: 3,496,062,984.87 (live_build.py had 3 495 317 868)
FP_POV = {"Entity": "2BUS Business", "Years": "FY26", "Period": "Sep", "Version": "Working",
          "All Balance Types": "Accumulated"}


def _bus_pulls() -> list[dict]:
    return [{"id": "m", "connection": "FinMax", "rows": [{"dimension": "All Accounts", "members": BUS_ACC}],
             "columns": [{"dimension": "All Time Periods", "members": MONTHS}],
             "pov": {**BUS_POV, "All Balance Types": "Periodic"}},
            {"id": "ytd", "connection": "FinMax", "rows": [{"dimension": "All Accounts", "members": BUS_ACC}],
             "columns": [{"dimension": "All Time Periods", "members": ["Sep 26", "Sep 25", "Dec 25"]}],
             "pov": {**BUS_POV, "All Balance Types": "Accumulated"}, "title": "Q1 FinMax YTD"}]


def _bus_spec(m: str = "m", ytd: str = "ytd") -> dict:
    return {
        "title": "Business Norway — EBITDA per month 2026 vs 2025", "subtitle": "NOK · FinMax · Actual · un.SY.2BUS",
        "key_points": ["June carries the year (one-off low opex)."],
        "columns": [{"header": "Actual 2026"}, {"header": "Actual 2025"}, {"header": "Variance"},
                    {"header": "Variance %", "kind": "prosent1"}],
        "sections": [
            {"title": "EBITDA per month (Periodic)",
             "repeat": {"over": "months:Jan..Dec", "line": {"key": "m{i}", "label": "{v}", "cells": [
                 {"pull": m, "row": E, "col": ["{v} 26"]}, {"pull": m, "row": E, "col": ["{v} 25"]},
                 {"var": [0, 1]}, {"varpct": [0, 1]}]}},
             "lines": [{"key": "ytd", "label": "YTD Jan–Sep", "kind": "grand",
                        "cells": [{"sum": "m[0-8]"}, {"sum": "m[0-8]"}, {"var": [0, 1]}, {"varpct": [0, 1]}]},
                       {"key": "fy", "label": "Full year 2025", "kind": "sum",
                        "cells": [None, {"sum": "m*"}, None, None]}]},
            {"title": "Reconciliation (Accumulated)", "lines": [
                {"key": "eacc", "label": "EBITDA YTD Sep", "cells": [{"pull": ytd, "row": E, "col": ["Sep 26"]},
                                                                    {"pull": ytd, "row": E, "col": ["Sep 25"]},
                                                                    None, None]},
                {"key": "fyacc", "label": "EBITDA Dec 25",
                 "cells": [None, {"pull": ytd, "row": E, "col": ["Dec 25"]}, None, None]}]}],
        "checks": [{"label": "Sum of months = Accumulated YTD", "expr": "ytd-eacc", "columns": [0, 1]},
                   {"label": "Sum of 2025 months = Accumulated Dec 25", "expr": "fy-fyacc", "columns": [1]}],
        "notes": ["Live matrix build."], "source": "Source: Oracle EPM Cloud (FinMax) via Smart View.",
        "chart": {"type": "column", "lines": "m*", "columns": [0, 1], "title": "EBITDA per month"}}


def _answer_ok(out: dict, what: str, allow_empty_checks: bool = False) -> None:
    checks = out.get("checks") or {}
    assert checks or allow_empty_checks, f"{what}: no checks in the answer"
    assert _all_ok(checks), f"{what}: checks {out.get('checks')}"
    errs = _no_formula_errors(out)
    assert not errs, f"{what}: {errs}"


@scenario("S2.build", ["build"], budget=30)
def s2_build(ctx):
    out, _ = ctx.client.must("sv_ws_build", {"question": "MATRIX S2 BUS EBITDA monthly", "pulls": _bus_pulls(),
                                             "spec": _bus_spec(), "target": "new"})
    _answer_ok(out, "build")
    assert out["formula_style"].get("ok"), f"formula_style {out['formula_style']}"
    assert BUS_YTD in out["table"] and BUS_VAR in out["table"], f"{BUS_YTD} / {BUS_VAR} not in table:\n{out['table']}"
    ctx.state.update(wb=out["workbook"], q=int(str(out["question"]).lstrip("Q")),
                     answer_sheet=out["answer_sheet"], data_sheets=out["data_sheets"])
    return {"notes": f"{out['answer_sheet']} + {out['data_sheets']}"}


@scenario("S2.patch", ["build"], budget=15)
def s2_patch(ctx):
    wb = _need(ctx, "wb", "S2.build")
    out, _ = ctx.client.must("sv_ws_answer", {"q": ctx.state["q"], "target": wb,
                                              "spec_patch": {"notes": ["Patched follow-up."]}})
    _answer_ok(out, "patch")
    assert BUS_YTD in out["table"], f"{BUS_YTD} gone after the patch:\n{out['table']}"


@scenario("S2.plan", ["build"], budget=45)
def s2_plan(ctx):
    """Same Q again with a FinPlan pull added; the FinMax pulls stay (their sheets are referenced by name)."""
    wb, q = _need(ctx, "wb", "S2.build"), ctx.state["q"]
    spec = _bus_spec(f"Q{q} FinMax", f"Q{q} FinMax YTD")
    spec["sections"][1]["lines"].append(
        {"key": "fp", "label": "EBITDA YTD Sep (FinPlan Actual)",
         "cells": [{"pull": "fp", "row": ["EBITDA"], "col": ["Actual"]}, None, None, None]})
    pulls = [{"id": "fp", "connection": "FinPlan", "rows": [{"dimension": "Account", "members": ["EBITDA"]}],
              "columns": [{"dimension": "Scenario", "members": ["Actual"]}], "pov": FP_POV,
              "title": f"Q{q} FinPlan"}]
    out, _ = ctx.client.must("sv_ws_build", {"question": "MATRIX S2 BUS EBITDA monthly + FinPlan", "pulls": pulls,
                                             "spec": spec, "target": wb, "q": q})
    _answer_ok(out, "plan")
    ov, _ = ctx.client.must("xl_overview", {"workbook": wb, "preview_rows": 0})
    names = [s["name"] for s in ov["sheets"]]
    ans = out["answer_sheet"]
    data = [n for n in names if n.startswith(f"Q{q} ") and n != ans]
    used = next(s["used"] for s in ov["sheets"] if s["name"] == ans)
    f = helpers.values(ctx.client, wb, ans, used.split("!")[-1].replace("$", ""), "formulas")["formulas"]
    text = " ".join(f.values())
    orphans = [n for n in data if f"'{n}'!" not in text and f"{n}!" not in text]
    assert not orphans, f"data sheets not referenced by {ans}: {orphans} (sheets {names})"
    late = [n for n in data if names.index(n) < names.index(ans)]
    assert not late, f"data sheets before the answer {ans}: {late} (order {names})"
    return {"notes": f"sheets {names}"}


# ================================================================== S3 templates (live_analysis.py)
TEMPLATES = {
    "division_pnl": ({"entity": "un.SY.2TEC", "period": "Sep 26"}, {"entity": "un.SY", "period": "Sep 26"}),
    "monthly_trend": ({"entity": "un.SY.2TEC", "through": "Sep 26"}, {"entity": "un.SY", "through": "Sep 26"}),
    "actual_vs_plan": ({"entity": "2TEC Technology", "month": "Sep", "year": "FY26"},
                       {"entity": "BANORWAY", "month": "Sep", "year": "FY26"}),
    "cost_base": ({"parent": "un.SY.2TEC", "period": "Sep 26", "top": 15},
                  {"parent": "un.SY", "period": "Sep 26", "top": 15}),
}
TEC_PNL_EBITDA = "(1 112 834 010)"


def _template(name: str, which: str):
    params = TEMPLATES[name][0 if which == "tec" else 1]

    def run(ctx):
        target = ctx.state.get("s3_wb") or "new"
        out, _ = ctx.client.must("sv_analyse", {"question": f"MATRIX S3 {name} {which}", "template": name,
                                                "params": params, "target": target})
        ctx.state["s3_wb"] = out["workbook"]
        _answer_ok(out, f"{name}.{which}", allow_empty_checks=name == "cost_base")
        if name == "division_pnl" and which == "tec":
            assert TEC_PNL_EBITDA in out["table"], f"{TEC_PNL_EBITDA} not in the TEC P&L:\n{out['table']}"
        return {"notes": out["answer_sheet"]}
    return run


for _t in TEMPLATES:
    for _w in ("tec", "sy"):
        scenario(f"S3.{_t}.{_w}", ["template"], budget=60)(_template(_t, _w))


# ================================================================== S4 concepts
@scenario("S4.scope", ["concept"], budget=20)
def s4_scope(ctx):
    sc, _ = ctx.client.must("sv_scope", {"text": "hva tjener vi på mobilabonnement"})
    ids = [c["id"] for c in sc.get("concepts", [])]
    assert ids and ids[0].startswith("fam_"), f"no fam_* concept first: {ids}"
    out, _ = ctx.client.must("sv_concept", {"concept": ids[0], "period": "Sep 26", "target": "new"})
    _answer_ok(out, ids[0])
    assert out["checks"], f"{ids[0]}: no checks (expected rev+cogs+opex = EBITDA before other items)"
    out2, _ = ctx.client.must("sv_concept", {"concept": "splitt", "period": "Sep 26", "target": out["workbook"]})
    _answer_ok(out2, "splitt")
    return {"notes": f"{ids[0]}: {list(out['checks'])}"}


# ================================================================== S5 drill (FinLite 12881 x un.SY x Sep 26)
def _drill(ctx, cell: dict, target: str = "none", **kw) -> dict:
    out, _ = ctx.client.must("sv_drill", {"question": "MATRIX S5 drill " + " x ".join(cell.values()),
                                          "cell": cell, "target": target, **kw})
    return out


def _exact(out: dict, lines: int | None = LINES_12881, total: float | None = SUM_12881) -> dict:
    rec = out["reconciliation"]
    assert rec["status"] == "exact" and abs(rec["gap"] or 0) < 0.005, f"not exact: {rec} {out.get('warnings')}"
    if lines is not None:
        assert rec["lines"] == lines, f"{rec['lines']} lines, expected {lines}"
    if total is not None:
        assert _close(rec["lines_sum"], total), f"lines sum {rec['lines_sum']}, expected {total}"
    return rec


@scenario("S5.drill.cold", ["drill"], budget=120)
def s5_cold(ctx):
    killed = helpers.kill_engine_edge()
    rec = _exact(_drill(ctx, CELL_12881))
    return {"notes": f"killed {len(killed)} engine Edge; {rec['lines']} lines"}


@scenario("S5.drill.warm", ["drill"], budget=60)
def s5_warm(ctx):
    _exact(_drill(ctx, CELL_12881))


@scenario("S5.drill.quarter", ["drill"], budget=180)
def s5_quarter(ctx):
    out = _drill(ctx, {**CELL_12881, "All Time Periods": "Q3 26"})
    rec = _exact(out, lines=None, total=None)
    assert not out.get("split_check"), f"split pieces do not add up to the cell: {out.get('split_check')}"
    return {"notes": f"{rec['lines']} lines, {len(out.get('pieces') or [])} pieces"}


@scenario("S5.drill.report", ["drill"], budget=90)
def s5_report(ctx):
    wb = _need(ctx, "wb", "S2.build")
    d = _drill(ctx, CELL_12881, target=wb)
    _exact(d)
    cp = d["cube_pull"]
    spec = {"title": "Account 12881 un.SY Sep 26 – drill-through", "subtitle": "FinLite · Periodic · NOK",
            "columns": [{"header": "NOK"}, {"header": "Lines"}],
            "sections": [{"title": "Reconciliation", "lines": [
                {"key": "cube", "label": "Cube value", "cells": [{"pull": cp["sheet"], "row": cp["row_keys"][0],
                                                                  "col": cp["col_keys"][0]}, None]},
                {"key": "sum", "label": "Sum of lines", "cells": [{"drill": "d", "where": {}},
                                                                  {"drill": "d", "where": {}, "agg": "count"}]},
                {"key": "gap", "label": "Gap", "kind": "grand", "cells": [{"expr": "cube-sum"}, None]},
                {"key": "debit", "label": "Debit", "cells": [{"drill": "d", "where": {}, "agg": "debit"}, None]},
                {"key": "credit", "label": "Credit", "cells": [{"drill": "d", "where": {}, "agg": "credit"}, None]}]}],
            "checks": [{"label": "Gap is zero", "expr": "gap", "columns": [0]},
                       {"label": "Debit + credit = sum", "expr": "debit+credit-sum", "columns": [0]}]}
    out, _ = ctx.client.must("sv_ws_answer", {"q": d["q"], "spec": spec, "target": wb})
    _answer_ok(out, "drill report")
    return {"notes": f"Q{d['q']} {out['answer_sheet']}"}


# ================================================================== S6 xl_* round trip
def _formats_per_cell(c, wb: str, sheet: str, rng: str) -> dict[str, tuple]:
    """{address: (number format, style name)} read one cell at a time (the per-cell truth the revert is held to)."""
    m = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", rng)
    c1, r1, c2, r2 = _col(m.group(1)), int(m.group(2)), _col(m.group(3)), int(m.group(4))
    out = {}
    for r in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            a = f"{_letters(col)}{r}"
            f = helpers.values(c, wb, sheet, a, "formats")
            out[a] = (f["number_formats"][0][0], next(iter(f.get("styles") or {}), None))
    return out


@scenario("S6.xl", ["xl"], budget=60)
def s6_xl(ctx):
    c = ctx.client
    p, _ = c.must("sv_ws_pull", {"question": "MATRIX S6 xl test book", "connection": "FinMax", "target": "new",
                                 "rows": [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
                                 "columns": [{"dimension": "All Time Periods", "members": ["Sep 26"]}],
                                 "pov": WHS_POV})
    wb = p["workbook"]
    path = Path(wb)
    mtime = path.stat().st_mtime if path.exists() else None
    sheet, rng = "XL test", "B2:G8"
    add, _ = c.must("xl_add_sheet", {"name": sheet, "workbook": wb})
    base, _ = c.must("xl_write", {"range": "B2", "sheet": sheet, "workbook": wb, "values": [
        ["Line", "2026", "2025", "Var", "Var %"], ["Revenue", 100, 80, None, None], ["COGS", -30, -20, None, None],
        ["Opex", -20, -25, None, None], ["Total", None, None, None, None]]})
    snap = {"both": helpers.values(c, wb, sheet, rng, "both"), "formats": _formats_per_cell(c, wb, sheet, rng)}
    nf = helpers.values(c, wb, sheet, rng, "formats")["number_formats"]
    assert (len(nf), len(nf[0])) == (7, 6), f"xl_read formats is not the range's shape: {nf}"
    done = []
    w1, _ = c.must("xl_write", {"range": "C6:D6", "sheet": sheet, "workbook": wb, "formula": "=SUM(C3:C5)"})
    done.append(w1["change_id"])
    r = helpers.values(c, wb, sheet, "C6:D6", "both")
    assert r["values"][0] == [50, 35] and r["formulas"].get("D6") == "=SUM(D3:D5)", f"formula fill: {r}"
    w2, _ = c.must("xl_write", {"range": "E3:E6", "sheet": sheet, "workbook": wb, "formula": "=C3-D3"})
    done.append(w2["change_id"])
    fs, _ = c.must("xl_format_sheet", {"rows": {"2": "header", "3-5": "line", "6": "sum"}, "last_col": "F",
                                       "sheet": sheet, "workbook": wb})
    done.append(fs["change_id"])
    ck, _ = c.must("xl_check", {"sheet": sheet, "workbook": wb})
    tr, _ = c.must("xl_trace", {"cell": "C6", "sheet": sheet, "workbook": wb})
    assert [x["address"] for x in tr.get("precedents", [])] == ["C3:C5"], f"trace C6: {tr}"
    c.must("xl_overview", {"workbook": wb, "sheet": sheet})
    chs, _ = c.must("xl_changes", {"workbook": wb, "limit": 20})
    assert chs, "xl_changes(workbook=<full path>) returns [] (path vs OneDrive URL)"
    ids = [x["id"] for x in chs]
    missing = [i for i in done + [base["change_id"], add["change_id"]] if i not in ids]
    assert not missing, f"xl_changes misses {missing}: {ids}"
    for cid in reversed(done):
        c.must("xl_revert", {"change_id": cid, "workbook": wb})
    after = {"both": helpers.values(c, wb, sheet, rng, "both"), "formats": _formats_per_cell(c, wb, sheet, rng)}
    diff = [k for k in ("values", "formulas") if after["both"].get(k) != snap["both"].get(k)]
    diff += [f"format {a}: {snap['formats'][a]} -> {after['formats'][a]}"
             for a in snap["formats"] if after["formats"][a] != snap["formats"][a]]
    assert not diff, f"after revert differs from the snapshot in {diff}"
    c.must("xl_revert", {"change_id": base["change_id"], "workbook": wb})
    c.must("xl_revert", {"change_id": add["change_id"], "workbook": wb})
    ov, _ = c.must("xl_overview", {"workbook": wb, "preview_rows": 0})
    assert sheet not in [s["name"] for s in ov["sheets"]], "the test sheet is still there after reverting add_sheet"
    if mtime is not None:
        assert path.stat().st_mtime == mtime, "the workbook file was saved by an xl_* call"
    return {"notes": f"xl_check {ck.get('counts', ck.get('ok'))}"}


# ================================================================== S7 refresh of S2's workbook
@scenario("S7.refresh", ["refresh"], budget=90)
def s7_refresh(ctx):
    import openpyxl

    wb = _need(ctx, "wb", "S2.build")
    ans = ctx.state["answer_sheet"]
    ov, _ = ctx.client.must("xl_overview", {"workbook": wb, "sheet": ans, "preview_rows": 0})
    used = ov["sheets"][0]["used"].split("!")[-1].replace("$", "")
    before = helpers.values(ctx.client, wb, ans, used)["values"]
    assert helpers.close_scratch_book(wb, save=True), f"could not close {wb} (needs it open in Excel, in _matrix)"
    ctx.state.pop("wb", None)              # closed now: later scenarios must not target it
    out, _, err = ctx.client.call("sv_refresh_workbook", {"path": wb, "output_name": "matrix_s7"})
    assert not err, err
    ctx.state.setdefault("files", []).append(out["output"])
    assert "Call was rejected" not in str(out), f"'Call was rejected' in the result: {out}"
    book = openpyxl.load_workbook(out["output"], data_only=True, read_only=True)
    try:
        after = [list(r) for r in book[ans].iter_rows(min_row=1, values_only=True)]
    finally:
        book.close()               # read-only keeps the file open; cleanup must be able to delete it
    r0 = int(re.match(r"[A-Z]+(\d+)", used).group(1))
    c0 = _col(re.match(r"([A-Z]+)", used).group(1))
    diff = []
    for i, row in enumerate(before):
        for j, v in enumerate(row):
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                got = after[r0 - 1 + i][c0 - 1 + j] if r0 - 1 + i < len(after) else None
                if not isinstance(got, (int, float)) or abs(got - v) > 0.01:
                    diff.append((r0 + i, c0 + j, v, got))
    assert not diff, f"{len(diff)} numbers changed by the refresh, first {diff[:5]}"
    return {"notes": out["output"]}


def _col(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _letters(col: int) -> str:
    s = ""
    while col:
        col, rem = divmod(col - 1, 26)
        s = chr(65 + rem) + s
    return s


# ================================================================== S8 bad-condition injections
HANG_S, CALL_TIMEOUT = 120, 130


class HANG(Exception):
    """A call still running at HANG_S (error class 'HANG: ...')."""


def _call(client, tool: str, args: dict, hang: float = HANG_S, timeout: float = CALL_TIMEOUT):
    """client.call with the S8 budget: (body, secs, err); raises HANG at `hang` s or on the client timeout."""
    try:
        body, secs, err = client.call(tool, args, timeout=timeout)
    except TimeoutError:
        raise HANG(f"{tool} still running at {timeout:.0f}s") from None
    if secs >= hang:
        raise HANG(f"{tool} took {secs:.0f}s" + (f" and failed: {err[:200]}" if err else ""))
    return body, secs, err


def _in_thread(fn, *a) -> tuple[threading.Thread, dict]:
    box: dict = {}

    def run():
        try:
            box["out"] = fn(*a)
        except BaseException as e:      # noqa: BLE001 - re-raised by _join
            box["exc"] = e
    t = threading.Thread(target=run, daemon=True)
    t.start()
    return t, box


def _join(t: threading.Thread, box: dict, timeout: float):
    t.join(timeout)
    if t.is_alive():
        raise HANG(f"call thread still running after {timeout:.0f}s")
    if "exc" in box:
        raise box["exc"]
    return box["out"]


def _followup(ctx) -> str:
    """The S1.query.finmax check: the engine still answers afterwards."""
    b, secs, err = _call(ctx.client, "sv_query", _ebitda_q())
    assert not err, f"follow-up sv_query failed: {err[:300]}"
    v = helpers.number(b["table"])
    assert _close(v, EBITDA_WHS_YTD), f"follow-up EBITDA WHS YTD Sep 26 = {v}, expected {EBITDA_WHS_YTD}"
    return f"follow-up {secs:.0f}s"


def _small_build(client, question: str, pov: dict | None = None, hang: float = HANG_S,
                 timeout: float = CALL_TIMEOUT) -> dict:
    """A one-line sv_ws_build into a new scratch workbook (EBITDA YTD Sep 26 of the given entity)."""
    pull = {"id": "p", "connection": "FinMax", "rows": [{"dimension": "All Accounts", "members": E}],
            "columns": [{"dimension": "All Time Periods", "members": ["Sep 26"]}], "pov": pov or WHS_POV}
    spec = {"title": question, "subtitle": "NOK · FinMax · Actual", "columns": [{"header": "YTD Sep 26"}],
            "sections": [{"title": "EBITDA", "lines": [{"key": "e", "label": "EBITDA YTD",
                                                       "cells": [{"pull": "p", "row": E, "col": ["Sep 26"]}]}]}],
            "checks": [{"label": "Line equals itself", "expr": "e-e", "columns": [0]}]}
    out, _, err = _call(client, "sv_ws_build", {"question": question, "pulls": [pull], "spec": spec, "target": "new"},
                        hang, timeout)
    assert not err, f"sv_ws_build '{question}': {err[:300]}"
    _answer_ok(out, question)
    return out


def _s8_book(ctx) -> dict:
    """One small scratch workbook (with a stored answer spec) shared by S8.dialog and S8.editmode."""
    if not ctx.state.get("s8_book"):
        out = _small_build(ctx.client, "MATRIX S8 scratch")
        ctx.state["s8_book"] = {"wb": out["workbook"], "q": int(str(out["question"]).lstrip("Q")),
                                "sheet": out["answer_sheet"]}
    return ctx.state["s8_book"]


def _names_dialog(err: str | None, *words: str) -> bool:
    return bool(err) and any(w in err.lower() for w in words)


@scenario("S8.sso", ["inject"], budget=150)
def s8_sso(ctx):
    notes = []
    with helpers.expired_sso() as moved:
        notes.append(f"moved {len(moved)} cookie file(s)")
        b, secs, err = _call(ctx.client, "sv_query", _ebitda_q())
        assert not err or _names_dialog(err, "login", "sign in", "sso"), f"sv_query: {err[:300]}"
        notes.append(f"sv_query {secs:.0f}s " + ("ok" if not err else "names login"))
        d, secs, err = _call(ctx.client, "sv_drill", {"question": "MATRIX S8 sso drill", "cell": CELL_12881,
                                                      "target": "none"})
        assert not err or _names_dialog(err, "login", "sign in", "sso"), f"sv_drill without cookies: {err[:300]}"
        notes.append(f"drill {secs:.0f}s " + ("ok" if not err else "names login"))
    notes.append(_followup(ctx))
    return {"notes": "; ".join(notes)}


def _blocked_answer(ctx, inject, words: tuple[str, ...]):
    bk = _s8_book(ctx)
    with helpers.scratch_excel(bk["wb"], ctx.watch) as pid:
        with inject(pid, bk) as st:
            b, secs, err = _call(ctx.client, "sv_ws_answer", {"q": bk["q"], "target": bk["wb"],
                                                              "spec_patch": {"notes": ["S8 injection."]}})
        assert err, f"sv_ws_answer passed ({secs:.0f}s) although Excel was blocked ({st})"
        assert _names_dialog(err, *words), f"sv_ws_answer ({secs:.0f}s) does not name {words}: {err[:300]}"
        note = f"pid {pid}; {st}; sv_ws_answer {secs:.0f}s: {err.splitlines()[0][:120]}"
    return {"notes": note + "; " + _followup(ctx)}


@scenario("S8.dialog", ["inject"], budget=60)
def s8_dialog(ctx):
    return _blocked_answer(ctx, lambda pid, bk: helpers.modal_dialog(pid, ctx.watch), ("dialog",))


@scenario("S8.editmode", ["inject"], budget=60)
def s8_editmode(ctx):
    return _blocked_answer(ctx, lambda pid, bk: helpers.edit_mode(pid, bk["sheet"]), ("dialog", "edit"))


@scenario("S8.parallel", ["inject"], budget=120)
def s8_parallel(ctx):
    from mcpclient import McpClient

    second = McpClient(cwd=ctx.scratch)
    try:
        with helpers.tracked(second):            # its hidden Excel is the matrix's too (DialogWatch)
            jobs = [(ctx.client, "MATRIX S8 parallel A WHS", WHS_POV, "1 858 190 823"),
                    (second, "MATRIX S8 parallel B BUS", {**BUS_POV, "All Balance Types": "Accumulated"}, BUS_YTD)]
            threads = [_in_thread(_small_build, c, q, pov) for c, q, pov, _ in jobs]
            outs, errs = [], []
            for (t, box), (_, q, _, want) in zip(threads, jobs):
                try:
                    out = _join(t, box, CALL_TIMEOUT + 10)
                    assert want in out["table"], f"{q}: {want} not in\n{out['table']}"
                    outs.append(out)
                except Exception as e:
                    errs.append(f"{q}: {type(e).__name__}: {e}")
            if any(e.split(": ", 2)[1] == "HANG" for e in errs):
                raise HANG("; ".join(errs))
            assert not errs, "; ".join(errs)
            assert outs[0]["workbook"] != outs[1]["workbook"], f"both answers went to {outs[0]['workbook']}"
    finally:
        second.close()
    return {"notes": f"{[Path(o['workbook']).name for o in outs]}; " + _followup(ctx)}


@scenario("S8.kill.excel", ["inject"], budget=90)
def s8_kill_excel(ctx):
    st, _, err = _call(ctx.client, "sv_status", {})
    assert not err, err
    if not st.get("pid"):
        _followup(ctx)                       # start the engine's Excel first
        st, _, _ = _call(ctx.client, "sv_status", {})
    pid = helpers.engine_hidden_excel_pid(st["pid"])
    import psutil
    big = _q([{"dimension": "All Accounts", "members": BUS_ACC}],
             [{"dimension": "All Time Periods", "members": MONTHS}], {**BUS_POV, "All Balance Types": "Periodic"})
    t, box = _in_thread(_call, ctx.client, "sv_query", big)
    time.sleep(1.0)
    if not t.is_alive():
        _join(t, box, 5)
        raise helpers.Limit("injection missed: the query finished before the kill")
    try:
        psutil.Process(pid).kill()
    except psutil.Error as e:
        _join(t, box, CALL_TIMEOUT + 10)
        raise helpers.Limit(f"injection missed: Excel {pid} not killed ({type(e).__name__})")
    b, secs, err = _join(t, box, CALL_TIMEOUT + 10)
    notes = [f"killed Excel {pid} during the query"]
    if err:
        if "sv_restart" in err:
            notes.append(f"error names sv_restart ({secs:.0f}s)")
        else:
            b2, secs2, err2 = _call(ctx.client, "sv_query", big)
            assert not err2, f"error without sv_restart ({err[:200]}) and the retry failed too: {err2[:200]}"
            notes.append(f"error without sv_restart, retry ok in {secs2:.0f}s: {err.splitlines()[0][:100]}")
    else:
        notes.append(f"query ok in {secs:.0f}s")
    notes.append(_followup(ctx))
    return {"notes": "; ".join(notes)}


@scenario("S8.kill.edge", ["inject"], budget=150)
def s8_kill_edge(ctx):
    args = {"question": "MATRIX S8 kill edge", "cell": CELL_12881, "target": "none"}
    t, box = _in_thread(_call, ctx.client, "sv_drill", args)
    time.sleep(2)                                # the drill is under way (not just an idle Edge from before)
    killed, running = helpers.kill_engine_edge_when_seen(lambda: not t.is_alive(), timeout=60)
    out, secs, err = _join(t, box, CALL_TIMEOUT + 10)
    if not killed or not running:
        raise helpers.Limit("injection missed: " + ("the drill finished before an engine Edge appeared"
                                                    if not running else "no engine Edge appeared within 60s"))
    notes = [f"killed {len(killed)} engine Edge during the drill"]
    if err:
        assert "next" in err.lower(), f"drill error without a next step ({secs:.0f}s): {err[:300]}"
        notes.append(f"error with a next step ({secs:.0f}s)")
    else:
        rec = _exact(out)
        notes.append(f"drill ok in {secs:.0f}s, {rec['lines']} lines")
    notes.append(_followup(ctx))
    return {"notes": "; ".join(notes)}


@scenario("S8.cold", ["inject"], budget=240)
def s8_cold(ctx):
    wb = ctx.state.pop("wb", None)               # S2's workbook when S2 ran; closed below, so S2's users stop here
    if not wb:
        wb = _small_build(ctx.client, "MATRIX S8 cold")["workbook"]
    assert helpers.close_scratch_book(wb, save=True), f"could not close {wb}"
    out, secs, err = _call(ctx.client, "sv_drill", {"question": "MATRIX S8 cold drill", "cell": CELL_12881,
                                                    "target": wb}, hang=240, timeout=250)
    assert not err, f"sv_drill into the closed {Path(wb).name} ({secs:.0f}s): {err[:300]}"
    rec = _exact(out)
    return {"notes": f"{Path(wb).name} reopened, drill {secs:.0f}s, {rec['lines']} lines; " + _followup(ctx)}


@scenario("S8.sleep", ["inject"], budget=230)
def s8_sleep(ctx):
    pids = helpers.server_pids(ctx.client)
    assert pids, "no child_server.py process under the relay"
    t, box = _in_thread(_call, ctx.client, "sv_query", _ebitda_q(), 90 + HANG_S, 90 + CALL_TIMEOUT)
    time.sleep(0.2)
    with helpers.suspended(pids):
        time.sleep(90)
    resumed = time.monotonic()
    b, secs, err = _join(t, box, CALL_TIMEOUT + 10)
    after = time.monotonic() - resumed
    assert after < HANG_S, f"sv_query answered {after:.0f}s after the resume"
    assert not err, f"sv_query after the resume ({after:.0f}s): {err[:300]}"
    assert _close(helpers.number(b["table"]), EBITDA_WHS_YTD), b["table"]
    return {"notes": f"suspended {pids} 90s; answered {after:.0f}s after the resume; " + _followup(ctx)}


# ================================================================== S9 fresh install (only with --fresh)
DOCTOR_CHECKS = {"config", "excel", "smartview", "epm", "drill_browser", "atlas", "catalog", "output", "http_api", "update"}


def _fresh(ctx) -> dict:
    return _need(ctx, "fresh", "S9.install")


def _fresh_doctor(env: dict, wait: float = 30) -> tuple[dict, str]:
    """`smartass doctor` of the fresh install; re-run for up to `wait` s while a check is not OK (the `serve` the
    installer started may still be binding its port)."""
    t0 = time.monotonic()
    while True:
        r = helpers.fresh_cli(env, "doctor", timeout=120)
        lines = helpers.doctor_lines(r.stdout)
        if (lines and all(s == "OK" for s, _ in lines.values())) or time.monotonic() - t0 > wait:
            return lines, r.stdout + r.stderr
        time.sleep(3)


@scenario("S9.install", ["fresh"], budget=400, fresh=True)
def s9_install(ctx):
    ctx.state.pop("fresh", None)
    env = helpers.fresh_profile()
    p = helpers.fresh_paths()
    t0 = time.monotonic()
    r = helpers.fresh_install(env)
    secs = time.monotonic() - t0
    assert r.returncode == 0, f"installer exit {r.returncode}: {(r.stdout + r.stderr)[-1500:]}"
    assert p["py"].exists(), f"no venv python at {p['py']}"
    ctx.state["fresh"] = {"env": env, "owners": {}}
    lines, text = _fresh_doctor(env)
    assert DOCTOR_CHECKS <= set(lines), f"doctor lines missing {sorted(DOCTOR_CHECKS - set(lines))}:\n{text[-1500:]}"
    bad = [f"{k} {s} {d}" for k, (s, d) in lines.items() if s != "OK"]
    assert not bad, "doctor not OK: " + "; ".join(bad)
    return {"notes": f"install {secs:.0f}s; doctor {len(lines)}/{len(lines)} OK"}


@scenario("S9.lookup", ["fresh"], budget=60, fresh=True)
def s9_lookup(ctx):
    """Only the fresh MCP server start + sv_status (doctor, atlas; it never starts Excel: server.host() is lazy).
    No Smart View call: the fresh engine's COM Excel inherits the simulated profile (no SSO, so it could only end
    as LIMIT) and would write its global options into the owner's real HKCU."""
    import psutil
    from mcpclient import McpClient

    fr, p = _fresh(ctx), helpers.fresh_paths()
    c = McpClient(env=fr["env"], cmd=[str(p["py"]), "-B", str(p["engine"] / "run_server.py")], cwd=p["engine"])
    try:
        st, _ = c.must("sv_status", {})
        servers = helpers.server_pids(c)
        assert servers, "no child_server.py under the fresh relay"
        cmd = " ".join(psutil.Process(servers[0]).cmdline()).lower()
        assert str(p["engine"]).lower() in cmd, f"the server does not run from the fresh install: {cmd}"
        fr["owners"].update(helpers.fresh_owners(c))
        assert not st.get("excel_running"), f"sv_status reports a running Excel in the fresh server: {st}"
    finally:
        c.close()
        killed = helpers.kill_fresh_excels(fr["owners"])
    assert not killed, f"the fresh server started a hidden Excel {killed}"
    raise helpers.Limit("fresh-profile lookup needs a real signed-in Excel; COM Excel inherits the simulated env — "
                        "verify manually on a new machine")


# ================================================================== S10 features
_CHANGED = re.compile(r"\b(created|updated|added|linked|copied|removed)\b")
_SSO_VALUE = re.compile(r"sso_token=(?!\*\*\*)([^&\s'\"]{6,})")


@scenario("S10.1.bundle", ["feature"], budget=60)
def s10_bundle(ctx):
    import subprocess
    import sys
    import zipfile

    from mcpclient import ENGINE
    from smartview_mcp import config

    home = Path(config.HOME)
    secrets = {config.HTTP_SECRET} if config.HTTP_SECRET else set()
    for f in (home / "logs").glob("*.log*"):          # any raw sso_token value the logs may hold
        try:
            secrets |= set(_SSO_VALUE.findall(f.read_text(encoding="utf-8", errors="replace")))
        except OSError:
            continue
    bundles = home / "bundles"
    before = set(bundles.glob("*.zip"))
    r = subprocess.run([sys.executable, "-m", "smartview_mcp.cli", "doctor", "--bundle"], cwd=ENGINE,
                       capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
    made = sorted(set(bundles.glob("*.zip")) - before)
    try:
        assert r.returncode == 0 and made, f"doctor --bundle exit {r.returncode}, no zip: {(r.stdout + r.stderr)[-800:]}"
        leaks, names = [], []
        with zipfile.ZipFile(made[-1]) as z:
            names = z.namelist()
            for n in names:
                text = z.read(n).decode("utf-8", errors="replace")
                if any(s in text for s in secrets):
                    leaks.append(f"{n}: a secret value")
                if _SSO_VALUE.search(text):
                    leaks.append(f"{n}: unredacted sso_token=")
        assert "doctor.json" in names and "versions.json" in names, f"members {names}"
        assert not leaks, "; ".join(leaks)
    finally:
        for z in made:
            z.unlink(missing_ok=True)
    return {"notes": f"{len(names)} members; {len(secrets)} secret value(s) searched; none found"}


@scenario("S10.2.setup", ["feature", "fresh"], budget=120, fresh=True)
def s10_setup(ctx):
    env, p = _fresh(ctx)["env"], helpers.fresh_paths()
    outs = []
    for _ in range(2):
        r = helpers.fresh_cli(env, "setup", timeout=300)
        assert r.returncode == 0, f"setup exit {r.returncode}: {(r.stdout + r.stderr)[-800:]}"
        outs.append(r.stdout)
    changed = [ln for ln in outs[1].splitlines() if _CHANGED.search(ln)]
    assert not changed, f"the second setup still changed: {changed}"
    for f in (p["copilot"], p["claude"]):
        servers = json.loads(f.read_text(encoding="utf-8"))["mcpServers"]
        assert "smartass" in servers and "smartview" not in servers, f"{f.name}: {sorted(servers)}"
    lnks = {f.name for f in p["startup"].glob("*.lnk")}
    assert {"smartass-engine.lnk", "smartass-sso-watcher.lnk"} <= lnks, f"Startup: {sorted(lnks)}"
    scripts = str(p["engine"] / ".venv" / "Scripts").lower()
    assert scripts in p["pathfile"].read_text(encoding="utf-8").lower(), "venv Scripts not on the (fake) user PATH"
    data = json.loads(p["copilot"].read_text(encoding="utf-8"))
    data["mcpServers"] = {"smartview": data["mcpServers"]["smartass"], "matrix-other": {"command": "x"}}
    p["copilot"].write_text(json.dumps(data), encoding="utf-8")
    r = helpers.fresh_cli(env, "setup", timeout=300)
    line = next((ln for ln in r.stdout.splitlines() if ln.startswith("copilot mcp:")), "")
    servers = json.loads(p["copilot"].read_text(encoding="utf-8"))["mcpServers"]
    assert "updated" in line, f"legacy entry: {line!r}"
    assert set(servers) == {"smartass", "matrix-other"}, f"after the legacy migration: {sorted(servers)}"
    unchanged = [ln.split(":")[0] for ln in outs[1].splitlines() if not _CHANGED.search(ln)]
    return {"notes": f"2nd run unchanged ({len(unchanged)} lines); smartview -> smartass migrated, other kept"}


@scenario("S10.3.reload", ["feature"], budget=90)
def s10_reload(ctx):
    import os

    from mcpclient import McpClient
    from smartview_mcp import config

    f = Path(config.__file__)
    st = f.stat()

    def changed(c) -> int:
        return sum(1 for m in list(c.notes) if m.get("method") == "notifications/tools/list_changed")
    c = McpClient(cwd=ctx.scratch)
    try:
        with helpers.tracked(c):
            c.must("sv_status", {})
            n0 = changed(c)
            os.utime(f)                                  # the relay sees a new fingerprint on the next request
            try:
                b, secs = c.must("sv_status", {})
            finally:                                     # the owner's relays must not see a change
                os.utime(f, ns=(st.st_atime_ns, st.st_mtime_ns))
            t0 = time.monotonic()
            while changed(c) - n0 < 1 and time.monotonic() - t0 < 30:
                time.sleep(0.5)
            time.sleep(3)
            n = changed(c) - n0
            assert b, "sv_status gave nothing after the reload"
            assert n == 1, f"{n} tools/list_changed notifications after touching config.py, expected 1"
    finally:
        if f.stat().st_mtime_ns != st.st_mtime_ns:
            os.utime(f, ns=(st.st_atime_ns, st.st_mtime_ns))
        c.close()
    return {"notes": f"1 list_changed; sv_status after reload {secs:.1f}s"}


@scenario("S10.4.update", ["feature", "fresh"], budget=600, fresh=True)
def s10_update(ctx):
    env, p = _fresh(ctx)["env"], helpers.fresh_paths()

    def version() -> str:
        return helpers.fresh_py(env, "from smartview_mcp import __version__; print(__version__)").stdout.strip()
    notes = []
    assert version() == "0.2.0", f"fresh install is {version()}"
    helpers.make_release("0.2.1")
    guard = helpers.fresh_py(env, "import time\nfrom smartview_mcp import update\n"
                                  "update.subprocess.Popen = lambda *a, **k: None\nnow = time.time()\n"
                                  "print(update.start_if_available(now), update.start_if_available(now + 60),"
                                  " update.start_if_available(now + 31 * 60))")
    (p["home"] / "update_attempt.json").unlink(missing_ok=True)
    assert guard.stdout.split() == ["True", "False", "True"], \
        f"loop guard (first, +1 min, +31 min): {guard.stdout.strip()} {guard.stderr[-500:]}"
    notes.append("guard True/False/True")
    r = helpers.fresh_cli(env, "update")
    assert "installed 0.2.1" in r.stdout, f"update to 0.2.1: {(r.stdout + r.stderr)[-800:]}"
    assert version() == "0.2.1", f"after the update: {version()}"
    r = helpers.fresh_cli(env, "update")
    assert r.stdout.startswith("no update"), f"second update: {r.stdout.strip()}"
    notes.append("0.2.1 installed, 2nd = no update")
    cfg = p["config"].read_text(encoding="utf-8")
    new, n = re.subn(r"(?m)^\[release\][ \t]*$", '[release]\nversion = "0.2.0"', cfg, count=1)
    assert n == 1, "no [release] section in the fresh config.toml"
    p["config"].write_text(new, encoding="utf-8")
    r = helpers.fresh_cli(env, "update")
    assert "installed 0.2.0" in r.stdout, f"pin 0.2.0: {(r.stdout + r.stderr)[-800:]}"
    assert version() == "0.2.0", f"after the pinned update: {version()}"
    r = helpers.fresh_cli(env, "update")
    assert r.stdout.startswith("no update"), f"update while pinned: {r.stdout.strip()}"
    notes.append("pin 0.2.0 -> downgraded and kept")
    return {"notes": "; ".join(notes)}


@scenario("S10.5.textfixes", ["feature"], budget=30)
def s10_textfixes(ctx):
    from smartview_mcp import config
    from smartview_mcp import drill_lines as dl

    f = Path(config.USER_TEXT_FIXES)
    old = f.read_bytes() if f.exists() else None
    key, val = "MATRIX\ufffdS10.5", "MATRIX-S10.5-ok"
    shipped = json.loads(Path(config.TEXT_FIXES).read_text(encoding="utf-8")).get("fixes", {})
    lost = next((k for k in shipped if "\ufffd" in k), None)
    assert lost, f"no shipped fix with a lost letter (U+FFFD) in {Path(config.TEXT_FIXES).name}"
    try:
        data = json.loads(old.decode("utf-8")) if old else {}
        fixes = {**data.get("fixes", {}), key: val, lost: "MATRIX-override"}
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_text(json.dumps({**data, "fixes": fixes}, ensure_ascii=False), encoding="utf-8")
        merged = dl.load_text_fixes(config.TEXT_FIXES, f)
        stats = {"fixed_mojibake": 0, "fixed_by_list": 0, "unfixed": set()}
        got = dl.repair_text(f"a {key} b", merged, stats)
        assert got == f"a {val} b", f"user fix not applied: {got!r}"
        got = dl.repair_text(lost, merged, stats)
        assert got == "MATRIX-override", f"the user's list does not win over the shipped one: {got!r}"
    finally:
        if old is None:
            f.unlink(missing_ok=True)
        else:
            f.write_bytes(old)
    # The 12881 drill lines (95, cached in drill.db 2026-10-23) have no lost letter, so they hit no shipped key
    # (the only one is H\ufffd\ufffdyhastighet) and a drill could not show the override.
    raise helpers.Limit("user fix + override of a shipped key verified in-process (repair_text); the drill part "
                        "is not tested: the 12881 lines hit no shipped text-fix key")


@scenario("S10.6.pack", ["feature"], budget=30)
def s10_pack(ctx):
    import subprocess
    import sys

    from mcpclient import ENGINE

    name = "source_product_list_2025.json"           # in packs\produktregnskap\knowledge only (not in core)
    code = ("import json; from smartview_mcp import config as c; "
            f"print(json.dumps({{'packs': c.PACKS, 'file': str(c.knowledge_file({name!r})), "
            "'pack': str(c.PACKS_DIR / 'produktregnskap')}))")
    r = subprocess.run([sys.executable, "-c", code], cwd=ENGINE, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=60)   # a new process reads config.toml as the child server does
    assert r.returncode == 0, f"config probe exit {r.returncode}: {r.stderr[-800:]}"
    got = json.loads(r.stdout.strip().splitlines()[-1])
    enabled = "produktregnskap" in got["packs"]
    under = got["file"].lower().startswith(got["pack"].lower().rstrip("\\") + "\\")
    if enabled:
        assert under, f"pack enabled but knowledge_file({name}) -> {got['file']}"
    else:
        assert not under, f"pack disabled but knowledge_file({name}) -> {got['file']}"
    raise helpers.Limit(f"pack has no runtime consumer; wiring only verified (pack {'en' if enabled else 'dis'}abled, "
                        f"{name} -> {got['file']})")


@scenario("S10.7.style", ["feature"], budget=60)
def s10_style(ctx):
    out = _small_build(ctx.client, "MATRIX S10.7 style", hang=240, timeout=300)
    fs = out.get("formula_style")
    assert isinstance(fs, dict) and fs.get("ok"), f"formula_style {fs!r}"
    cells = _formats_per_cell(ctx.client, out["workbook"], out["answer_sheet"], "A1:C5")
    styles = sorted({s for _nf, s in cells.values() if s})
    assert any(s.startswith("TN ") for s in styles), f"no TN style in {out['answer_sheet']} A1:C5: {styles}"
    return {"notes": f"formula_style ok; styles {styles}"}


@scenario("S10.8.coldexcel", ["feature"], budget=240)
def s10_coldexcel(ctx):
    import pythoncom

    vis, owned = helpers.visible_excel_pids(), helpers.engine_excel_pids(ctx.scratch)
    if vis - owned:
        raise helpers.Limit(f"the user's own Excel is open (pid {sorted(vis - owned)}): a cold start needs none")
    for pid in vis & owned:
        helpers.quit_owned_excel(pid, ctx.scratch)
    left = helpers.visible_excel_pids()
    assert not left, f"could not close the matrix's Excel {sorted(left)}"
    _, _, err = _call(ctx.client, "sv_restart", {})
    assert not err, f"sv_restart: {err[:300]}"
    new: set[int] = set()
    try:
        t0 = time.monotonic()
        out = _small_build(ctx.client, "MATRIX S10.8 cold excel", hang=240, timeout=300)
        secs = time.monotonic() - t0
        new = helpers.visible_excel_pids() & helpers.engine_excel_pids(ctx.scratch)
        assert new, f"no visible matrix Excel after the cold build ({secs:.0f}s)"
        pythoncom.CoInitialize()
        try:
            app, _ = helpers.app_of(sorted(new)[0])
            visible = bool(app.Visible) if app is not None else None
            del app
        finally:
            pythoncom.CoUninitialize()
        assert visible, f"Excel {sorted(new)} Visible={visible}"
    finally:
        for pid in new or (helpers.visible_excel_pids() & helpers.engine_excel_pids(ctx.scratch)):
            helpers.quit_owned_excel(pid, ctx.scratch)
    return {"notes": f"cold build {secs:.0f}s into {Path(out['workbook']).name}; Visible=True"}


@scenario("S10.9.options", ["feature"], budget=120)
def s10_options(ctx):
    from smartview_mcp import config

    orig = Path(config.HOME) / "smartview_options_original.json"
    before = helpers.sv_options()
    orig_before = orig.read_text() if orig.exists() else None
    try:
        _, secs, err = _call(ctx.client, "sv_restart", {})
        assert not err, f"sv_restart: {err[:300]}"
        follow = _followup(ctx)                          # the new Excel has applied the engine's options
        after = helpers.sv_options()
        assert after == before, f"Smart View options changed by sv_restart: {before} -> {after}"
        if orig_before is not None and orig.exists():
            assert orig.read_text() == orig_before, f"{orig.name} changed: {orig_before} -> {orig.read_text()}"
    finally:
        if helpers.sv_options() != before:
            helpers.write_sv_options(before)
    return {"notes": f"{before} unchanged; restart {secs:.0f}s; {follow}"}


@scenario("S10.10.http", ["feature"], budget=30)
def s10_http(ctx):
    import urllib.error
    import urllib.request

    from smartview_mcp import config

    url = f"http://127.0.0.1:{config.HTTP_PORT}/health"
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

    def get(headers: dict) -> int:
        try:
            with opener.open(urllib.request.Request(url, headers=headers), timeout=30) as r:
                return r.status
        except urllib.error.HTTPError as e:
            return e.code
    try:
        none = get({})
    except (urllib.error.URLError, OSError) as e:
        raise helpers.Limit(f"`smartass serve` is not answering on {url}: {e}") from None
    wrong = get({"X-Smartass-Secret": "matrix-wrong-secret"})
    right = get({"X-Smartass-Secret": config.HTTP_SECRET})
    assert none in (401, 403) and wrong in (401, 403), f"without/wrong secret: {none}/{wrong}"
    assert right == 200, f"with the secret: {right}"
    return {"notes": f"no secret {none}, wrong {wrong}, right {right}"}
