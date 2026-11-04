"""Benchmarks B01-B16 (docs/specs/2026-11-02-benchmarks.md) over stdio MCP, on the one route: the pipeline
(sv_gate -> sv_answer, follow-ups and recipes; B12 uses the sv_drill utility).

Each case is scored on:
- correct: every oracle value (an independent single-cell HTTP fetch in the bench process, run after the MCP server
  exited, so it can neither warm the route nor share its browser) appears on the answer sheet within 1 NOK (or as
  MNOK/kNOK rounded), every check OK, no formula errors; an oracle may be a list of cells, scored as their sum;
- first_visible_s: first new sheet in a visible Excel, polled from the side (B16 fails above 2.5 s);
- total_s: the route's tool calls.
Gate cases (B10, B11) and the trap questions T01-T12 are scored by sv_gate (suggested outcome = expected; the report
adds false-go and needless-ask counts). The pinned cases call sv_gate first and pass deviation="bench pins recipe
<name>" only where the gate names another recipe (or a finance default is switched off). Prints no figures.

Usage (from engine): uv run python tests\\live\\bench.py [--only B01,B02] [--keep] [--out NAME]
Writes docs\\audit\\<date>-<out>.json and .md (--out defaults to bench-core).
"""
from __future__ import annotations

import argparse
import json
import sys
import threading
import time
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

import helpers
import scenarios as S
from mcpclient import ENGINE, McpClient

from smartass.core import config

AUDIT = ENGINE.parent / "docs" / "audit"
PERIOD = "Sep 26"
TOL = 1.0


# ================================================================== scoring (pure; tested in test_matrix_logic.py)
def found(want: float, numbers: list[float]) -> str | None:
    """How `want` appears among the sheet's numbers: 'exact' (±1 NOK), 'kNOK'/'MNOK' (rounded to that unit), or None.
    Sign is ignored: answer sheets may show costs positive."""
    w = abs(want)
    if any(abs(abs(v) - w) <= TOL for v in numbers):
        return "exact"
    for unit, s in (("kNOK", 1e3), ("MNOK", 1e6)):
        if w >= s and any(abs(abs(v) * s - w) <= s / 2 for v in numbers):
            return unit
    return None


def numbers_of(grid: list[list]) -> list[float]:
    return [float(v) for row in grid for v in row if isinstance(v, (int, float)) and not isinstance(v, bool)]


def verdict(r: dict) -> str:
    if r.get("error"):
        return "ERROR"
    if "gate_got" in r:     # a gate case is scored by sv_gate's suggested outcome only (oracle_pass re-scores all)
        return "PASS" if r["gate_got"] == r["gate"] else "FAIL"
    ok = all(r["oracle"].values()) and r["checks_ok"] and not r["formula_errors"] and not r["problems"]
    return "PASS" if ok else "FAIL"


# ================================================================== side poller: time to first visible output
def _visible_sheets() -> set[tuple[str, str]]:
    """(workbook, sheet) of every visible Excel, reached through its windows: the running object table misses new
    unsaved workbooks of an automation-started Excel, which is where answers land."""
    from smartview_mcp.excel.instances import _app_from_window, _excel_windows
    out: set[tuple[str, str]] = set()
    for _pid, _main, ex7s in _excel_windows():
        app = next((a for a in (_app_from_window(h) for h, _ in ex7s[:1]) if a is not None), None)
        try:
            for wb in app.Workbooks if app is not None else []:
                for ws in wb.Worksheets:
                    out.add((str(wb.Name), str(ws.Name)))
        except Exception:   # noqa: BLE001 - busy or gone: try again next poll
            continue
    return out


class FirstVisible(threading.Thread):
    """Seconds from start() to the first new sheet in a visible Excel. Read-only COM from its own apartment."""

    def __init__(self, poll: float = 0.5):
        super().__init__(daemon=True)
        self.poll, self.ready, self.stop = poll, threading.Event(), threading.Event()
        self.t0 = 0.0
        self.secs: float | None = None
        self.sheets: list[tuple[str, str]] = []

    def run(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            before = _visible_sheets()
            self.t0 = time.monotonic()
            self.ready.set()
            while not self.stop.is_set():
                new = _visible_sheets() - before
                if new:
                    self.secs = round(time.monotonic() - self.t0, 2)
                    self.sheets = sorted(new)
                    return
                self.stop.wait(self.poll)
        finally:
            self.ready.set()
            pythoncom.CoUninitialize()

    def __enter__(self):
        self.start()
        self.ready.wait(30)
        return self

    def __exit__(self, *exc):
        self.stop.set()
        self.join(5)


# ================================================================== cases
def cell(conn: str, pov: dict, row: tuple[str, str], col: tuple[str, str]) -> dict:
    """One oracle cell: the keyword arguments of helpers.oracle_value."""
    return {"connection": conn, "pov": pov, "row": row, "col": col}


def fm(entity: str, bal: str = "Accumulated") -> dict:
    return {"All Entities": entity, "All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": bal}


def fp(entity: str, version: str, period: str = "Sep", year: str = "FY26") -> dict:
    return {"Entity": entity, "Years": year, "Period": period, "Version": version, "All Balance Types": "Accumulated"}


def ly(p: str) -> str:
    return f"{p[:3]} {int(p[4:]) - 1:02d}"


@dataclass
class Case:
    id: str
    question: str
    gate: str                                        # clear | ambiguous | infeasible | followup
    run: Callable | None = None                      # fn(ctx) -> Run dict on the pipeline
    oracle: Callable | None = None                   # fn(ctx, run) -> {label: cell(...)}
    agent: bool = False                              # scored by sv_gate (gate quality), not a script
    needs: str | None = None                         # a case whose state this follow-up edits


CASES: dict[str, Case] = {}


def case(id: str, question: str, gate: str = "clear", oracle=None, needs=None, agent=False):
    def deco(fn):
        CASES[id] = Case(id, question, gate, fn, oracle, agent, needs)
        return fn
    return deco


def _run(out: dict, **kw) -> dict:
    return {"out": out, "workbook": out.get("workbook"), "sheet": out.get("answer_sheet"), **kw}


def _critical(out: dict) -> list[str]:
    return [f"{x['severity']}: {x['text']}" for x in out["review"]["findings"] if x["severity"] == "Critical"]


def recipe_v1(case_id: str, name: str, params: dict, checklist: dict | None = None) -> dict:
    """A new question on the core route as a recipe spec (the gate is clear: the bench questions are scoped)."""
    ck = {"period": "the closed month (FinMax Current Reporting Month)", "scenario": "Actual (and plan)",
          "entity": str(params.get("entity") or params.get("area") or params.get("parent") or "un.SY"),
          "measure": "as the question", "currency_sign": "NOK, EPM signs", "granularity": "as the recipe",
          **(checklist or {})}
    return {"version": 1, "question": CASES[case_id].question, "gate": {"outcome": "clear", "checklist": ck},
            "recipe": {"name": name, "params": params}}


def pin_deviation(gate: dict, spec: dict) -> str | None:
    """The deviation a pinned bench spec needs: sv_answer refuses a recipe other than the one sv_gate named (or a
    free-form spec where it named one) and recipe params that switch a finance default off, unless spec.deviation
    says why. None when the spec is the standard analysis."""
    from smartview_mcp.core import recipes
    named, rec = (gate.get("recipe") or {}).get("name"), spec.get("recipe") or {}
    departs = bool(named) and rec.get("name") != named
    departs = departs or bool(rec and recipes.departures(rec.get("name"), rec.get("params") or {}))
    departs = departs or bool(recipes.layout_departures(spec.get("layout")))
    return f"bench pins recipe {rec.get('name') or 'free-form'}" if departs else None


def _answer(ctx, spec: dict) -> tuple[dict, str]:
    """sv_answer on a new workbook with the bench's pinned spec, after sv_gate on its question; the pin is a
    deviation only where the gate (or a finance default) disagrees. Returns (reply, note)."""
    gate, _ = ctx.client.must("sv_gate", {"question": spec["question"]})
    dev = pin_deviation(gate, spec)
    out, _ = ctx.client.must("sv_answer", {"spec": {**spec, "deviation": dev} if dev else spec, "target": "new"})
    named = (gate.get("recipe") or {}).get("name")
    return out, (f"{dev} (gate named {named or 'no recipe'}); " if dev else "")


def _core(case_id: str, name: str, params: dict, keep: str | None = None, check=None, **run_kw):
    def run(ctx):
        out, note = _answer(ctx, recipe_v1(case_id, name, params))
        if keep:
            ctx.state.update({f"{keep}_wb": out["workbook"], f"{keep}_q": int(str(out["question"]).lstrip("Q")),
                              f"{keep}_out": out})
        return _run(out, problems=_critical(out) + (check(out) if check else []),
                    notes=f"{note}period {out.get('period')}", **run_kw)
    return run


def _followup(state_key: str, patch_fn, texts=None, check=None):
    def run(ctx):
        out, _ = ctx.client.must("sv_answer", {"followup": patch_fn(ctx), "q": ctx.state[f"{state_key}_q"],
                                               "target": ctx.state[f"{state_key}_wb"]})
        problems = _critical(out) + (check(out) if check else [])
        return _run(out, problems=problems, texts=texts or [], notes=f"reused {out.get('reused')}")
    return run


def _fin_oracle(entity: str, fp_entity: str | None, accounts=("ac.EBITDA",)):
    def oracle(ctx, run):
        p = run["out"]["period"]
        o = {f"{a} YTD": cell("FinMax", fm(entity), ("All Accounts", a), ("All Time Periods", p)) for a in accounts}
        o["EBITDA YTD LY"] = cell("FinMax", fm(entity), ("All Accounts", "ac.EBITDA"), ("All Time Periods", ly(p)))
        if fp_entity:
            o["EBITDA budget YTD"] = cell("FinPlan", fp(fp_entity, "Final", p[:3], f"FY{p[4:]}"),
                                          ("Account", "EBITDA"), ("Scenario", "Target"))
        return o
    return oracle


case("B01", "Wholesale EBITDA YTD Sep 26 vs last year and budget", oracle=_fin_oracle("un.SY.2WHS", "2WHS Wholesale"))(
    _core("B01", "pnl", {"area": "Wholesale"}))


BUS_O = {"EBITDA YTD Sep 26": cell("FinMax", fm("un.SY.2BUS"), ("All Accounts", "ac.EBITDA"), ("All Time Periods", PERIOD)),
         "EBITDA YTD Sep 25": cell("FinMax", fm("un.SY.2BUS"), ("All Accounts", "ac.EBITDA"), ("All Time Periods", "Sep 25")),
         "EBITDA Jun 26": cell("FinMax", fm("un.SY.2BUS", "Periodic"), ("All Accounts", "ac.EBITDA"),
                               ("All Time Periods", "Jun 26"))}


def b02_v1() -> dict:
    """B02 as an answer spec v1 for sv_answer."""
    pulls = S._bus_pulls()
    pulls[1].pop("title", None)                     # sheet names come from the pipeline
    return {"version": 1, "question": CASES["B02"].question,
            "gate": {"outcome": "clear", "checklist": {
                "period": "Jan-Dec 2026 and 2025 (closed months), YTD Sep", "scenario": "Actual",
                "entity": "un.SY.2BUS Business Norway", "measure": "EBITDA (ac.EBITDA)",
                "currency_sign": "NOK, FinMax signs", "granularity": "month (Periodic) + Accumulated YTD"}},
            "assumptions": ["Actual, NOK; months are Periodic; the reconciliation uses Accumulated YTD Sep"],
            "pulls": pulls, "layout": S._bus_spec()}


@case("B02", "Business Norway EBITDA per month 2026 vs 2025, reconciled to accumulated YTD",
      oracle=lambda ctx, run: BUS_O)
def b02(ctx):
    out, note = _answer(ctx, b02_v1())
    ctx.state.update(b02_wb=out["workbook"], b02_q=int(str(out["question"]).lstrip("Q")), b02_outline=out["outline"])
    return _run(out, problems=_critical(out), notes=f"{note}routes {out['routes']}")


def b03_patch(ctx) -> dict:
    """B03 as a follow-up patch on B02: the FinPlan pull and one line in B02's second section (found by its title)."""
    pull = {"id": "fp", "connection": "FinPlan", "rows": [{"dimension": "Account", "members": ["EBITDA"]}],
            "columns": [{"dimension": "Scenario", "members": ["Actual"]}], "pov": S.FP_POV}
    title = ctx.state["b02_outline"]["sections"][1]["title"]
    line = {"key": "fp", "label": "EBITDA YTD Sep (FinPlan Actual)",
            "cells": [{"pull": "fp", "row": ["EBITDA"], "col": ["Actual"]}, None, None, None]}
    return {"pulls": [pull], "layout": {"sections": [{"title": title, "lines": [line]}]}}


case("B03", "Follow-up to B02: add FinPlan Actual EBITDA YTD", "followup", needs="B02",
     oracle=lambda ctx, run: {**BUS_O, "FinPlan Actual EBITDA YTD": cell("FinPlan", S.FP_POV, ("Account", "EBITDA"),
                                                                         ("Scenario", "Actual"))})(
    _followup("b02", b03_patch))


B04_TITLE = "Business Norway EBITDA 2026 vs 2025 (bench B04)"


case("B04", "Follow-up to B02: add a note and change the title", "followup", needs="B02",
     oracle=lambda ctx, run: BUS_O)(_followup(
         "b02", lambda ctx: {"layout": {"title": B04_TITLE, "notes": {"append": ["Bench note B04."]}}},
         texts=[B04_TITLE, "Bench note B04."],
         check=lambda out: [f"re-fetched {s} ({r})" for s, r in out["routes"].items() if r != "reused"]))


def _period(run: dict) -> str:
    """The recipe's closed month."""
    return run["out"].get("period") or PERIOD


def _b05_oracle(ctx, run):
    p = _period(run)
    return {"EBITDA actual YTD": cell("FinPlan", fp("2TEC Technology", "Working", p[:3], f"FY{p[4:]}"),
                                      ("Account", "EBITDA"), ("Scenario", "Actual")),
            "EBITDA target YTD": cell("FinPlan", fp("2TEC Technology", "Final", p[:3], f"FY{p[4:]}"),
                                      ("Account", "EBITDA"), ("Scenario", "Target"))}


def _b07_oracle(ctx, run):
    p = _period(run)
    return {f"C30 {p}": cell("FinMax", fm("un.SY", "Periodic"), ("All Accounts", "ac.C30"), ("All Time Periods", p)),
            f"C30 Jan {p[4:]}": cell("FinMax", fm("un.SY", "Periodic"), ("All Accounts", "ac.C30"),
                                     ("All Time Periods", f"Jan {p[4:]}")),
            f"C30 YTD {p}": cell("FinMax", fm("un.SY"), ("All Accounts", "ac.C30"), ("All Time Periods", p))}


case("B05", "Technology actual vs plan, Sep 26", oracle=_b05_oracle)(
    _core("B05", "actual_vs_plan", {"entity": "2TEC Technology"}))
case("B06", "Top 15 cost centres by opex under Technology, Sep 26",
     oracle=lambda ctx, run: {"C70 YTD un.SY.2TEC": cell("FinMax", fm("un.SY.2TEC"), ("All Accounts", "ac.C70"),
                                                         ("All Time Periods", _period(run)))})(
    _core("B06", "cost_base", {"parent": "un.SY.2TEC", "top": 15}))
case("B07", "Telenor Norge monthly revenue trend through Sep 26", oracle=_b07_oracle)(
    _core("B07", "monthly_trend", {"entity": "un.SY", "accounts": ["ac.C30"]}, keep="b07"))

BA_KEYS = ("wholesale", "business", "ssm", "ei", "technology", "it", "adm", "nordics")


def b08_followup(period: str) -> dict:
    """Follow-up to B07 (C30 per month, un.SY): the same months per business area, the remainder as
    'Other / eliminations', and a check that the split adds up to the trend line a0."""
    from smartview_mcp import catalog, templates
    months = templates.months_to(period)
    bas = [e for e in catalog.load().entities if e.key in BA_KEYS and not e.abolished]
    n = len(months)

    def per_month(m: str) -> list:
        return [{"pull": "ba", "row": [m], "col": [mo]} for mo in months] + [{"rowsum": list(range(n))}]
    return {
        "pulls": [{"id": "ba", "connection": "FinMax",
                   "rows": [{"dimension": "All Entities", "members": [e.finmax for e in bas] + ["un.SY"]}],
                   "columns": [{"dimension": "All Time Periods", "members": months}],
                   "pov": {"All Accounts": "ac.C30", "All Scenarios": "Actual", "All Currencies": "NOK",
                           "All Balance Types": "Periodic"}}],
        "layout": {"sections": [{
            "title": "By business area",
            "repeat": {"over": [{"m": e.finmax, "label": e.label} for e in bas],
                       "line": {"key": "ba{i}", "label": "{label}", "kind": "line", "cells": per_month("{m}")}},
            "lines": [{"key": "split_sum", "label": "Sum of business areas", "aux": True,
                       "cells": [{"sum": "ba[0-9]*"} for _ in range(n + 1)]},
                      {"key": "split_other", "label": "Other / eliminations", "kind": "line",
                       "cells": [{"expr": "split_tot-split_sum"} for _ in range(n + 1)]},
                      {"key": "split_tot", "label": "Telenor Norge (un.SY)", "kind": "sum",
                       "cells": per_month("un.SY")}]}],
            "checks": [{"label": "Business areas + other = Telenor Norge (trend)", "expr": "split_tot-a0",
                        "columns": list(range(n + 1))}]}}


def _b08_oracle(ctx, run):
    p = ctx.state["b07_out"]["period"]
    return {f"C30 {p} {e}": cell("FinMax", fm(e, "Periodic"), ("All Accounts", "ac.C30"), ("All Time Periods", p))
            for e in ("un.SY.2BUS", "un.SY.2TEC")}


case("B08", "Follow-up to B07: split by business area", "followup", oracle=_b08_oracle, needs="B07")(
    _followup("b07", lambda ctx: b08_followup(ctx.state["b07_out"]["period"])))


@case("B09", "What do we earn on mobile subscriptions?")
def b09(ctx):
    from smartview_mcp import concepts
    ids = [c["id"] for c in concepts.find("what do we earn on mobile subscriptions")]
    if not ids:
        raise AssertionError("the concept library found no concept")
    run = _core("B09", "concept", {"concept": ids[0]})(ctx)
    return {**run, "notes": f"concept {ids[0]}; {run['notes']}"}


CASES["B10"] = Case("B10", "EBITDA per product for Mobile Consumer", "infeasible", agent=True)
CASES["B11"] = Case("B11", "How are we doing on costs?", "ambiguous", agent=True)

# Trap questions (docs/specs/2026-11-02-benchmarks.md): gate quality only. Expected outcomes are owner decisions.
for _id, _q, _g in [
        ("T01", "What's our budget for Business?", "ambiguous"),
        ("T02", "Opex for Wholesale this year", "ambiguous"),
        ("T03", "EBITDA margin per customer", "infeasible"),
        ("T04", "Revenue in EUR for Technology", "infeasible"),
        ("T05", "Profit per product", "infeasible"),
        ("T06", "Business EBITDA Sep 26 vs Sep 25", "clear"),
        ("T07", "Q3 numbers for Consumer", "ambiguous"),
        ("T08", "Cost per FTE in IT", "clear"),
        ("T09", "How much did we spend on consultants?", "clear"),
        ("T10", "Forecast for December 27 actuals", "infeasible"),
        ("T11", "Hvilke konsulentselskaper fakturerte mest til Telenor Norge i 2026? og på hvilke initiativer / "
                "prosjekt", "clear"),
        ("T12", "Which suppliers invoiced most on travel in 2026?", "clear")]:
    CASES[_id] = Case(_id, _q, _g, agent=True)


@case("B12", "Lines behind account 12881, un.SY, Sep 26",
      oracle=lambda ctx, run: {"12881 un.SY Sep 26": cell("FinLite", fm("un.SY", "Periodic"),
                                                          ("All Accounts", "ac.12881"), ("All Time Periods", PERIOD))})
def b12(ctx):
    out, _ = ctx.client.must("sv_drill", {"question": "BENCH B12", "cell": S.CELL_12881, "target": "new"})
    rec = out["reconciliation"]
    problems = [] if rec["status"] == "exact" and abs(rec["gap"] or 0) < 0.005 else [f"reconciliation {rec['status']}"]
    return _run(out, numbers=[rec["lines_sum"]], problems=problems, checks_required=False)


def _last_year_only(out: dict) -> list[str]:
    return [] if any("compared with last year only" in w for w in out.get("warnings") or []) else \
        ["no last-year-only warning"]


case("B13", "EBITDA YTD for unit un.SY.5OK_W", oracle=_fin_oracle("un.SY.5OK_W", None))(
    _core("B13", "pnl", {"area": "un.SY.5OK_W", "detail": "full"}, check=_last_year_only))


def _b14_oracle(ctx, run):
    from smartview_mcp.templates import latest_forecast
    o = {"EBITDA target FY": cell("FinPlan", fp("2BUS Business", "Final", "Dec"), ("Account", "EBITDA"),
                                  ("Scenario", "Target"))}
    f = latest_forecast(26)
    if f:
        o[f"EBITDA {f} FY"] = cell("FinPlan", fp("2BUS Business", "Final", "Dec"), ("Account", "EBITDA"), ("Scenario", f))
    return o


case("B14", "Business full-year EBITDA: latest forecast vs budget", oracle=_b14_oracle)(
    _core("B14", "actual_vs_plan", {"entity": "2BUS Business", "month": "Dec", "year": "FY26"}))
case("B15", "Wholesale full P&L (revenue, COGS, opex, EBITDA) actual vs budget vs last year, YTD",
     oracle=_fin_oracle("un.SY.2WHS", "2WHS Wholesale", ("ac.C30", "ac.C40S", "ac.C70", "ac.EBITDA")))(
    _core("B15", "pnl", {"area": "Wholesale", "detail": "full"}))

FRAME_LIMIT_S = 2.5


def b16_problems(out: dict, first_visible_s: float | None) -> list[str]:
    """What B16 (the consultant ranking) must show: the Answer and By project sheets, every check OK, the cube total
    in key_values (the read-back the agent quotes), and a first frame in a visible Excel within 2.5 s."""
    sheets = out.get("sheets") or []
    problems = [f"no Q<n> {s} sheet" for s in ("Answer", "By project") if not any(x.endswith(f" {s}") for x in sheets)]
    checks = out.get("checks") or {}
    problems += [f"check {k}: {v}" for k, v in checks.items() if v != "OK"] or ([] if checks else ["no checks"])
    if "Total (cube)" not in (out.get("key_values") or {}):
        problems.append("'Total (cube)' missing from key_values")
    t = out.get("timings") or {}
    frame = t.get("first frame")
    after_open = None if frame is None else round(frame - (t.get("open target") or 0.0), 2)
    if after_open is None or after_open > FRAME_LIMIT_S:      # the owner's target: <= 2 s after the workbook is open
        problems.append(f"first frame {_s(after_open, 'never')} s after the workbook opened > {FRAME_LIMIT_S} s"
                        f" (seen in Excel at {_s(first_visible_s, 'never')} s)")
    return problems


def _b16_oracle(ctx, run):
    """The cube total: FinLite un.SY YTD C60080 + C60090 at the answer's month (a list of cells = their sum)."""
    month = str(run["out"]["period"]).removeprefix("YTD ").strip()
    return {f"consultants YTD {month}": [cell("FinLite", fm("un.SY"), ("All Accounts", a), ("All Time Periods", month))
                                         for a in ("ac.C60080", "ac.C60090")]}


@case("B16", "Hvilke konsulentselskaper fakturerte mest til Telenor Norge i 2026? og på hvilke initiativer / prosjekt",
      oracle=_b16_oracle)
def b16(ctx):
    q = CASES["B16"].question
    gate, _ = ctx.client.must("sv_gate", {"question": q})
    if gate["suggested_outcome"] != "clear" or not gate.get("recipe"):
        raise AssertionError(f"gate {gate['suggested_outcome']}, recipe {gate.get('recipe')}")
    spec = {"version": 1, "question": q, "recipe": gate["recipe"],
            "gate": {"outcome": "clear", "checklist": {k: str(v.get("value") or v.get("gap") or "")
                                                       for k, v in gate["checklist"].items()}}}
    out, _ = ctx.client.must("sv_answer", {"spec": spec, "target": "new"})
    return _run(out, problems=_critical(out), frame_check=lambda secs: b16_problems(out, secs),
                notes=f"recipe {gate['recipe']['name']}; period {out.get('period')}; timings {out.get('timings')}")


# ================================================================== runner
def sheet_values(client, wb: str, sheet: str) -> list[list]:
    ov, _ = client.must("xl_overview", {"workbook": wb, "preview_rows": 0})
    used = next(s["used"] for s in ov["sheets"] if s["name"] == sheet)
    body, _ = client.must("xl_read", {"workbook": wb, "sheet": sheet, "range": used.split("!")[-1].replace("$", ""),
                                      "what": "values", "max_cells": 50000})
    return body["values"]


def run_case(c: Case, ctx) -> dict:
    """Run one case on the pipeline. The oracle cells are only recorded here (`_oracle_q`, with the sheet's numbers
    in `_nums`); oracle_pass fetches and scores them after the server exited. The provisional result ignores them."""
    r = {"id": c.id, "question": c.question, "gate": c.gate, "route": "core", "first_visible_s": None,
         "total_s": None, "oracle": {}, "match": {}, "checks_ok": False, "formula_errors": [], "problems": [],
         "error": None, "notes": ""}
    if c.agent:
        try:
            t0 = time.monotonic()
            out, _ = ctx.client.must("sv_gate", {"question": c.question})
            r["total_s"] = r["first_visible_s"] = round(time.monotonic() - t0, 2)
            r["gate_got"] = out["suggested_outcome"]
            r["notes"] = f"gate {r['gate_got']}" + (f": {out['reason']}" if out.get("reason") else "") + (
                f" / asks {len(out['questions'])}" if out.get("questions") else "")
            r["result"] = "PASS" if r["gate_got"] == c.gate else "FAIL"
        except Exception as e:  # noqa: BLE001
            r["error"], r["result"] = f"{type(e).__name__}: {e}"[:400], "ERROR"
        return r
    if c.run is None:
        return {**r, "result": "UNSUPPORTED", "notes": "no route for this case"}
    if c.needs and c.needs not in ctx.state.setdefault("passed", set()):
        return {**r, "result": "SKIP", "notes": f"needs {c.needs} to pass in the same run"}
    try:
        t0 = time.monotonic()
        with FirstVisible() as fv:
            run = c.run(ctx)
        r["total_s"] = round(time.monotonic() - t0, 2)
        r["first_visible_s"] = fv.secs
        out = run["out"]
        checks = out.get("checks") or {}
        review = out.get("review")
        # the review grades each check (a 'Warning' check may differ by design); sv_drill has no review
        ok = (not any(x["severity"] == "Critical" and x["text"].startswith("check '") for x in review["findings"])
              if review is not None else S._all_ok(checks))
        r["checks_ok"] = ok and (bool(checks) or not run.get("checks_required", True))
        r["formula_errors"] = S._no_formula_errors(out)
        r["problems"] = list(run.get("problems") or [])
        if run.get("frame_check"):
            r["problems"] += run["frame_check"](fv.secs)
        r["notes"] = run.get("notes") or f"{run.get('sheet')}"
        if run.get("numbers") is not None:
            nums, texts = list(run["numbers"]), []
        else:
            grid = sheet_values(ctx.client, run["workbook"], run["sheet"])
            nums, texts = numbers_of(grid), [str(v) for row in grid for v in row if isinstance(v, str)]
        for t in run.get("texts") or []:
            if not any(t in s for s in texts):
                r["problems"].append(f"text missing: {t!r}")
        r["_oracle_q"] = c.oracle(ctx, run) if c.oracle else {}
        r["_nums"] = nums
    except Exception as e:  # noqa: BLE001 - one case must not stop the run
        r["error"] = f"{type(e).__name__}: {e}"[:400]
        traceback.print_exc()
    r["result"] = verdict(r)
    if r["result"] == "PASS":
        ctx.state.setdefault("passed", set()).add(c.id)
    return r


def oracle_pass(results: list[dict]) -> None:
    """Score the oracle cells after the MCP server has exited (its hidden Edge freed the drill profile): one
    in-process HTTP fetch per cell, so the oracle never warms the route it checks. Re-scores every result."""
    from smartview_mcp import drill_browser
    try:
        for r in results:
            nums = r.pop("_nums", None) or []
            for label, q in (r.pop("_oracle_q", None) or {}).items():
                try:
                    vals = [helpers.oracle_value(**c) for c in (q if isinstance(q, list) else [q])]
                    v = None if any(x is None for x in vals) else sum(vals)
                except Exception as e:  # noqa: BLE001 - one bad oracle cell must not stop the scoring
                    r["oracle"][label], r["match"][label] = False, f"oracle failed: {str(e)[:120]}"
                    continue
                how = found(v, nums) if v is not None else None
                r["oracle"][label] = how is not None
                r["match"][label] = how or ("no oracle value" if v is None else "missing")
            if r["result"] in ("PASS", "FAIL"):
                r["result"] = verdict(r)
    finally:
        drill_browser.shutdown_shared()


def _s(v, none="-"):
    return none if v is None else v


def gate_score(results: list[dict]) -> dict:
    """Advisor quality over the gate cases scored by sv_gate. false-go = should have asked or refused, but went ahead
    (clear); needless-ask = was clear, but asked."""
    scored = [r for r in results if r.get("gate_got")]
    return {"right": sum(r["gate_got"] == r["gate"] for r in scored), "total": len(scored),
            "false_go": sum(r["gate"] in ("ambiguous", "infeasible") and r["gate_got"] == "clear" for r in scored),
            "needless_ask": sum(r["gate"] == "clear" and r["gate_got"] == "ambiguous" for r in scored)}


def write_report(name: str, started: datetime, results: list[dict]) -> Path:
    AUDIT.mkdir(parents=True, exist_ok=True)
    base = AUDIT / f"{started:%Y-%m-%d}-{name}"
    if base.with_suffix(".json").exists():
        base = AUDIT / f"{started:%Y-%m-%d}-{name}-{started:%H%M}"
    base.with_suffix(".json").write_text(json.dumps({"started": started.isoformat(timespec="seconds"), "route": "core",
                                                      "results": results}, indent=1), encoding="utf-8")
    lines = [f"# Benchmarks — route `core` ({started:%Y-%m-%d %H:%M})", "",
             "| ID | Result | First visible s | Total s | Oracle | Notes |", "|---|---|---|---|---|---|"]
    for r in results:
        orc = f"{sum(r['oracle'].values())}/{len(r['oracle'])}" if r["oracle"] else "—"
        why = r["error"] or "; ".join(r["problems"] + [f"{k}: {v}" for k, v in r["match"].items() if v not in
                                                        ("exact", "kNOK", "MNOK")]) or r["notes"]
        lines.append(f"| {r['id']} | {r['result']} | {_s(r['first_visible_s'], '—')} | {_s(r['total_s'], '—')} | {orc} | "
                     f"{why.replace('|', '/')} |")
    g = gate_score(results)
    if g["total"]:
        lines += ["", f"Gate: {g['right']}/{g['total']} right; false-go {g['false_go']} (should have asked or refused, "
                      f"went ahead); needless-ask {g['needless_ask']} (clear, but asked)."]
    base.with_suffix(".md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return base


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Benchmarks B01-B16, trap questions T01-T12")
    ap.add_argument("--route", default="core", choices=["core"], help="the one route (kept for the documented commands)")
    ap.add_argument("--only", default="", help="comma-separated case ids")
    ap.add_argument("--keep", action="store_true", help="leave the scratch workbooks and _bench folder")
    ap.add_argument("--out", default="bench-core", help="report name")
    a = ap.parse_args(argv)
    only = {s.strip().upper() for s in a.only.split(",") if s.strip()}
    chosen = [c for c in CASES.values() if not only or c.id in only]
    scratch = Path(config.ANALYSER_DIR) / "_matrix"     # helpers.cleanup closes workbooks in a _matrix folder only
    scratch.mkdir(parents=True, exist_ok=True)
    ctx = S.Ctx(client=None, scratch=scratch, keep=a.keep)
    ctx.state["excels_before"] = helpers.excel_pids()
    started, results, client = datetime.now(), [], None
    try:
        client = McpClient(cwd=scratch)
        ctx.client = client
        helpers._CLIENTS.append(client)
        for c in sorted(chosen, key=lambda c: c.id):
            r = run_case(c, ctx)
            results.append(r)
            print(f"ran {c.id} ({r['result']} before oracle)", flush=True)
    finally:
        try:
            helpers.cleanup(ctx)
        except Exception as e:  # noqa: BLE001
            print(f"cleanup failed: {e}", file=sys.stderr)
        if client is not None:
            helpers._CLIENTS.remove(client)
            client.close()
    try:
        oracle_pass(results)
    finally:
        for r in results:
            print(f"{r['result']:11} {r['id']}  first {_s(r['first_visible_s'])}s  total {_s(r['total_s'])}s"
                  + (f"  {r['error']}" if r["error"] else ""), flush=True)
        print(f"report: {write_report(a.out, started, results).with_suffix('.md')}")
    return 1 if any(r["result"] in ("FAIL", "ERROR") for r in results) else 0


if __name__ == "__main__":
    sys.exit(main())
