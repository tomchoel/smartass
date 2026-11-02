"""The workbook for fin_pnl: the build request (live Smart View source sheets + an answer spec whose every
number is a formula over them) and the expected value of every source cell the answer uses, from the calc engine's
facts - so the build can prove live read = sheet (reconcile_pull, run after the pull).

Answer columns (one set for every section; the section title says what the comparison is):
  0 month | 1 comparison month | 2 variance | 3 variance % | 4 YTD | 5 comparison YTD | 6 variance | 7 variance % |
  8 material (YTD). Bridges and the landing use column 4.

The sheet follows the engine's rules exactly (specs.evaluate gives the engine's facts):
- variance = current - comparison, positive = better; for a better = 'lower' KPI (cost per FTE) the variance
  column is comparison - current (the var columns are swapped) and has no % (its base would differ);
- a KPI with sign = -1 is multiplied by the labelled input cell 'Sign' (-1), so cost per FTE is shown positive;
- no constants in formulas (tn-formelstil): months in a year / in the period and the sign are labelled inputs;
- averages of a stock (FTE) are AVERAGEA over the opening level + each month end (n+1 points);
- the variance % is left blank where |comparison| < the catalogue's pct_floor_nok (decided from the facts,
  like the Material column);
- every comparison shows the labels of its own cube (FinMax: alias + short name; FinPlan: its own names).
Every section (last year, budget, forecast, bridges, landing) is always shown; `compare` only limits the flags."""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

from .analytics import periods
from .analytics.bridge import components
from .analytics.reconcile import TOL
from .catalog import Catalog
from .fin_pnl import (NAMES, PROVISIONAL, Result, finplan_warnings, flag_text, freshness_text, grand,
                      landing_warnings, open_period_note, ordered_kpis, shown_pct)
from .templates import _fmt

COLUMNS = ["{p}", "Comparison", "Variance", "Variance %", "YTD {p}", "Comparison YTD", "Variance YTD",
           "Variance % YTD", "Material (YTD)"]
KINDS = ["heltall", "heltall", "heltall", "prosent1", "heltall", "heltall", "heltall", "prosent1", "heltall"]
VALUE_COLS = (0, 1, 4, 5)
INPUT_KEYS = ("months_year", "months_ytd", "kpi_sign")      # the only lines with typed values (labelled inputs)
FTE_TOL = 0.05                                             # FTE: live read = sheet within 0.05 FTE

Expected = tuple   # (row, col, value, fact id, tolerance)


@dataclass
class Request:
    question: str
    pulls: list[dict]                                        # pull fields: id, connection, rows, columns, pov
    spec: dict
    expected: dict[str, list[Expected]]                      # pull id -> (row, col, value, fact id, tolerance)
    source: str = "Live read"                                # what the sheets are checked against (check labels)


def _ax(dim: str, members: list[str]) -> dict:
    return {"dimension": dim, "members": list(dict.fromkeys(members))}


def _pull(pull: str, row: list[str], col: list[str]) -> dict:
    return {"pull": pull, "row": row, "col": col}


def _var(better: str, a: int, b: int) -> dict:
    """The variance cell of columns a (current) and b (comparison): positive = better."""
    return {"var": [b, a]} if better == "lower" else {"var": [a, b]}


def _prefixed(expr: str, prefix: str) -> str:
    return re.sub(r"[a-z_]\w*", lambda m: prefix + m.group(0), expr)


def _at4(cell, more=()) -> list:
    """A line's cells with `cell` in column 4 (bridges, landing), then `more`."""
    return [None] * 4 + [cell, *more]


@dataclass
class _Builder:
    res: Result
    cat: Catalog
    pulls: list[dict] = field(default_factory=list)
    expected: dict[str, list] = field(default_factory=dict)
    sections: list[dict] = field(default_factory=list)
    checks: list[dict] = field(default_factory=list)
    method_notes: list[str] = field(default_factory=list)
    chart: dict | None = None
    source: str = "Live read"

    def __post_init__(self):
        s = self.res.scope
        self.p, self.ly = s.period, self.res.ly
        self.mon, self.fy = periods.finplan(self.p)
        self.fy_ly = periods.finplan(self.ly)[1]
        self.fm_lines, self.fp_lines = self.cat.pnl["FinMax"], self.cat.pnl["FinPlan"]
        self.g, self.gp = grand(self.fm_lines), grand(self.fp_lines)
        self.m = self.cat.materiality

    # -------------------------------------------------------------- helpers
    def value(self, fid: str):
        f = self.res.facts.get(fid)
        return None if f is None else f.value

    def expect(self, pull: str, row: list[str], col: list[str], fid: str, tol: float = TOL) -> None:
        self.expected.setdefault(pull, []).append((row, col, self.value(fid), fid, tol))

    def pct(self, v, better: str, a: int, b: int):
        """The variance % cell, blank for a lower-is-better line or a comparison under pct_floor_nok."""
        if better == "lower" or v is None:
            return None
        c = v.comparison.value
        if v.abs.unit == "nok" and c is not None and abs(c) < self.m.pct_floor_nok:
            return None
        return {"varpct": [a, b]}

    def flag(self, comparison: str, key: str):
        t = flag_text(self.res, comparison, key)
        return None if t is None else {"text": t}

    def label(self, cube: str, key: str) -> str:
        return self.res.label(cube, key)

    # -------------------------------------------------------------- A. FinMax vs last year (+ FTE levels)
    def finmax(self) -> None:
        res, p, ly = self.res, self.p, self.ly
        e = res.scope.entity.finmax
        stocks = [x for x in self.fm_lines if x.measure == "stock"]
        base = {"All Entities": e, "All Scenarios": "Actual", "All Currencies": "NOK"}
        self.pulls.append({"id": "pnl", "connection": "FinMax", "pov": base,
                           "rows": [_ax("All Accounts", [x.account for x in self.fm_lines])],
                           "columns": [_ax("All Balance Types", ["Periodic", "Accumulated"]),
                                       _ax("All Time Periods", [p, ly])]})
        self.expected["pnl"] = []
        lines = []
        for x in self.fm_lines:
            a, vs = [x.account], res.fm[x.key]
            if x.measure == "stock":
                cells = [_pull("pnl", a, ["Accumulated", p]), _pull("pnl", a, ["Accumulated", ly]),
                         _var(x.better, 0, 1), self.pct(vs["m"], x.better, 0, 1),
                         {"ref": {"key": f"{x.key}_avg", "col": 4}}, {"ref": {"key": f"{x.key}_avg", "col": 5}},
                         _var(x.better, 4, 5), self.pct(vs["ytd"], x.better, 4, 5), None]
                self.expect("pnl", a, ["Accumulated", p], f"fm.{x.key}.m.cur", FTE_TOL)
                self.expect("pnl", a, ["Accumulated", ly], f"fm.{x.key}.m.ly", FTE_TOL)
                label = f"{self.label('FinMax', x.key)} (month end; YTD = average)"
            else:
                cells = [_pull("pnl", a, ["Periodic", p]), _pull("pnl", a, ["Periodic", ly]), _var(x.better, 0, 1),
                         self.pct(vs["m"], x.better, 0, 1), _pull("pnl", a, ["Accumulated", p]),
                         _pull("pnl", a, ["Accumulated", ly]), _var(x.better, 4, 5),
                         self.pct(vs["ytd"], x.better, 4, 5), self.flag("ly", x.key)]
                for bal, per, part in (("Periodic", p, "m.cur"), ("Periodic", ly, "m.ly"),
                                       ("Accumulated", p, "ytd.cur"), ("Accumulated", ly, "ytd.ly")):
                    self.expect("pnl", a, [bal, per], f"fm.{x.key}.{part}")
                label = self.label("FinMax", x.key)
            ln = {"key": x.key, "label": label, "kind": x.kind, "cells": cells}
            if x.num:
                ln["num"] = x.num
            lines.append(ln)
        self.sections.append({"title": f"P&L vs last year (FinMax): {p} vs {ly}, YTD vs YTD {ly}", "lines": lines})
        self.checks += [{"label": c.label, "expr": c.expr, "columns": list(VALUE_COLS)}
                        for c in self.cat.checks["FinMax"]]
        if stocks:
            self.fte_levels(base, stocks)

    def fte_levels(self, base: dict, stocks) -> None:
        res = self.res
        cur_pts = [periods.opening(self.p), *res.ytd]
        ly_pts = [periods.opening(self.ly), *res.ly_ytd]
        self.pulls.append({"id": "fte", "connection": "FinMax", "pov": {**base, "All Balance Types": "Accumulated"},
                           "rows": [_ax("All Accounts", [x.account for x in stocks])],
                           "columns": [_ax("All Time Periods", cur_pts + ly_pts)]})
        self.expected["fte"] = []
        for x in stocks:
            name = self.label("FinMax", x.key)
            lv = []
            for i, (c, l) in enumerate(zip(cur_pts, ly_pts)):
                lv.append({"key": f"{x.key}_l{i}", "label": f"{name} {c} / {l}", "kind": "line", "aux": True,
                           "num": x.num or "tall1",
                           "cells": [None] * 4 + [_pull("fte", [x.account], [c]), _pull("fte", [x.account], [l])]})
                self.expect("fte", [x.account], [c], f"fm.{x.key}.lvl.{c}", FTE_TOL)
                self.expect("fte", [x.account], [l], f"fm.{x.key}.lvl.{l}", FTE_TOL)
            keys = [ln["key"] for ln in lv]
            lv.append({"key": f"{x.key}_avg", "label": f"Average {name} ({len(keys)} points: opening + month ends)",
                       "kind": "sum", "aux": True, "num": x.num or "tall1",
                       "cells": [None] * 4 + [{"avg": keys}, {"avg": keys}, _var(x.better, 4, 5),
                                              self.pct(res.fm[x.key]["ytd"], x.better, 4, 5)]})
            self.sections.append({"title": f"{name} levels for the average (opening December + each month end)",
                                  "lines": lv})

    # -------------------------------------------------------------- KPIs and the labelled inputs
    def kpis(self) -> None:
        res = self.res
        kl, need_months, need_sign = [], False, False
        for k in ordered_kpis(self.cat, "FinMax"):
            if k.key not in res.kpis:
                continue
            f = res.kpis[k.key]
            den = f"{k.denominator}_avg" if k.basis == "average_stock" else k.denominator
            expr = f"{k.numerator}/{den}"
            if k.annualise:
                expr, need_months = expr + "*months_year/months_ytd", True
            if k.sign == -1:
                expr, need_sign = expr + "*kpi_sign", True

            def cell(part, expr=expr, f=f):
                fact = f.get(part)
                if fact is None:
                    return None
                return {"expr": expr} if fact.value is not None else {"text": "n/a"}
            cells = [cell("m.cur"), cell("m.ly"), None, None, cell("ytd.cur"), cell("ytd.ly"), None, None, None]
            for a, b, j in ((0, 1, 2), (4, 5, 6)):
                if cells[a] and cells[b] and "expr" in cells[a] and "expr" in cells[b]:
                    cells[j] = _var(k.better, a, b)
            label = k.label + (" (annualised)" if k.annualise else "") + (" (headline)" if k.headline else "")
            kl.append({"key": f"k_{k.key}", "label": label, "kind": "line",
                       "num": "prosent1" if k.unit == "pct" else "heltall", "cells": cells})
        ld = res.landing
        if ld is not None and not ld.seasonal:
            need_months = True
        inputs = []
        if need_months:
            inputs += [{"key": "months_year", "label": "Months in a year", "kind": "line", "aux": True,
                        "cells": [{"value": 12}, {"value": 12}, None, None, {"value": 12}, {"value": 12}]},
                       {"key": "months_ytd", "label": f"Months in the period (month; YTD {self.p} / YTD {self.ly})",
                        "kind": "line", "aux": True,
                        "cells": [{"value": 1}, {"value": 1}, None, None, {"value": len(res.ytd)},
                                  {"value": len(res.ly_ytd)}]}]
        if need_sign:
            inputs.append({"key": "kpi_sign", "label": "Sign: cost per FTE shown positive (FinMax books costs "
                                                       "negative)", "kind": "line", "aux": True,
                           "cells": [{"value": -1}, {"value": -1}, None, None, {"value": -1}, {"value": -1}]})
        if inputs:
            self.sections.append({"title": "Inputs", "lines": inputs})
        if kl:
            self.sections.append({"title": "KPIs (FinMax; variance in points or NOK, positive = better)",
                                  "lines": kl})

    # -------------------------------------------------------------- B/C. FinPlan vs budget and the forecast
    def finplan(self) -> None:
        res, s, mon, fy = self.res, self.res.scope, self.mon, self.fy
        ef = s.entity.finplan
        if not (ef and res.budget):
            return
        months = [mon] if mon == "Dec" else [mon, "Dec"]
        accts = [x.account for x in self.fp_lines]
        self.pulls.append({"id": "fpa", "connection": "FinPlan",
                           "pov": {"Entity": ef, "Scenario": "Actual", "Version": "Working"},
                           "rows": [_ax("Account", accts)],
                           "columns": [_ax("Years", [fy, self.fy_ly]),
                                       _ax("All Balance Types", ["Periodic", "Accumulated"]), _ax("Period", months)]})
        scen = ["Target"] + ([s.forecast] if s.forecast and res.forecast else [])
        self.pulls.append({"id": "fpp", "connection": "FinPlan",
                           "pov": {"Entity": ef, "Years": fy, "Version": "Final"}, "rows": [_ax("Account", accts)],
                           "columns": [_ax("Scenario", scen), _ax("All Balance Types", ["Periodic", "Accumulated"]),
                                       _ax("Period", months)]})
        self.expected["fpa"], self.expected["fpp"] = [], []
        bl, fl = [], []
        for x in self.fp_lines:
            a, b, bv = [x.account], f"fp.{x.key}", res.budget[x.key]
            bl.append({"key": f"b_{x.key}", "label": self.label("FinPlan", x.key), "kind": x.kind, "cells": [
                _pull("fpa", a, [fy, "Periodic", mon]), _pull("fpp", a, ["Target", "Periodic", mon]),
                _var(x.better, 0, 1), self.pct(bv["m"], x.better, 0, 1), _pull("fpa", a, [fy, "Accumulated", mon]),
                _pull("fpp", a, ["Target", "Accumulated", mon]), _var(x.better, 4, 5),
                self.pct(bv["ytd"], x.better, 4, 5), self.flag("budget", x.key)]})
            self.expect("fpa", a, [fy, "Periodic", mon], f"{b}.actual.m")
            self.expect("fpa", a, [fy, "Accumulated", mon], f"{b}.actual.ytd")
            self.expect("fpp", a, ["Target", "Periodic", mon], f"{b}.budget.m")
            self.expect("fpp", a, ["Target", "Accumulated", mon], f"{b}.budget.ytd")
            if len(scen) > 1:
                fc, fv = s.forecast, res.forecast[x.key]
                fl.append({"key": f"f_{x.key}", "label": self.label("FinPlan", x.key), "kind": x.kind, "cells": [
                    {"ref": {"key": f"b_{x.key}", "col": 0}}, _pull("fpp", a, [fc, "Periodic", mon]),
                    _var(x.better, 0, 1), self.pct(fv["m"], x.better, 0, 1), {"ref": {"key": f"b_{x.key}", "col": 4}},
                    _pull("fpp", a, [fc, "Accumulated", mon]), _var(x.better, 4, 5),
                    self.pct(fv["ytd"], x.better, 4, 5), self.flag("forecast", x.key)]})
                self.expect("fpp", a, [fc, "Periodic", mon], f"{b}.forecast.m")
                self.expect("fpp", a, [fc, "Accumulated", mon], f"{b}.forecast.ytd")
        bl.append({"key": "recon", "label": "EBITDA: FinPlan actual less FinMax (definitions differ)", "kind": "line",
                   "aux": True, "cells": [{"expr": f"b_{self.gp}-{self.g}"}, None, None, None,
                                          {"expr": f"b_{self.gp}-{self.g}"}]})
        self.sections.append({"title": f"P&L vs budget (FinPlan; actual = Working, Target = Final): {mon} {fy}",
                              "lines": bl})
        self.checks += [{"label": c.label, "expr": _prefixed(c.expr, "b_"), "columns": list(VALUE_COLS)}
                        for c in self.cat.checks["FinPlan"]]
        if fl:
            self.sections.append({"title": f"P&L vs the latest forecast ({s.forecast}, Final)", "lines": fl})
            self.checks += [{"label": f"{c.label} (forecast)", "expr": _prefixed(c.expr, "f_"),
                             "columns": list(VALUE_COLS)} for c in self.cat.checks["FinPlan"]]

    # -------------------------------------------------------------- D. EBITDA bridge by line
    def line_bridge(self) -> None:
        lb, g, p, ly = self.res.line_bridge, self.g, self.p, self.ly
        if lb is None:
            return
        if any(sign < 0 for sign, _ in components(self.cat.checks["FinMax"], self.fm_lines, g)):
            raise ValueError("a negative leaf in the EBITDA chain is not supported by the line bridge")
        leaf = [st for st in lb.steps if st.key not in ("rounding", "unexplained")]
        gap = next((st for st in lb.steps if st.key in ("rounding", "unexplained")), None)
        rows = [{"key": "lb_start", "label": f"EBITDA YTD {ly}", "kind": "line",
                 "cells": _at4({"ref": {"key": g, "col": 5}})}]
        rows += [{"key": f"lb_{st.key}", "label": st.label, "kind": "line",
                  "cells": _at4({"ref": {"key": st.key, "col": 6}})} for st in leaf]
        steps = ["lb_start"] + [f"lb_{st.key}" for st in leaf]
        if gap is not None:
            rows += [{"key": "lb_steps", "label": "Sum of the line steps", "kind": "line", "aux": True,
                      "cells": _at4({"sum": [f"lb_{st.key}" for st in leaf]})},
                     {"key": f"lb_{gap.key}", "label": gap.label + " (end - start - the line steps)", "kind": "line",
                      "cells": _at4({"expr": "lb_end-lb_start-lb_steps"})}]
            steps.append(f"lb_{gap.key}")
        rows.append({"key": "lb_end", "label": f"EBITDA YTD {p}", "kind": "sum",
                     "cells": _at4({"ref": {"key": g, "col": 4}})})
        rows.append({"key": "lb_sum", "label": "Start + steps (check)", "kind": "line", "aux": True,
                     "cells": _at4({"sum": steps})})
        self.sections.append({"title": f"EBITDA bridge by line, YTD {ly} → YTD {p} (FinMax)", "lines": rows})
        self.checks.append({"label": "Line bridge closes", "expr": "lb_sum-lb_end", "columns": [4]})
        if lb.unexplained:
            self.checks.append({"label": "Unexplained (line bridge)", "static": "DIFF",
                                "detail": f"NOK {_fmt(lb.unexplained)} not explained by the identity chain"})
        self.chart = {"type": "waterfall", "lines": steps + ["lb_end"], "column": 4,
                      "title": f"EBITDA bridge YTD {ly} → YTD {p}"}

    # -------------------------------------------------------------- E. EBITDA bridge by sub-unit
    def unit_bridge(self) -> None:
        res, s, g, p, ly = self.res, self.res.scope, self.g, self.p, self.ly
        ub = res.unit_bridge
        if ub is None or not s.units:
            return
        acct = next(x.account for x in self.fm_lines if x.key == g)
        self.pulls.append({"id": "units", "connection": "FinMax",
                           "pov": {"All Scenarios": "Actual", "All Currencies": "NOK", "All Accounts": acct,
                                   "All Balance Types": "Accumulated"},
                           "rows": [_ax("All Entities", [u for u, _ in s.units])],
                           "columns": [_ax("All Time Periods", [p, ly])]})
        self.expected["units"] = []
        for u, _ in s.units:
            self.expect("units", [u], [p], f"unit.{u}.cur")
            self.expect("units", [u], [ly], f"unit.{u}.ly")
        top = [st for st in ub.steps if st.key != "other"]
        other = next((st for st in ub.steps if st.key == "other"), None)
        small = self.m.pct_floor_nok

        def upct(fid):
            c = self.value(fid)
            return None if c is not None and abs(c) < small else {"varpct": [4, 5]}
        ul = [{"key": f"u{i}", "label": st.label, "kind": "line", "cells": _at4(
            _pull("units", [st.key], [p]), [_pull("units", [st.key], [ly]), {"var": [4, 5]},
                                            upct(f"unit.{st.key}.ly")])}
            for i, st in enumerate(top, 1)]
        tops = [f"u{i}" for i in range(1, len(top) + 1)]
        ul.append({"key": "u_top", "label": f"Sum of the {len(top)} units shown", "kind": "line", "aux": True,
                   "cells": _at4({"sum": tops}, [{"sum": tops}, {"var": [4, 5]}])})
        ul.append({"key": "u_other", "label": "Other units and eliminations", "kind": "line",
                   "cells": _at4({"expr": f"{g}-u_top"}, [{"expr": f"{g}-u_top"}, {"var": [4, 5]}])})
        self.sections.append({"title": f"EBITDA by sub-unit, YTD {p} vs YTD {ly} (largest variances)", "lines": ul})
        rows = [{"key": "ub_start", "label": f"EBITDA YTD {ly}", "kind": "line",
                 "cells": _at4({"ref": {"key": g, "col": 5}})}]
        rows += [{"key": f"ub{i}", "label": st.label, "kind": "line",
                  "cells": _at4({"ref": {"key": f"u{i}", "col": 6}})} for i, st in enumerate(top, 1)]
        steps = ["ub_start"] + [f"ub{i}" for i in range(1, len(top) + 1)]
        if other is not None:
            rows.append({"key": "ub_other", "label": other.label, "kind": "line",
                         "cells": _at4({"ref": {"key": "u_other", "col": 6}})})
            steps.append("ub_other")
        rows.append({"key": "ub_end", "label": f"EBITDA YTD {p}", "kind": "sum",
                     "cells": _at4({"ref": {"key": g, "col": 4}})})
        rows.append({"key": "ub_sum", "label": "Start + steps (check)", "kind": "line", "aux": True,
                     "cells": _at4({"sum": steps})})
        self.sections.append({"title": f"EBITDA bridge by sub-unit, YTD {ly} → YTD {p}", "lines": rows})
        self.checks.append({"label": "Sub-unit bridge closes", "expr": "ub_sum-ub_end", "columns": [4]})

    # -------------------------------------------------------------- F. Year-end landing
    def landing(self) -> None:
        res, s, mon, fy, gp = self.res, self.res.scope, self.mon, self.fy, self.gp
        ld = res.landing
        if ld is None or not (s.entity.finplan and res.budget):
            return
        acct = next(x.account for x in self.fp_lines if x.key == gp)
        boi = self.cat.line_by_concept("FinPlan", "ebitda_boi")
        lines = [{"key": "l_ytd", "label": f"Actual YTD {mon} {fy}", "kind": "line",
                  "cells": _at4({"ref": {"key": f"b_{gp}", "col": 4}})}]
        has_fc, fc_ok = bool(res.forecast), ld.forecast.value is not None
        if has_fc:
            fc = s.forecast
            lines += [
                {"key": "l_fc_fy", "label": f"{fc}: full year", "kind": "line", "aux": True,
                 "cells": _at4(_pull("fpp", [acct], [fc, "Accumulated", "Dec"]))},
                {"key": "l_fc_ytd", "label": f"{fc}: YTD {mon}", "kind": "line", "aux": True,
                 "cells": _at4({"ref": {"key": f"f_{gp}", "col": 5}})},
                {"key": "l_rem", "label": "Remaining months of the forecast", "kind": "line",
                 "cells": _at4({"expr": "l_fc_fy-l_fc_ytd"} if fc_ok else {"text": "n/a"})},
                {"key": "l_fc", "label": "Landing: actual YTD + remaining forecast", "kind": "sum",
                 "cells": _at4({"expr": "l_ytd+l_rem"} if fc_ok else {"text": "n/a"})}]
            self.expect("fpp", [acct], [fc, "Accumulated", "Dec"], f"fp.{gp}.forecast.fy")
            if not fc_ok:
                self.method_notes.append(f"Landing with the forecast: n/a ({ld.forecast.reason}).")
        lines.append({"key": "l_tgt", "label": "Target (budget): full year", "kind": "line",
                      "cells": _at4(_pull("fpp", [acct], ["Target", "Accumulated", "Dec"]))})
        self.expect("fpp", [acct], ["Target", "Accumulated", "Dec"], f"fp.{gp}.budget.fy")
        if has_fc:
            lines.append({"key": "l_gap", "label": "Landing less target", "kind": "line",
                          "cells": _at4({"expr": "l_fc-l_tgt"} if fc_ok else {"text": "n/a"})})
        if ld.seasonal and boi is not None:
            ba, name = boi.account, self.label("FinPlan", boi.key)
            lines += [
                {"key": "l_ly_ytd", "label": f"Last year: {name} YTD {mon} {self.fy_ly}", "kind": "line",
                 "aux": True, "cells": _at4(_pull("fpa", [ba], [self.fy_ly, "Accumulated", mon]))},
                {"key": "l_ly_fy", "label": f"Last year: {name} full year {self.fy_ly}", "kind": "line",
                 "aux": True, "cells": _at4(_pull("fpa", [ba], [self.fy_ly, "Accumulated", "Dec"]))},
                {"key": "l_share", "label": "Last year's YTD share of its full year", "kind": "line", "aux": True,
                 "num": "prosent1", "cells": _at4({"expr": "l_ly_ytd/l_ly_fy"})},
                {"key": "l_rr", "label": "Run-rate landing (YTD / last year's share)", "kind": "sum",
                 "cells": _at4({"expr": "l_ytd/l_share"})}]
            self.expect("fpa", [ba], [self.fy_ly, "Accumulated", mon], f"fp.{boi.key}.actual.ly_ytd")
            self.expect("fpa", [ba], [self.fy_ly, "Accumulated", "Dec"], f"fp.{boi.key}.actual.ly_fy")
        else:
            lines.append({"key": "l_rr", "label": f"Run-rate landing (YTD / {len(res.ytd)} x 12)", "kind": "sum",
                          "cells": _at4({"expr": "l_ytd/months_ytd*months_year"})})
        lines.append({"key": "l_rr_gap", "label": "Run-rate landing less target", "kind": "line",
                      "cells": _at4({"expr": "l_rr-l_tgt"})})
        self.sections.append({"title": f"Year-end landing {fy} (FinPlan EBITDA; values in the YTD column)",
                              "lines": lines})
        self.method_notes.append(f"Landing method: {ld.method}.")

    # -------------------------------------------------------------- notes
    def missing_notes(self) -> list[str]:
        """Missing figures: those on a source sheet show as DIFF in the live read = sheet checks (counted); the
        others (derived) are listed with their reasons."""
        on_sheet = {e[3] for cells in self.expected.values() for e in cells}
        gone = self.res.missing()
        sheet = [f for f in gone if f.id in on_sheet]
        other = [f for f in gone if f.id not in on_sheet]
        out = []
        if sheet:
            src = self.source.lower()
            out.append(f"{len(sheet)} source figure(s) have no {src} value; they show as DIFF in the {src} = "
                       f"source sheet checks (first: {sheet[0].id}: {sheet[0].reason}).")
        if other:
            shown = "; ".join(f"{f.id}: {f.reason}" for f in other[:6])
            more = f"; and {len(other) - 6} more" if len(other) > 6 else ""
            out.append(f"Not computed ({len(other)}): {shown}{more}.")
        return out

    def notes(self, mode: str, refresh_notes) -> list[str]:
        """Order: open period, forecast warnings, problems, missing figures, then the method notes."""
        res, m = self.res, self.m
        where = self.source.lower()

        def nok(v):
            return f"NOK {v:,.0f}".replace(",", " ")
        head = [n for n in [open_period_note(res)] if n] + finplan_warnings(res) + landing_warnings(res)
        head += [f"Problem: {x}" for x in res.problems] + self.missing_notes()
        method = [freshness_text(res, mode, refresh_notes), *res.notes, *self.method_notes,
                  "Signs as in FinMax: costs are negative; variance = current − comparison (positive = better). Cost "
                  "per FTE is shown positive (x the Sign input), so its variance is comparison − current.",
                  f"Material: the YTD variance at {m.pct:.0%} of |comparison| (counted from {nok(m.pct_floor_nok)}; "
                  f"a smaller comparison counts as zero) {'or' if m.rule == 'either' else 'and'} at {nok(m.abs_nok)}, "
                  f"or a single month moving at least {nok(m.abs_nok)}; detail lines only (subtotals are not flagged "
                  f"again). 'comment' = always commented ({', '.join(m.always)}); 'unknown' = could not be judged. "
                  f"The flags are set by smartass from the {where}; they do not change when the sheets are "
                  "refreshed." + (f" Flags marked provisional: {PROVISIONAL}." if res.provisional else ""),
                  f"Variance % is left blank where |comparison| is under {nok(m.pct_floor_nok)}.",
                  f"The '{self.source} = source sheet' checks are static: they compare the sheets as built with the "
                  f"{where} and do not re-run when the sheets are refreshed."]
        if res.scope.entity.finplan:
            method += ["Last year, budget, forecast and the landing are always shown; the question's comparisons "
                       "only decide which are flagged and commented.",
                       "Actual vs budget and forecast is read entirely in FinPlan (its definitions differ from FinMax; "
                       "the EBITDA difference line shows by how much)."]
        return head + method + ["History is restated to today's organisation (e.g. the 1 June 2026 reorg)."]


def build(res: Result, cat: Catalog, now: str | None = None, mode: str = "live read",
          refresh_notes: list[str] | None = None) -> Request:
    b = _Builder(res, cat)
    b.finmax()
    b.kpis()
    b.finplan()
    b.line_bridge()
    b.unit_bridge()
    b.landing()
    s, p = res.scope, res.scope.period
    closed_ok = not periods.after(p, s.closed)
    b.checks.append({"label": "Period is closed", "static": "OK" if closed_ok else "DIFF",
                     "detail": f"{p} vs Current Reporting Month {s.closed}"})
    plan = bool(s.entity.finplan)                                 # most units below an area have no FinPlan member
    spec = {"title": f"{s.entity.label} — P&L {p} and YTD vs last year" + (", budget and forecast" if plan else ""),
            "subtitle": (f"NOK · FinMax and FinPlan · {s.entity.finmax} / {s.entity.finplan}" if plan
                         else f"NOK · FinMax · {s.entity.finmax} (no FinPlan entity)"),
            "auto_points": False, "key_points": key_points(res, cat),
            "columns": [{"header": h.format(p=p), "kind": k} for h, k in zip(COLUMNS, KINDS)],
            "sections": b.sections, "checks": b.checks, "notes": b.notes(mode, refresh_notes),
            "source": f"Source: Oracle EPM Cloud ({'FinMax, FinPlan' if plan else 'FinMax'}) via Smart View, retrieved "
                      f"{now or dt.datetime.now().strftime('%Y-%m-%d %H:%M')}.",
            "chart": b.chart}
    return Request(f"{s.entity.label} P&L {p}", b.pulls, spec, b.expected, b.source)


def _kpi_text(v: float | None, unit: str) -> str:
    if v is None:
        return "n/a"
    return f"{v * 100:.1f} %".replace(".", ",") if unit == "pct" else _fmt(v)


def _pct_text(p: float | None) -> str:
    return "" if p is None else ", " + f"{p * 100:+.1f} %".replace(".", ",").replace("-", "−")


def key_points(res: Result, cat: Catalog) -> list[str]:
    """Factual points drafted from the facts (interpretation is the agent's, Stage 5). Labels are those of each
    comparison's cube; an open period's flags are provisional."""
    from .fin_pnl import CUBE
    m = cat.materiality
    g, gp = grand(cat.pnl["FinMax"]), grand(cat.pnl["FinPlan"])
    prov = f" ({PROVISIONAL})" if res.provisional else ""
    out = []
    v = res.fm.get(g, {}).get("ytd")
    if v is not None:
        out.append(f"EBITDA YTD {res.scope.period}: {_fmt(v.current.value)} vs {_fmt(v.comparison.value)} YTD "
                   f"{res.ly} (variance {_fmt(v.abs.value)}{_pct_text(shown_pct(v, m))}){prov}.")
    if res.headline and res.headline in res.kpi_vars:
        k = next(x for x in cat.kpis if x.key == res.headline)
        hv = res.kpi_vars[res.headline]["ytd"]
        ann = " (annualised)" if k.annualise else ""
        out.append(f"{k.label} YTD {res.scope.period}{ann}: {_kpi_text(hv.current.value, k.unit)} vs "
                   f"{_kpi_text(hv.comparison.value, k.unit)} YTD {res.ly} (variance {_kpi_text(hv.abs.value, k.unit)}"
                   f"; {'lower' if k.better == 'lower' else 'higher'} is better).")
    b = res.budget.get(gp, {}).get("ytd")
    if b is not None:
        out.append(f"Against budget (FinPlan): EBITDA YTD {_fmt(b.current.value)} vs target {_fmt(b.comparison.value)} "
                   f"(variance {_fmt(b.abs.value)}{_pct_text(shown_pct(b, m))}){prov}.")
    mat = [(c, k, f) for c, k, f in res.ranking() if f.material][:3]
    if mat:
        parts = []
        for c, k, f in mat:
            vs = res.table_of(c)[k]
            name = res.label(CUBE[c], k)
            if f.basis == "month":
                parts.append(f"{name} vs {NAMES[c]}: the month {_fmt(vs['m'].abs.value)} (a single-month swing)")
            else:
                parts.append(f"{name} vs {NAMES[c]} {_fmt(vs['ytd'].abs.value)}"
                             + (" (YTD and the month)" if f.basis == "ytd+month" else ""))
        out.append("Material variances: " + "; ".join(parts) + f"{prov}.")
    ld = res.landing
    if ld is not None and ld.run_rate.value is not None:
        text = f"Year-end landing {periods.finplan(res.scope.period)[1]}: "
        if ld.forecast.value is not None:
            text += f"{_fmt(ld.forecast.value)} with the latest forecast ({res.scope.forecast}), "
        text += f"{_fmt(ld.run_rate.value)} at run rate; target {_fmt(ld.target.value)}."
        if res.scope.forecast_note:
            text += f" {res.scope.forecast_note}."
        if ld.warning:
            text += f" Warning: {ld.warning}."
        out.append(text)
    ub = res.unit_bridge
    if ub is not None and ub.steps and ub.steps[0].key != "other" and ub.steps[0].fact.value is not None:
        out.append(f"Largest sub-unit EBITDA variance vs YTD {res.ly}: {ub.steps[0].label} "
                   f"{_fmt(ub.steps[0].fact.value)}.")
    if res.line_bridge is not None and res.line_bridge.unexplained:
        out.append(f"Unexplained: the EBITDA line bridge leaves {_fmt(res.line_bridge.unexplained)}.")
    return out


def _sheet_value(p, row: list[str], col: list[str]):
    """The pulled cell as Smart View put it on the source sheet: a number, 0.0 for blank / #Missing, a label such as
    '#NoAccess', or KeyError when the sheet has no such cell."""
    j = p.layout.col_tuples.index(tuple(col))
    for r in p.parsed.rows:
        if tuple(r[:p.layout.nr]) == tuple(row):
            v = r[p.layout.nr + j]
            return 0.0 if v is None or (isinstance(v, str) and v.strip() in ("#Missing", "")) else v
    raise KeyError(tuple(row))


def reconcile_pull(p, cells: list[Expected], tol: float = TOL, source: str = "Live read") -> dict:
    """The static check '<source> = source sheet' for one pull (each cell within its own tolerance: NOK 1, FTE
    0.05). A figure the live read could not give is a DIFF."""
    src = source.lower()
    worst, bad = 0.0, []
    for cell in cells:
        row, col, want, fid = cell[:4]
        t = cell[4] if len(cell) > 4 else tol
        where = f"{'/'.join(row)} {'/'.join(col)}"
        try:
            got = _sheet_value(p, row, col)
        except (KeyError, ValueError):
            bad.append(f"{where} not on the sheet")
            continue
        if want is None:
            bad.append(f"{where}: no {src} value ({fid})")
            continue
        if not isinstance(got, (int, float)):
            bad.append(f"{where}: the sheet shows {got}")
            continue
        d = abs(float(got) - want)
        worst = max(worst, d)
        if d > t:
            bad.append(f"{where}: sheet {_fmt(got)}, {src} {_fmt(want)}")
    detail = f"{len(cells)} cells, largest difference {worst:,.2f}".replace(",", " ")
    if bad:
        detail += f"; {len(bad)} off, first: {bad[0]}"
    return {"label": f"{source} = source sheet '{p.sheet}'", "static": "DIFF" if bad else "OK", "detail": detail}


def with_reconciliation(spec: dict, expected: dict, pulls_by_id: dict, source: str = "Live read") -> dict:
    """The before-answer hook: one '<source> = source sheet' check per source sheet (source:
    Request.source), after the others. Returns a new spec dict (the hook must return the spec)."""
    extra = [reconcile_pull(pulls_by_id[pid], cells, source=source) for pid, cells in expected.items()
             if pid in pulls_by_id]
    return {**spec, "checks": list(spec.get("checks") or []) + extra}
