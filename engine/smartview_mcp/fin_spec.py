"""The workbook for fin_pnl: the sv_ws_build request (live Smart View source sheets + an answer spec whose every
number is a formula over them) and the expected value of every source cell the answer uses, from the calc engine's
facts - so the build can prove warehouse = sheet (with_reconciliation, run as _build_job's before_answer hook).

Answer columns (one set for every section; the section title says what the comparison is):
  0 month | 1 comparison month | 2 variance | 3 variance % | 4 YTD | 5 comparison YTD | 6 variance | 7 variance % |
  8 material (YTD). Bridges and the landing use column 4.

The sheet follows the engine's rules exactly (specs.evaluate gives the engine's facts):
- variance = current - comparison, positive = better; for a better = 'lower' KPI (cost per FTE) the variance
  column is comparison - current (the var columns are swapped) and has no % (its base would differ);
- a KPI with sign = -1 is multiplied by the labelled input cell 'Sign' (-1), so cost per FTE is shown positive;
- no constants in formulas (tn-formelstil): months in a year / in the period and the sign are labelled inputs;
- averages of a stock (FTE) are AVERAGEA over the opening level + each month end (n+1 points)."""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass

from .analytics import periods
from .analytics.bridge import components
from .analytics.reconcile import TOL
from .catalog import Catalog
from .fin_pnl import Result, freshness_text, grand, ordered_kpis
from .templates import _fmt, _pct

COLUMNS = ["{p}", "Comparison", "Variance", "Variance %", "YTD {p}", "Comparison YTD", "Variance YTD",
           "Variance % YTD", "Material (YTD)"]
KINDS = ["heltall", "heltall", "heltall", "prosent1", "heltall", "heltall", "heltall", "prosent1", "heltall"]
VALUE_COLS = (0, 1, 4, 5)
INPUT_KEYS = ("months_year", "months_ytd", "kpi_sign")      # the only lines with typed values (labelled inputs)


@dataclass
class Request:
    question: str
    pulls: list[dict]                                        # PullIn fields: id, connection, rows, columns, pov
    spec: dict
    expected: dict[str, list[tuple[list[str], list[str], float | None, str]]]   # pull id -> (row, col, value, fact)


def _ax(dim: str, members: list[str]) -> dict:
    return {"dimension": dim, "members": list(dict.fromkeys(members))}


def _pull(pull: str, row: list[str], col: list[str]) -> dict:
    return {"pull": pull, "row": row, "col": col}


def _var(better: str, a: int, b: int) -> dict:
    """The variance cell of columns a (current) and b (comparison): positive = better."""
    return {"var": [b, a]} if better == "lower" else {"var": [a, b]}


def _varpct(better: str, a: int, b: int) -> dict | None:
    return None if better == "lower" else {"varpct": [a, b]}


def _flag_text(res: Result, comparison: str, key: str):
    f = res.flags.get(comparison, {}).get(key)
    if f is None or not f.comment:
        return None
    if not f.material:
        return {"text": "comment"}
    return {"text": {"ytd": "material", "month": "material (month)", "ytd+month": "material (YTD + month)"}.get(
        f.basis, "material")}


def _value(res: Result, fid: str):
    f = res.facts.get(fid)
    return None if f is None else f.value


def _prefixed(expr: str, prefix: str) -> str:
    return re.sub(r"[a-z_]\w*", lambda m: prefix + m.group(0), expr)


def _at4(cell, more=()) -> list:
    """A line's cells with `cell` in column 4 (bridges, landing), then `more`."""
    return [None] * 4 + [cell, *more]


def build(res: Result, cat: Catalog, now: str | None = None, mode: str = "warehouse",
          refresh_notes: list[str] | None = None) -> Request:
    s = res.scope
    p, ly = s.period, res.ly
    mon, fy = periods.finplan(p)
    fy_ly = periods.finplan(ly)[1]
    fm_lines, fp_lines = cat.pnl["FinMax"], cat.pnl["FinPlan"]
    g, gp = grand(fm_lines), grand(fp_lines)
    stocks = [x for x in fm_lines if x.measure == "stock"]
    e = s.entity.finmax
    base = {"All Entities": e, "All Scenarios": "Actual", "All Currencies": "NOK"}
    pulls = [{"id": "pnl", "connection": "FinMax", "pov": base,
              "rows": [_ax("All Accounts", [x.account for x in fm_lines])],
              "columns": [_ax("All Balance Types", ["Periodic", "Accumulated"]), _ax("All Time Periods", [p, ly])]}]
    expected: dict[str, list] = {"pnl": []}
    cur_pts = [periods.opening(p), *res.ytd]
    ly_pts = [periods.opening(ly), *res.ly_ytd]
    if stocks:
        pulls.append({"id": "fte", "connection": "FinMax", "pov": {**base, "All Balance Types": "Accumulated"},
                      "rows": [_ax("All Accounts", [x.account for x in stocks])],
                      "columns": [_ax("All Time Periods", cur_pts + ly_pts)]})
        expected["fte"] = []
    sections, checks, notes, warnings = [], [], [], []

    # ---- A. FinMax P&L vs last year
    lines = []
    for x in fm_lines:
        a = [x.account]
        if x.measure == "stock":
            cells = [_pull("pnl", a, ["Accumulated", p]), _pull("pnl", a, ["Accumulated", ly]),
                     _var(x.better, 0, 1), _varpct(x.better, 0, 1), {"ref": {"key": f"{x.key}_avg", "col": 4}},
                     {"ref": {"key": f"{x.key}_avg", "col": 5}}, _var(x.better, 4, 5), _varpct(x.better, 4, 5),
                     None]
            expected["pnl"] += [(a, ["Accumulated", p], _value(res, f"fm.{x.key}.m.cur"), f"fm.{x.key}.m.cur"),
                                (a, ["Accumulated", ly], _value(res, f"fm.{x.key}.m.ly"), f"fm.{x.key}.m.ly")]
            label = f"{x.label} (month end; YTD = average)"
        else:
            cells = [_pull("pnl", a, ["Periodic", p]), _pull("pnl", a, ["Periodic", ly]), _var(x.better, 0, 1),
                     _varpct(x.better, 0, 1), _pull("pnl", a, ["Accumulated", p]),
                     _pull("pnl", a, ["Accumulated", ly]), _var(x.better, 4, 5), _varpct(x.better, 4, 5),
                     _flag_text(res, "ly", x.key)]
            for bal, per, part in (("Periodic", p, "m.cur"), ("Periodic", ly, "m.ly"), ("Accumulated", p, "ytd.cur"),
                                   ("Accumulated", ly, "ytd.ly")):
                fid = f"fm.{x.key}.{part}"
                expected["pnl"].append((a, [bal, per], _value(res, fid), fid))
            label = x.label
        ln = {"key": x.key, "label": label, "kind": x.kind, "cells": cells}
        if x.num:
            ln["num"] = x.num
        lines.append(ln)
    sections.append({"title": f"P&L vs last year (FinMax): {p} vs {ly}, YTD vs YTD {ly}", "lines": lines})
    checks += [{"label": c.label, "expr": c.expr, "columns": list(VALUE_COLS)} for c in cat.checks["FinMax"]]

    # ---- FTE levels for the n+1 average
    for x in stocks:
        lv = []
        for i, (c, l) in enumerate(zip(cur_pts, ly_pts)):
            lv.append({"key": f"{x.key}_l{i}", "label": f"{x.label} {c} / {l}", "kind": "line", "aux": True,
                       "num": x.num or "tall1",
                       "cells": [None] * 4 + [_pull("fte", [x.account], [c]), _pull("fte", [x.account], [l])]})
            expected["fte"] += [([x.account], [c], _value(res, f"fm.{x.key}.lvl.{c}"), f"fm.{x.key}.lvl.{c}"),
                                ([x.account], [l], _value(res, f"fm.{x.key}.lvl.{l}"), f"fm.{x.key}.lvl.{l}")]
        keys = [ln["key"] for ln in lv]
        lv.append({"key": f"{x.key}_avg", "label": f"Average {x.label} ({len(keys)} points: opening + month ends)",
                   "kind": "sum", "aux": True, "num": x.num or "tall1",
                   "cells": [None] * 4 + [{"avg": keys}, {"avg": keys}, _var(x.better, 4, 5),
                                          _varpct(x.better, 4, 5)]})
        sections.append({"title": f"{x.label} levels for the average (opening December + each month end)",
                         "lines": lv})

    # ---- KPIs (the headline KPI first) and their labelled inputs
    kl, need_months, need_sign = [], False, False
    for k in ordered_kpis(cat, "FinMax"):
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
                   {"key": "months_ytd", "label": f"Months in the period (month; YTD {p} / YTD {ly})", "kind": "line",
                    "aux": True, "cells": [{"value": 1}, {"value": 1}, None, None, {"value": len(res.ytd)},
                                           {"value": len(res.ly_ytd)}]}]
    if need_sign:
        inputs.append({"key": "kpi_sign", "label": "Sign: cost per FTE shown positive (FinMax books costs negative)",
                       "kind": "line", "aux": True,
                       "cells": [{"value": -1}, {"value": -1}, None, None, {"value": -1}, {"value": -1}]})
    if inputs:
        sections.append({"title": "Inputs", "lines": inputs})
    if kl:
        sections.append({"title": "KPIs (FinMax; variance in points or NOK, positive = better)", "lines": kl})

    # ---- B/C. FinPlan vs budget and the latest forecast
    ef = s.entity.finplan
    boi = cat.line_by_concept("FinPlan", "ebitda_boi")
    if ef and res.budget:
        months = [mon] if mon == "Dec" else [mon, "Dec"]
        pulls.append({"id": "fpa", "connection": "FinPlan",
                      "pov": {"Entity": ef, "Scenario": "Actual", "Version": "Working"},
                      "rows": [_ax("Account", [x.account for x in fp_lines])],
                      "columns": [_ax("Years", [fy, fy_ly]), _ax("All Balance Types", ["Periodic", "Accumulated"]),
                                  _ax("Period", months)]})
        scen = ["Target"] + ([s.forecast] if s.forecast and res.forecast else [])
        pulls.append({"id": "fpp", "connection": "FinPlan", "pov": {"Entity": ef, "Years": fy, "Version": "Final"},
                      "rows": [_ax("Account", [x.account for x in fp_lines])],
                      "columns": [_ax("Scenario", scen), _ax("All Balance Types", ["Periodic", "Accumulated"]),
                                  _ax("Period", months)]})
        expected["fpa"], expected["fpp"] = [], []
        bl, fl = [], []
        for x in fp_lines:
            a, b = [x.account], f"fp.{x.key}"
            bl.append({"key": f"b_{x.key}", "label": x.label, "kind": x.kind, "cells": [
                _pull("fpa", a, [fy, "Periodic", mon]), _pull("fpp", a, ["Target", "Periodic", mon]),
                _var(x.better, 0, 1), _varpct(x.better, 0, 1), _pull("fpa", a, [fy, "Accumulated", mon]),
                _pull("fpp", a, ["Target", "Accumulated", mon]), _var(x.better, 4, 5), _varpct(x.better, 4, 5),
                _flag_text(res, "budget", x.key)]})
            expected["fpa"] += [(a, [fy, "Periodic", mon], _value(res, f"{b}.actual.m"), f"{b}.actual.m"),
                                (a, [fy, "Accumulated", mon], _value(res, f"{b}.actual.ytd"), f"{b}.actual.ytd")]
            expected["fpp"] += [(a, ["Target", "Periodic", mon], _value(res, f"{b}.budget.m"), f"{b}.budget.m"),
                                (a, ["Target", "Accumulated", mon], _value(res, f"{b}.budget.ytd"), f"{b}.budget.ytd")]
            if len(scen) > 1:
                fc = s.forecast
                fl.append({"key": f"f_{x.key}", "label": x.label, "kind": x.kind, "cells": [
                    {"ref": {"key": f"b_{x.key}", "col": 0}}, _pull("fpp", a, [fc, "Periodic", mon]),
                    _var(x.better, 0, 1), _varpct(x.better, 0, 1), {"ref": {"key": f"b_{x.key}", "col": 4}},
                    _pull("fpp", a, [fc, "Accumulated", mon]), _var(x.better, 4, 5), _varpct(x.better, 4, 5),
                    _flag_text(res, "forecast", x.key)]})
                expected["fpp"] += [
                    (a, [fc, "Periodic", mon], _value(res, f"{b}.forecast.m"), f"{b}.forecast.m"),
                    (a, [fc, "Accumulated", mon], _value(res, f"{b}.forecast.ytd"), f"{b}.forecast.ytd")]
        bl.append({"key": "recon", "label": "EBITDA: FinPlan actual less FinMax (definitions differ)", "kind": "line",
                   "aux": True, "cells": [{"expr": f"b_{gp}-{g}"}, None, None, None, {"expr": f"b_{gp}-{g}"}]})
        sections.append({"title": f"P&L vs budget (FinPlan; actual = Working, Target = Final): {mon} {fy}",
                         "lines": bl})
        checks += [{"label": c.label, "expr": _prefixed(c.expr, "b_"), "columns": list(VALUE_COLS)}
                   for c in cat.checks["FinPlan"]]
        if fl:
            sections.append({"title": f"P&L vs the latest forecast ({s.forecast}, Final)", "lines": fl})

    # ---- D. EBITDA bridge by line (FinMax YTD)
    chart = None
    lb = res.line_bridge
    if lb is not None:
        if any(sign < 0 for sign, _ in components(cat.checks["FinMax"], fm_lines, g)):
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
        sections.append({"title": f"EBITDA bridge by line, YTD {ly} → YTD {p} (FinMax)", "lines": rows})
        checks.append({"label": "Line bridge closes", "expr": "lb_sum-lb_end", "columns": [4]})
        if lb.unexplained:
            checks.append({"label": "Unexplained (line bridge)", "static": "DIFF",
                           "detail": f"NOK {_fmt(lb.unexplained)} not explained by the identity chain"})
        chart = {"type": "waterfall", "lines": steps + ["lb_end"], "column": 4,
                 "title": f"EBITDA bridge YTD {ly} → YTD {p}"}

    # ---- E. EBITDA bridge by sub-unit
    ub = res.unit_bridge
    if ub is not None and s.units:
        acct = next(x.account for x in fm_lines if x.key == g)
        pulls.append({"id": "units", "connection": "FinMax",
                      "pov": {"All Scenarios": "Actual", "All Currencies": "NOK", "All Accounts": acct,
                              "All Balance Types": "Accumulated"},
                      "rows": [_ax("All Entities", [u for u, _ in s.units])],
                      "columns": [_ax("All Time Periods", [p, ly])]})
        expected["units"] = [r for u, _ in s.units for r in (
            ([u], [p], _value(res, f"unit.{u}.cur"), f"unit.{u}.cur"),
            ([u], [ly], _value(res, f"unit.{u}.ly"), f"unit.{u}.ly"))]
        top = [st for st in ub.steps if st.key != "other"]
        other = next((st for st in ub.steps if st.key == "other"), None)
        ul = [{"key": f"u{i}", "label": st.label, "kind": "line", "cells": _at4(
            _pull("units", [st.key], [p]), [_pull("units", [st.key], [ly]), {"var": [4, 5]}, {"varpct": [4, 5]}])}
            for i, st in enumerate(top, 1)]
        ul.append({"key": "u_top", "label": f"Sum of the {len(top)} units shown", "kind": "line", "aux": True,
                   "cells": _at4({"sum": [f"u{i}" for i in range(1, len(top) + 1)]},
                                 [{"sum": [f"u{i}" for i in range(1, len(top) + 1)]}, {"var": [4, 5]}])})
        ul.append({"key": "u_other", "label": "Other units and eliminations", "kind": "line",
                   "cells": _at4({"expr": f"{g}-u_top"}, [{"expr": f"{g}-u_top"}, {"var": [4, 5]},
                                                         {"varpct": [4, 5]}])})
        sections.append({"title": f"EBITDA by sub-unit, YTD {p} vs YTD {ly} (largest variances)", "lines": ul})
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
        sections.append({"title": f"EBITDA bridge by sub-unit, YTD {ly} → YTD {p}", "lines": rows})
        checks.append({"label": "Sub-unit bridge closes", "expr": "ub_sum-ub_end", "columns": [4]})

    # ---- F. Year-end landing (FinPlan EBITDA; seasonality on EBITDA before other items)
    if ld is not None and ef and res.budget:
        acct = next(x.account for x in fp_lines if x.key == gp)
        lines = [{"key": "l_ytd", "label": f"Actual YTD {mon} {fy}", "kind": "line",
                  "cells": _at4({"ref": {"key": f"b_{gp}", "col": 4}})}]
        has_fc = bool(res.forecast)
        if has_fc:
            fc = s.forecast
            fc_ok = ld.forecast.value is not None
            lines += [
                {"key": "l_fc_fy", "label": f"{fc}: full year", "kind": "line", "aux": True,
                 "cells": _at4(_pull("fpp", [acct], [fc, "Accumulated", "Dec"]))},
                {"key": "l_fc_ytd", "label": f"{fc}: YTD {mon}", "kind": "line", "aux": True,
                 "cells": _at4({"ref": {"key": f"f_{gp}", "col": 5}})},
                {"key": "l_rem", "label": "Remaining months of the forecast", "kind": "line",
                 "cells": _at4({"expr": "l_fc_fy-l_fc_ytd"} if fc_ok else {"text": "n/a"})},
                {"key": "l_fc", "label": "Landing: actual YTD + remaining forecast", "kind": "sum",
                 "cells": _at4({"expr": "l_ytd+l_rem"} if fc_ok else {"text": "n/a"})}]
            expected["fpp"].append(([acct], [fc, "Accumulated", "Dec"], _value(res, f"fp.{gp}.forecast.fy"),
                                    f"fp.{gp}.forecast.fy"))
            if not fc_ok:
                notes.append(f"Landing with the forecast: n/a ({ld.forecast.reason}).")
        lines.append({"key": "l_tgt", "label": "Target (budget): full year", "kind": "line",
                      "cells": _at4(_pull("fpp", [acct], ["Target", "Accumulated", "Dec"]))})
        expected["fpp"].append(([acct], ["Target", "Accumulated", "Dec"], _value(res, f"fp.{gp}.budget.fy"),
                                f"fp.{gp}.budget.fy"))
        if has_fc:
            lines.append({"key": "l_gap", "label": "Landing less target", "kind": "line",
                          "cells": _at4({"expr": "l_fc-l_tgt"} if ld.forecast.value is not None
                                        else {"text": "n/a"})})
        if ld.seasonal and boi is not None:
            ba = boi.account
            lines += [
                {"key": "l_ly_ytd", "label": f"Last year: {boi.label} YTD {mon} {fy_ly}", "kind": "line",
                 "aux": True, "cells": _at4(_pull("fpa", [ba], [fy_ly, "Accumulated", mon]))},
                {"key": "l_ly_fy", "label": f"Last year: {boi.label} full year {fy_ly}", "kind": "line",
                 "aux": True, "cells": _at4(_pull("fpa", [ba], [fy_ly, "Accumulated", "Dec"]))},
                {"key": "l_share", "label": "Last year's YTD share of its full year", "kind": "line", "aux": True,
                 "num": "prosent1", "cells": _at4({"expr": "l_ly_ytd/l_ly_fy"})},
                {"key": "l_rr", "label": "Run-rate landing (YTD / last year's share)", "kind": "sum",
                 "cells": _at4({"expr": "l_ytd/l_share"})}]
            expected["fpa"] += [
                ([ba], [fy_ly, "Accumulated", mon], _value(res, f"fp.{boi.key}.actual.ly_ytd"),
                 f"fp.{boi.key}.actual.ly_ytd"),
                ([ba], [fy_ly, "Accumulated", "Dec"], _value(res, f"fp.{boi.key}.actual.ly_fy"),
                 f"fp.{boi.key}.actual.ly_fy")]
        else:
            lines.append({"key": "l_rr", "label": f"Run-rate landing (YTD / {len(res.ytd)} x 12)", "kind": "sum",
                          "cells": _at4({"expr": "l_ytd/months_ytd*months_year"})})
        lines.append({"key": "l_rr_gap", "label": "Run-rate landing less target", "kind": "line",
                      "cells": _at4({"expr": "l_rr-l_tgt"})})
        sections.append({"title": f"Year-end landing {fy} (FinPlan EBITDA; values in the YTD column)",
                         "lines": lines})
        notes.append(f"Landing method: {ld.method}.")
        if ld.warning:
            warnings.append(f"Warning: {ld.warning} - the forecast landing may be out of date; the run-rate "
                            "landing is shown alongside.")

    closed_ok = not periods.after(p, s.closed)
    checks.append({"label": "Period is closed", "static": "OK" if closed_ok else "DIFF",
                   "detail": f"{p} vs Current Reporting Month {s.closed}"})
    m = cat.materiality

    def nok(v):
        return f"NOK {v:,.0f}".replace(",", " ")
    notes = warnings + [freshness_text(res, mode, refresh_notes),
             "Signs as in FinMax: costs are negative; variance = current − comparison (positive = better). Cost "
             "per FTE is shown positive (x the Sign input), so its variance is comparison − current.",
             f"Material: the YTD variance at {m.pct:.0%} of |comparison| (counted from {nok(m.pct_floor_nok)}; a "
             f"smaller comparison counts as zero) {'or' if m.rule == 'either' else 'and'} at {nok(m.abs_nok)}, or a "
             f"single month moving at least {nok(m.abs_nok)}; detail lines only (subtotals are not flagged again). "
             f"'comment' = always commented ({', '.join(m.always)}). The flags are set by smartass from the "
             "warehouse; they do not change when the sheets are refreshed.",
             "Actual vs budget and forecast is read entirely in FinPlan (its definitions differ from FinMax; the "
             "EBITDA difference line shows by how much).",
             "History is restated to today's organisation (e.g. the 1 June 2026 reorg)."] + notes
    notes += [f"Problem: {x}" for x in res.problems]
    missing = res.missing()
    if missing:
        notes.append(f"{len(missing)} figure(s) could not be computed from the warehouse (shown as DIFF in the "
                     f"warehouse = source sheet checks), first: {missing[0].id}: {missing[0].reason}.")
    if not closed_ok:
        notes.insert(0, f"{p} is after the Current Reporting Month ({s.closed}): the period is still open.")
    title = f"{s.entity.label} — P&L {p} and YTD vs last year, budget and forecast"
    spec = {"title": title,
            "subtitle": f"NOK · FinMax and FinPlan · {s.entity.finmax} / {s.entity.finplan or '-'}",
            "auto_points": False, "key_points": key_points(res, cat),
            "columns": [{"header": h.format(p=p), "kind": k} for h, k in zip(COLUMNS, KINDS)],
            "sections": sections, "checks": checks, "notes": notes,
            "source": f"Source: Oracle EPM Cloud (FinMax, FinPlan) via Smart View, retrieved "
                      f"{now or dt.datetime.now().strftime('%Y-%m-%d %H:%M')}.",
            "chart": chart}
    return Request(f"{s.entity.label} P&L {p}", pulls, spec, expected)


def _kpi_text(v: float | None, unit: str) -> str:
    if v is None:
        return "n/a"
    return f"{v * 100:.1f} %".replace(".", ",") if unit == "pct" else _fmt(v)


def key_points(res: Result, cat: Catalog) -> list[str]:
    """Factual points drafted from the facts (interpretation is the agent's, Stage 5)."""
    g, gp = grand(cat.pnl["FinMax"]), grand(cat.pnl["FinPlan"])
    labels = {**{x.key: x.label for x in cat.pnl["FinPlan"]}, **{x.key: x.label for x in cat.pnl["FinMax"]}}
    out = []
    v = res.fm.get(g, {}).get("ytd")
    if v is not None:
        out.append(f"EBITDA YTD {res.scope.period}: {_fmt(v.current.value)} vs {_fmt(v.comparison.value)} YTD "
                   f"{res.ly} (variance {_fmt(v.abs.value)}, {_pct(v.current.value, v.comparison.value)}).")
    if res.headline and res.headline in res.kpi_vars:
        k = next(x for x in cat.kpis if x.key == res.headline)
        hv = res.kpi_vars[res.headline]["ytd"]
        out.append(f"{k.label} YTD {res.scope.period} (annualised): {_kpi_text(hv.current.value, k.unit)} vs "
                   f"{_kpi_text(hv.comparison.value, k.unit)} YTD {res.ly} (variance {_kpi_text(hv.abs.value, k.unit)}"
                   f"; {'lower' if k.better == 'lower' else 'higher'} is better).")
    b = res.budget.get(gp, {}).get("ytd")
    if b is not None:
        out.append(f"Against budget (FinPlan): EBITDA YTD {_fmt(b.current.value)} vs target {_fmt(b.comparison.value)} "
                   f"(variance {_fmt(b.abs.value)}, {_pct(b.current.value, b.comparison.value)}).")
    names = {"ly": "last year", "budget": "budget", "forecast": "forecast"}
    mat = [(c, k, f) for c, k, f in res.ranking() if f.material][:3]
    if mat:
        out.append("Material YTD variances: " + "; ".join(
            f"{labels.get(k, k)} vs {names[c]} {_fmt(_var_value(res, c, k))}" for c, k, f in mat) + ".")
    ld = res.landing
    if ld is not None and ld.run_rate.value is not None:
        text = f"Year-end landing {periods.finplan(res.scope.period)[1]}: "
        if ld.forecast.value is not None:
            text += f"{_fmt(ld.forecast.value)} with the latest forecast, "
        text += f"{_fmt(ld.run_rate.value)} at run rate; target {_fmt(ld.target.value)}."
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


def _var_value(res: Result, comparison: str, key: str) -> float | None:
    src = {"ly": res.fm, "budget": res.budget, "forecast": res.forecast}[comparison]
    v = src.get(key, {}).get("ytd")
    return None if v is None else v.abs.value


def _sheet_value(p, row: list[str], col: list[str]):
    """The pulled cell as Smart View put it on the source sheet: a number, 0.0 for blank / #Missing, a label such as
    '#NoAccess', or KeyError when the sheet has no such cell."""
    j = p.layout.col_tuples.index(tuple(col))
    for r in p.parsed.rows:
        if tuple(r[:p.layout.nr]) == tuple(row):
            v = r[p.layout.nr + j]
            return 0.0 if v is None or (isinstance(v, str) and v.strip() in ("#Missing", "")) else v
    raise KeyError(tuple(row))


def reconcile_pull(p, cells: list[tuple[list[str], list[str], float | None, str]], tol: float = TOL) -> dict:
    """The static check 'warehouse = source sheet' for one pull. A figure the warehouse could not give is a DIFF."""
    worst, bad = 0.0, []
    for row, col, want, fid in cells:
        where = f"{'/'.join(row)} {'/'.join(col)}"
        try:
            got = _sheet_value(p, row, col)
        except (KeyError, ValueError):
            bad.append(f"{where} not on the sheet")
            continue
        if want is None:
            bad.append(f"{where}: no warehouse value ({fid})")
            continue
        if not isinstance(got, (int, float)):
            bad.append(f"{where}: the sheet shows {got}")
            continue
        d = abs(float(got) - want)
        worst = max(worst, d)
        if d > tol:
            bad.append(f"{where}: sheet {_fmt(got)}, warehouse {_fmt(want)}")
    detail = f"{len(cells)} cells, largest difference NOK {worst:,.0f}".replace(",", " ")
    if bad:
        detail += f"; {len(bad)} off, first: {bad[0]}"
    return {"label": f"Warehouse = source sheet '{p.sheet}'", "static": "DIFF" if bad else "OK", "detail": detail}


def with_reconciliation(spec: dict, expected: dict, pulls_by_id: dict) -> dict:
    """_build_job's before_answer hook: one 'warehouse = sheet' check per source sheet, after the others. Returns a
    new spec dict (the hook must return the spec)."""
    extra = [reconcile_pull(pulls_by_id[pid], cells) for pid, cells in expected.items() if pid in pulls_by_id]
    return {**spec, "checks": list(spec.get("checks") or []) + extra}
