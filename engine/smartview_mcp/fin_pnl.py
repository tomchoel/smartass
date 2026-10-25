"""fin_pnl, the first slice (Stage 2+3): a business area's P&L for a month and YTD against last year (FinMax) and
budget and the latest forecast (FinPlan: actual = version Working, plan = Final), with materiality flags, an EBITDA
bridge by line and by sub-unit, KPIs and the year-end landing - decided by the calc engine from the warehouse, in
milliseconds and without Excel. fin_spec.py turns the Result into the sv_ws_build request (live source sheets +
formulas); the engine's numbers then reconcile the sheets. templates.division_pnl is this slice's predecessor.

Fact ids (a contract with fin_spec.py and the live checks; unique within a Result):
- FinMax: fm.<line>.m.cur|m.ly|ytd.cur|ytd.ly, variances fm.<line>.m|ytd .var/.pct; FTE levels fm.<line>.lvl.<month>.
- FinPlan: fp.<line>.actual.m|ytd|ly_ytd|ly_fy, fp.<line>.budget|forecast.m|ytd|fy (+ .m/.ytd .var/.pct),
  fp.<grand>.forecast.remaining.
- Sub-units: unit.<member>.cur|ly; unit bridge steps units.<member>.var, units.other; line bridge bridge.rounding|unexplained.
- KPIs: kpi.<key>.ytd.cur|ytd.ly (+ m.cur|m.ly for basis 'period'), variances kpi.<key>.ytd|m .var/.pct.
- Landing: land.forecast|run_rate|gap_forecast|gap_run_rate."""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from .analytics import periods
from .analytics.bridge import Bridge, by_line, by_unit, components
from .analytics.facts import Fact, missing
from .analytics.frame import Frame, from_rows
from .analytics.kpis import kpi
from .analytics.landing import Landing, landing, remaining_forecast
from .analytics.materiality import Flag, flags
from .analytics.measures import amount, average_level, flow, level
from .analytics.variance import Variance, variance
from .catalog import Catalog, Entity
from .extract import FINMAX_SCENARIOS

COMPARISONS = ("ly", "budget", "forecast")


@dataclass(frozen=True)
class Scope:
    entity: Entity
    period: str                                     # 'Sep 26'
    closed: str                                     # the last closed month
    forecast: str | None                            # the latest 'Forecast Q<q><yy>' for the period's year
    units: tuple[tuple[str, str], ...] = ()         # (member, label) sub-units for the bridge
    compare: tuple[str, ...] = COMPARISONS


@dataclass
class Result:
    scope: Scope
    ly: str
    ytd: list[str]
    ly_ytd: list[str]
    facts: dict[str, Fact] = field(default_factory=dict)
    fm: dict[str, dict[str, Variance]] = field(default_factory=dict)       # line -> {"m", "ytd"} vs last year
    budget: dict[str, dict[str, Variance]] = field(default_factory=dict)   # FinPlan line -> {"m", "ytd"}
    forecast: dict[str, dict[str, Variance]] = field(default_factory=dict)
    flags: dict[str, dict[str, Flag]] = field(default_factory=dict)        # comparison -> line -> Flag
    line_bridge: Bridge | None = None
    unit_bridge: Bridge | None = None
    kpis: dict[str, dict[str, Fact]] = field(default_factory=dict)         # kpi -> {"m.cur", "ytd.ly", ...}
    kpi_vars: dict[str, dict[str, Variance]] = field(default_factory=dict)  # kpi -> {"m", "ytd"} vs last year
    headline: str | None = None                                           # the headline KPI key (shown first)
    landing: Landing | None = None
    freshness: dict = field(default_factory=dict)
    problems: list[str] = field(default_factory=list)                     # e.g. an unexplained bridge residual

    def ranking(self) -> list[tuple[str, str, Flag]]:
        """(comparison, line, flag) for every line to comment (material or always), largest |variance| first."""
        out = [(c, k, f) for c, fs in self.flags.items() for k, f in fs.items() if f.comment]
        return sorted(out, key=lambda x: -x[2].size)

    def missing(self) -> list[Fact]:
        """Every fact that could not be computed, with its reason (never a silent zero)."""
        return [f for f in self.facts.values() if f.value is None]


def scenario_keys(cat: Catalog, cube: str, forecast: str | None) -> dict[tuple[str, str], str]:
    """(scenario member, version) -> catalogue scenario key for analytics.frame.from_rows, from [scenarios]; the
    latest-forecast rule maps to `forecast` (left out when there is none)."""
    out = {}
    for s in cat.scenarios.values():
        if s.connection != cube:
            continue
        member = forecast if s.rule == "latest_forecast" else s.scenario
        if member:
            out[(member, s.version or "")] = s.key
    return out


def grand(lines) -> str:
    return next(x.key for x in lines if x.kind == "grand")


def ordered_kpis(cat: Catalog, connection: str = "FinMax") -> list:
    """The cube's KPIs, the headline KPI first (catalogue order otherwise)."""
    return sorted((k for k in cat.kpis if k.connection == connection), key=lambda k: not k.headline)


def _keep(res: Result, *facts: Fact) -> None:
    for f in facts:
        old = res.facts.get(f.id)
        if old is not None and old != f:
            raise ValueError(f"two different facts share the id '{f.id}'")
        res.facts[f.id] = f


def _cell(rows: dict, key: tuple, id: str, definition: str) -> Fact:
    if key not in rows:
        return missing(id, "nok", definition, f"no warehouse value for {' / '.join(key[3:6])}")
    v = rows[key]
    return Fact(id, 0.0 if v is None else float(v), "nok", definition, (key,))


def _flags(res: Result, cat: Catalog, name: str, table: dict, lines) -> None:
    fs = flags([v["ytd"] for v in table.values()], cat.materiality, months={k: v["m"] for k, v in table.items()},
               lines=lines)
    res.flags[name] = {f.line: f for f in fs}


def _finmax(cat: Catalog, s: Scope, store, res: Result) -> Frame:
    e, p, ly = s.entity.finmax, s.period, res.ly
    lines = cat.pnl["FinMax"]
    closed_only = not periods.after(p, s.closed)
    months = res.ytd + res.ly_ytd
    flows = [x.account for x in lines if x.measure == "flow"]
    stocks = [x.account for x in lines if x.measure == "stock"]
    rows = store.get("FinMax", [e], flows, months, "Actual", "", "Periodic", closed_only=closed_only)
    if stocks:
        rows |= store.get("FinMax", [e], stocks, months + [periods.opening(p), periods.opening(ly)], "Actual", "",
                          "Accumulated", closed_only=closed_only)
    fm = from_rows("FinMax", lines, rows, FINMAX_SCENARIOS)
    for x in lines:
        b = f"fm.{x.key}"
        m_cur = amount(fm, e, x.key, "actual", [p], f"{b}.m.cur")
        m_ly = amount(fm, e, x.key, "actual", [ly], f"{b}.m.ly")
        if x.measure == "stock":
            y_cur = average_level(fm, e, x.key, "actual", res.ytd, f"{b}.ytd.cur")
            y_ly = average_level(fm, e, x.key, "actual", res.ly_ytd, f"{b}.ytd.ly")
            for m in [periods.opening(p), *res.ytd, periods.opening(ly), *res.ly_ytd]:
                _keep(res, level(fm, e, x.key, "actual", m, f"{b}.lvl.{m}"))
        else:
            y_cur = flow(fm, e, x.key, "actual", res.ytd, f"{b}.ytd.cur")
            y_ly = flow(fm, e, x.key, "actual", res.ly_ytd, f"{b}.ytd.ly")
        res.fm[x.key] = {"m": variance(x, m_cur, m_ly, f"{b}.m"), "ytd": variance(x, y_cur, y_ly, f"{b}.ytd")}
        for v in res.fm[x.key].values():
            _keep(res, v.current, v.comparison, v.abs, v.pct)
    _flags(res, cat, "ly", res.fm, lines)
    return fm


def _finplan(cat: Catalog, s: Scope, store, res: Result) -> Frame | None:
    ef = s.entity.finplan
    if not ef:
        return None
    lines = cat.pnl["FinPlan"]
    accts = [x.account for x in lines]
    yy = periods.split_period(s.period)[1]
    year, last = periods.year(yy), periods.year(yy - 1)
    closed_only = not periods.after(s.period, s.closed)
    rows = store.get("FinPlan", [ef], accts, year + last, "Actual", "Working", "Periodic", closed_only=closed_only)
    rows |= store.get("FinPlan", [ef], accts, year, "Target", "Final", "Periodic")
    if s.forecast:
        rows |= store.get("FinPlan", [ef], accts, year, s.forecast, "Final", "Periodic")
    fp = from_rows("FinPlan", lines, rows, scenario_keys(cat, "FinPlan", s.forecast))
    for x in lines:
        b = f"fp.{x.key}"
        a = {"m": flow(fp, ef, x.key, "actual_finplan", [s.period], f"{b}.actual.m"),
             "ytd": flow(fp, ef, x.key, "actual_finplan", res.ytd, f"{b}.actual.ytd"),
             "ly_ytd": flow(fp, ef, x.key, "actual_finplan", res.ly_ytd, f"{b}.actual.ly_ytd"),
             "ly_fy": flow(fp, ef, x.key, "actual_finplan", last, f"{b}.actual.ly_fy")}
        _keep(res, *a.values())
        for name, table in (("budget", res.budget), ("forecast", res.forecast)):
            if name == "forecast" and not s.forecast:
                continue
            pm = {"m": flow(fp, ef, x.key, name, [s.period], f"{b}.{name}.m"),
                  "ytd": flow(fp, ef, x.key, name, res.ytd, f"{b}.{name}.ytd"),
                  "fy": flow(fp, ef, x.key, name, year, f"{b}.{name}.fy")}
            _keep(res, *pm.values())
            table[x.key] = {"m": variance(x, a["m"], pm["m"], f"{b}.{name}.m"),
                            "ytd": variance(x, a["ytd"], pm["ytd"], f"{b}.{name}.ytd")}
            for v in table[x.key].values():
                _keep(res, v.abs, v.pct)
    _flags(res, cat, "budget", res.budget, lines)
    if res.forecast:
        _flags(res, cat, "forecast", res.forecast, lines)
    g = grand(lines)
    rid = f"fp.{g}.forecast.remaining"
    rem = (remaining_forecast(fp, ef, g, "forecast", periods.remaining(s.period), rid) if s.forecast
           else missing(rid, "nok", "remaining months of the latest forecast", "no latest forecast for the year"))
    boi = cat.line_by_concept("FinPlan", "ebitda_boi")        # last year's seasonality: EBITDA before other items
    ly_ytd = res.facts[f"fp.{boi.key}.actual.ly_ytd"] if boi else missing(
        "fp.ebitda_boi.actual.ly_ytd", "nok", "last year's YTD EBITDA before other items", "no ebitda_boi line")
    ly_fy = res.facts[f"fp.{boi.key}.actual.ly_fy"] if boi else missing(
        "fp.ebitda_boi.actual.ly_fy", "nok", "last year's EBITDA before other items", "no ebitda_boi line")
    res.landing = landing(res.facts[f"fp.{g}.actual.ytd"], rem, res.facts[f"fp.{g}.budget.fy"], ly_ytd, ly_fy,
                          len(res.ytd), forecast_name=s.forecast, closed_month=s.closed)
    ld = res.landing
    _keep(res, rem, ld.forecast, ld.run_rate, ld.gap_forecast, ld.gap_run_rate)
    return fp


def _units(cat: Catalog, s: Scope, store, res: Result, total_ly: Fact, total_cur: Fact) -> None:
    lines = cat.pnl["FinMax"]
    g = grand(lines)
    acct = next(x.account for x in lines if x.key == g)
    ly = res.ly
    names = [u for u, _ in s.units]
    rows = store.get("FinMax", names, [acct], [s.period, ly], "Actual", "", "Accumulated",
                     closed_only=not periods.after(s.period, s.closed))
    units = []
    for u, label in s.units:
        cur = _cell(rows, ("FinMax", "Actual", "", u, acct, s.period, "Accumulated"), f"unit.{u}.cur",
                    f"FinMax {u} {g} YTD {s.period} (Accumulated)")
        cmp = _cell(rows, ("FinMax", "Actual", "", u, acct, ly, "Accumulated"), f"unit.{u}.ly",
                    f"FinMax {u} {g} YTD {ly} (Accumulated)")
        _keep(res, cur, cmp)
        units.append((u, label, cur, cmp))
    res.unit_bridge = by_unit(total_ly, total_cur, units)


def _kpis(cat: Catalog, s: Scope, fm: Frame, res: Result) -> None:
    e, ly = s.entity.finmax, res.ly
    for k in ordered_kpis(cat, "FinMax"):
        out = {"ytd.cur": kpi(fm, k, e, "actual", res.ytd, f"kpi.{k.key}.ytd.cur"),
               "ytd.ly": kpi(fm, k, e, "actual", res.ly_ytd, f"kpi.{k.key}.ytd.ly")}
        vs = {"ytd": variance(k, out["ytd.cur"], out["ytd.ly"], f"kpi.{k.key}.ytd")}
        if k.basis == "period":                     # per-FTE KPIs (average stock) only for the YTD
            out["m.cur"] = kpi(fm, k, e, "actual", [s.period], f"kpi.{k.key}.m.cur")
            out["m.ly"] = kpi(fm, k, e, "actual", [ly], f"kpi.{k.key}.m.ly")
            vs["m"] = variance(k, out["m.cur"], out["m.ly"], f"kpi.{k.key}.m")
        res.kpis[k.key], res.kpi_vars[k.key] = out, vs
        if k.headline:
            res.headline = k.key
        _keep(res, *out.values(), *(f for v in vs.values() for f in (v.abs, v.pct)))


def compute(cat: Catalog, s: Scope, store) -> Result:
    """Every number of the answer, from the warehouse (store: warehouse.Warehouse or MemoryWarehouse)."""
    if s.entity.abolished:
        raise ValueError(f"{s.entity.label} was abolished on {s.entity.abolished}: ask for its successors")
    if not s.entity.finmax:
        raise ValueError(f"{s.entity.label} has no FinMax entity")
    ly = periods.ly(s.period)
    res = Result(s, ly, periods.ytd(s.period), periods.ytd(ly))
    fm = _finmax(cat, s, store, res)
    _finplan(cat, s, store, res)
    for c in COMPARISONS:
        if c not in s.compare:
            res.flags.pop(c, None)
    lines = cat.pnl["FinMax"]
    g = grand(lines)
    total_ly, total_cur = res.facts[f"fm.{g}.ytd.ly"], res.facts[f"fm.{g}.ytd.cur"]
    leaves = components(cat.checks["FinMax"], lines, g)
    res.line_bridge = by_line(total_ly, total_cur, {k: v["ytd"] for k, v in res.fm.items()}, leaves,
                              {x.key: x.label for x in lines})
    if s.units:
        _units(cat, s, store, res, total_ly, total_cur)
    for b in (res.line_bridge, res.unit_bridge):
        if b is not None:
            _keep(res, *(st.fact for st in b.steps))
    if res.line_bridge.unexplained:
        res.problems.append(f"Unexplained: the EBITDA line bridge leaves NOK {res.line_bridge.unexplained:,.0f} "
                            "that the catalogue's identity chain does not explain".replace(",", " "))
    _kpis(cat, s, fm, res)
    res.freshness = {"FinMax": store.freshness("FinMax", [s.entity.finmax])}
    if s.entity.finplan:
        res.freshness["FinPlan"] = store.freshness("FinPlan", [s.entity.finplan])
    return res


def freshness_text(res: Result, mode: str, notes: list[str] | None = None) -> str:
    """'Data: warehouse; closed history to Sep 26, oldest cell read 2026-10-25 08:12; the latest forecast and any
    open month read live (newest 10:31).' notes: refresh()'s notes (a refresh skipped because another process is
    refreshing, a paused job), appended so the answer says the closed history may be older."""
    def ts(x):
        return dt.datetime.fromtimestamp(x).strftime("%Y-%m-%d %H:%M") if x else "never"
    stamps = [f for f in res.freshness.values() if f.get("rows")]
    oldest = min((f["oldest"] for f in stamps), default=None)
    newest = max((f["newest"] for f in stamps), default=None)
    text = (f"Data: {mode}; closed history to {res.scope.closed}, oldest cell read {ts(oldest)}; the latest "
            f"forecast and any open month read live (newest {ts(newest)}).")
    for n in notes or []:
        text += f" Note: {n.rstrip('.')}."
    return text


def summary(res: Result, mode: str, fills: list, notes: list[str] | None = None) -> dict:
    """What the agent writes the answer from: facts (ids, values, definitions), flags, missing facts with their
    reasons, landing and freshness (with refresh()'s notes)."""
    keep = ("fm.", "fp.", "kpi.", "land.", "units.", "unit.", "bridge.")
    facts = [f.as_dict(max_inputs=3) for k, f in res.facts.items()
             if k.startswith(keep) and ".lvl." not in k and not k.endswith(".pct")]
    flagged = [{"comparison": c, "line": k, "material": f.material, "always": f.always, "basis": f.basis,
                "reason": f.reason, "variance": f.size} for c, k, f in res.ranking()]
    out = {"area": res.scope.entity.label, "period": res.scope.period, "closed_month": res.scope.closed,
           "forecast": res.scope.forecast, "freshness": freshness_text(res, mode, notes), "flags": flagged,
           "headline_kpi": res.headline, "facts": facts,
           "missing": [{"id": f.id, "reason": f.reason} for f in res.missing()][:40],
           "problems": list(res.problems), "refresh_notes": list(notes or []),
           "fills": [{"job": r.fill.job, "result": r.fill.result, "cells": r.fill.cells, "seconds": r.fill.seconds,
                      "errors": r.fill.detail.get("errors", [])[:3], "problems": r.fill.detail.get("problems", [])[:3]}
                     for r in fills]}
    if res.landing is not None:
        ld = res.landing
        out["landing"] = {"forecast": ld.forecast.value, "run_rate": ld.run_rate.value, "target": ld.target.value,
                          "gap_forecast": ld.gap_forecast.value, "gap_run_rate": ld.gap_run_rate.value,
                          "method": ld.method, "stale": ld.stale, "warning": ld.warning}
    return out
