"""Analysis templates: plan the pulls, then turn the pulled data into an answer spec (see answer.py).

Rules agreed with the user (2026-10-13): actuals from FinMax; plan/forecast from FinPlan; actual-vs-plan entirely in
FinPlan (Actual = version Working, Target/forecast = version Final) with a FinMax EBITDA reconciliation line;
P&L questions compare with the same period last year by default; FinMax signs kept (costs negative), so
variance = current - comparison is positive when better.
"""
from __future__ import annotations

import contextlib
import datetime as dt
import logging
import re
from dataclasses import dataclass, field
from typing import Callable

from . import atlas, catalog, ops
from .grid import Axis
from .host import ExcelHost, HostError, InputError

log = logging.getLogger(__name__)

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


@dataclass
class PullSpec:
    id: str
    connection: str
    rows: list[Axis]
    cols: list[Axis]
    pov: dict
    title_suffix: str = ""


@dataclass
class Plan:
    title: str
    pulls: list[PullSpec]
    answer: Callable[[dict], dict]            # {pull id: sheets.Pull} -> answer spec
    followups: list[Callable] = field(default_factory=list)
    summary: Callable[[dict], dict] | None = None


# ------------------------------------------------------------------ periods
COM_CLOSED_MONTH_TTL_S = 2 * 3600       # a slow COM answer is reused this long (not the whole day)


def closed_month(h: ExcelHost, connection: str = "FinMax") -> str:
    """'Current Reporting Month' -> e.g. 'Sep 26'. Smart View's HTTP protocol first (svhttp, ~0.2 s, no Excel); the
    COM route (FinMax member ChildMbrName '[CurrRepMonth].[Y26M09]' in the hidden server Excel, ~10 s + a login on a
    fresh server) is the fallback. Cached on disk per connection as {month, at, source}: an HTTP answer is reused for
    config.CLOSED_MONTH_TTL_S, a (slow) COM answer for COM_CLOSED_MONTH_TTL_S (2 h). The old {date, <conn>: month} format is
    read as a COM-sourced entry of that date."""
    import json
    import os
    import time
    from . import config, svhttp
    cache = config.SVKIT / "closed_month.json"
    now = time.time()
    try:
        c = json.loads(cache.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        c = {}
    if not isinstance(c, dict):
        c = {}
    e = c.get(connection)
    if isinstance(e, str) and c.get("date") == dt.date.today().isoformat():      # old format: valid for that day
        return e
    if isinstance(e, dict) and e.get("month") and isinstance(e.get("at"), (int, float)):
        if e.get("source") == "http" and now - e["at"] < config.CLOSED_MONTH_TTL_S:
            return e["month"]
        if e.get("source") == "com" and now - e["at"] < COM_CLOSED_MONTH_TTL_S:
            return e["month"]
    val, source = None, "http"
    try:
        val = svhttp.shared().closed_month(connection)
    except (svhttp.SvHttpError, svhttp.DrillError) as ex:
        log.info("closed month over HTTP failed (%s); asking Smart View in Excel", ex)
    except Exception as ex:   # noqa: BLE001 - anything else is a bug worth seeing, but the COM route still answers
        log.warning("closed month over HTTP failed unexpectedly (%s); asking Smart View in Excel", ex, exc_info=True)
    if val is None:
        source = "com"
        info = h.on_host(lambda: ops.find_member(h, connection, "Current Reporting Month")).get("info") or {}
        m = re.search(r"Y(\d\d)M(\d\d)", str(info.get("ChildMbrName", "")))
        if not m:
            raise HostError("Could not resolve 'Current Reporting Month' in FinMax; pass the period explicitly "
                            "(e.g. 'Sep 26').")
        mo, yy = int(m.group(2)), int(m.group(1))
        if not 1 <= mo <= 12:
            raise HostError(f"Smart View in Excel returned the closed month '{m.group(0)}', which is not a month "
                            "(1-12); pass the period explicitly (e.g. 'Sep 26').")
        val = f"{MONTHS[mo - 1]} {yy:02d}"
        today = dt.date.today()
        back = (today.year - (2000 + yy)) * 12 + today.month - mo
        if back < 0 or back > 3:
            log.warning("closed month %s from Smart View in Excel looks implausible (%d months %s today); using it",
                        val, abs(back), "ahead of" if back < 0 else "behind")
    tmp = cache.with_name(f"{cache.name}.{os.getpid()}.tmp")
    try:
        c.pop("date", None)
        c[connection] = {"month": val, "at": now, "source": source}
        tmp.write_text(json.dumps(c), encoding="utf-8")
        os.replace(tmp, cache)
    except OSError:
        with contextlib.suppress(OSError):
            tmp.unlink(missing_ok=True)                  # no stray temp files when the replace fails
    return val


def split_period(p: str) -> tuple[int, int]:
    m = re.fullmatch(r"([A-Z][a-z]{2}) (\d\d)", p.strip())
    if not m or m.group(1) not in MONTHS:
        raise InputError(f"Period '{p}' is not a month like 'Sep 26'.")
    return MONTHS.index(m.group(1)) + 1, int(m.group(2))


def same_period_ly(p: str) -> str:
    mo, yy = split_period(p)
    return f"{MONTHS[mo - 1]} {yy - 1:02d}"


def months_to(p: str) -> list[str]:
    mo, yy = split_period(p)
    return [f"{MONTHS[i]} {yy:02d}" for i in range(mo)]


def _clean(label: str | None) -> str:
    """'C30 REVENUES (ac)' -> 'C30 REVENUES'; 'EBITDA EBITDA (ac)' -> 'EBITDA' (alias repeating the code)."""
    if not label:
        return ""
    s = re.sub(r"\s*\((ac|un|un\.[A-Za-z0-9]+|pd|pd\.[A-Za-z0-9]+|cp|fu|sy|pj\.[A-Za-z0-9]+)\)\s*$", "", str(label).strip())
    parts = s.split(" ", 1)
    if len(parts) == 2 and parts[0] == parts[1]:
        s = parts[0]
    return s


def _label_of(p, row_members, default):
    labs = p.row_labels.get(tuple(row_members))
    return _clean(labs[-1]) if labs else default


def _raw(p, row_members, col_members):
    """The parsed cell as is (a number, None or a label such as '#NoAccess')."""
    try:
        j = p.layout.col_tuples.index(tuple(col_members))
    except ValueError:
        return None
    for r in p.parsed.rows:
        if tuple(r[:p.layout.nr]) == tuple(row_members):
            return r[p.layout.nr + j]
    return None


def _val(p, row_members, col_members):
    """Parsed value from a pull (for ranking / key points; the sheet itself uses live formulas)."""
    v = _raw(p, row_members, col_members)
    return v if isinstance(v, (int, float)) else None


def _recon_point(ea, fm, entity: str | None, period: str, note: str | None) -> str:
    """FinPlan vs FinMax EBITDA key point; says why when there is nothing to reconcile (never 'within n/a')."""
    if fm is None:
        why = (note or "no FinMax entity").rstrip(".").split(";")[0]
        return f"FinMax reconciliation unavailable ({why[0].lower() + why[1:]})."
    raw = _raw(fm, ["ac.EBITDA"], [period])
    efm = raw if isinstance(raw, (int, float)) else None
    if efm is None:
        why = (f"FinMax returned {raw} for {entity} {period}: no read access or member not valid in this cube"
               if isinstance(raw, str) and raw.startswith("#") else f"no FinMax EBITDA for {entity} {period}")
        return f"FinMax reconciliation unavailable ({why})."
    if ea is None:
        return "FinMax reconciliation unavailable (no FinPlan Actual EBITDA)."
    return f"FinPlan Actual vs FinMax EBITDA ({entity}): difference {_fmt(ea - efm)}."


def member_label(connection: str, name: str) -> str:
    """Display label from the atlas ('2TEC Technology (un.SY)' -> '2TEC Technology'); the name if unknown."""
    hits = [x for x in atlas.search(atlas.connect(), name, connection, None, 5) if x["name"] == name]
    return _clean(hits[0]["label"]) if hits else name


def _fmt(v: float | None) -> str:
    """Amount for running TEXT (key points, notes, chat): minus sign, never parentheses (tables use parentheses)."""
    if v is None:
        return "n/a"
    s = f"{abs(v):,.0f}".replace(",", " ")
    return f"−{s}" if v < -0.5 else s


def _pct(a, b) -> str:
    if a is None or not b:
        return "n/a"
    return f"{(a - b) / abs(b) * 100:+.1f} %".replace(".", ",").replace("-", "−")


def _now() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M")


FINPLAN_TOP = {"BANORWAY"}                    # FinPlan's Telenor Norge total = FinMax 'un.SY'


def map_finmax_entity(entity: str, explicit: str | None = None) -> tuple[str | None, str | None]:
    """FinMax entity for a FinPlan entity ('2TEC Technology' -> 'un.SY.2TEC', 'BANORWAY' -> 'un.SY'), checked
    against the atlas. Returns (member, None), or (None, note) when nothing matches. Without a usable FinMax atlas
    the 'un.SY.<code>' guess is kept (the error-cell scan still flags it if it is wrong)."""
    if explicit:
        return explicit, None
    code = entity.split(" ")[0]
    guess = f"un.SY.{code}"
    db = atlas.connect()
    known = atlas.usable(db, "FinMax", "All Entities")
    if known and atlas.has_member(db, "FinMax", "All Entities", guess):
        return guess, None
    if code in FINPLAN_TOP or entity in FINPLAN_TOP:
        return "un.SY", None
    if not known:
        return guess, None
    return None, f"No FinMax entity matches {entity}; pass finmax_entity to reconcile."


# ------------------------------------------------------------------ B. cost base per cost centre
def cost_base_members(connection: str, parent: str) -> tuple[list[str], set[str]]:
    """(descendants + the parent, in tree order, de-duplicated; the leaf set) from the atlas."""
    db = atlas.connect()
    ents = [e["name"] for e in atlas.descendants(db, connection, parent, limit=3000) if e["dimension"] == "All Entities"]
    leaves = {e["name"] for e in atlas.descendants(db, connection, parent, leaves_only=True, limit=3000)}
    return list(dict.fromkeys(ents + [parent])), leaves


# ------------------------------------------------------------------ C. monthly trend
def monthly_trend(h: ExcelHost, entity: str, accounts: list[str] | None = None, through: str | None = None,
                  connection: str = "FinMax") -> Plan:
    through = through or closed_month(h, "FinMax")
    months = months_to(through)
    accounts = accounts or ["ac.C30", "ac.C40S", "ac.C70", "ac.EBITDA"]
    pov = {"All Entities": entity, "All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"}
    pulls = [PullSpec("tr", connection, [Axis("All Accounts", accounts)], [Axis("All Time Periods", months)], pov),
             PullSpec("ytd", connection, [Axis("All Accounts", accounts)], [Axis("All Time Periods", [through])],
                      {**pov, "All Balance Types": "Accumulated"}, title_suffix="YTD")]
    ytd_col = len(months)

    def answer(P):
        p, y = P["tr"], P["ytd"]
        lines, recon, checks = [], [], []
        for i, a in enumerate(accounts):
            cells = [{"pull": p.sheet, "row": [a], "col": [m]} for m in months] + [{"rowsum": list(range(len(months)))}]
            lines.append({"key": f"a{i}", "label": _label_of(p, [a], a),
                          "kind": "grand" if "EBITDA" in a else "line", "cells": cells})
            recon.append({"key": f"y{i}", "label": f"{_label_of(p, [a], a)} — Accumulated {through}", "kind": "line",
                          "cells": [None] * ytd_col + [{"pull": y.sheet, "row": [a], "col": [through]}]})
            checks.append({"label": f"Sum of months = YTD: {_label_of(p, [a], a)}", "expr": f"a{i}-y{i}",
                           "columns": [ytd_col]})
        cols = [{"header": m, "kind": "heltall"} for m in months] + [{"header": f"YTD {through}", "kind": "heltall"}]
        last, prev = (_val(p, [accounts[-1]], [months[-1]]), _val(p, [accounts[-1]], [months[-2]])) if len(months) > 1 else (None, None)
        kp = [f"{_label_of(p, [accounts[-1]], accounts[-1])} {months[-1]}: {_fmt(last)} vs {_fmt(prev)} in "
              f"{months[-2] if len(months) > 1 else '-'} ({_pct(last, prev)})."]
        return {"title": f"{member_label(connection, entity)} — monthly trend {months[0]}–{through}",
                "subtitle": f"NOK · {connection} · Actual · {entity} · Periodic",
                "key_points": kp, "columns": cols,
                "sections": [{"title": "Per month", "lines": lines},
                             {"title": "Reconciliation to Accumulated (YTD) in FinMax", "lines": recon}],
                "checks": checks,
                "notes": ["Periodic = the month's own movement; the YTD column sums the months shown and is checked "
                          "against FinMax Accumulated."],
                "source": f"Source: Oracle EPM Cloud ({connection}) via Smart View, retrieved {_now()}.",
                "chart": {"type": "line", "row": f"a{len(accounts) - 1}", "columns": list(range(len(months))),
                          "title": f"{_label_of(p, [accounts[-1]], accounts[-1])} per month"}}

    return Plan(f"{member_label(connection, entity)} trend {through}", pulls, answer)


# ------------------------------------------------------------------ D. actual vs budget / forecast (FinPlan)
# Stops at EBITDA: Target (budget) carries no depreciation/write-downs/other items, so EBIT variances would mislead.
def fp_lines() -> list[tuple]:
    """key, FinPlan account, kind"""
    return [(x.key, x.account, x.kind) for x in catalog.load().pnl["FinPlan"]]


def fp_checks() -> list[tuple]:
    return [(c.label, c.expr) for c in catalog.load().checks["FinPlan"]]


def forecasts(year_yy: int) -> list[str]:
    """The FinPlan 'Forecast Q<q><yy>' members of a year in the atlas, newest quarter first."""
    with contextlib.closing(atlas.connect()) as db:
        names = [r[0] for r in db.execute("SELECT DISTINCT name FROM member WHERE connection='FinPlan' AND "
                                          "dimension='Scenario' AND name LIKE 'Forecast Q%'")]
    cands = []
    for n in names:
        m = re.fullmatch(r"Forecast Q(\d)(\d\d)", n)
        if m and int(m.group(2)) == year_yy:
            cands.append((int(m.group(1)), n))
    return [n for _, n in sorted(cands, reverse=True)]


def latest_forecast(year_yy: int) -> str | None:
    names = forecasts(year_yy)
    return names[0] if names else None


def actual_vs_plan(h: ExcelHost, entity: str, month: str | None = None, year: str | None = None, basis: str = "ytd",
                   forecast: str | None = None, finmax_entity: str | None = None, closed: str | None = None) -> Plan:
    closed = closed or closed_month(h, "FinMax")
    mo, yy = split_period(closed)
    month = month or MONTHS[mo - 1]
    year = year or f"FY{yy:02d}"
    fyy = int(year[2:])
    forecast = forecast or latest_forecast(fyy)
    bal = "Accumulated" if basis == "ytd" else "Periodic"
    accts = [a for _, a, _ in fp_lines()]
    base = {"Entity": entity, "Years": year, "Period": month, "All Balance Types": bal}
    pulls = [PullSpec("act", "FinPlan", [Axis("Account", accts)], [Axis("Scenario", ["Actual"])],
                      {**base, "Version": "Working"}, title_suffix="actual"),
             PullSpec("plan", "FinPlan", [Axis("Account", accts)],
                      [Axis("Scenario", ["Target"] + ([forecast] if forecast else []))], {**base, "Version": "Final"},
                      title_suffix="plan")]
    finmax_entity, recon_note = map_finmax_entity(entity, finmax_entity)
    fm_period = f"{month} {fyy:02d}"
    if finmax_entity:
        pulls.append(PullSpec("fm", "FinMax", [Axis("All Accounts", ["ac.EBITDA"])],
                              [Axis("All Time Periods", [fm_period])],
                              {"All Entities": finmax_entity, "All Scenarios": "Actual", "All Currencies": "NOK",
                               "All Balance Types": bal}, title_suffix="recon"))
    what = "YTD " if basis == "ytd" else ""

    def answer(P):
        a, pl, fm = P["act"], P["plan"], P.get("fm")
        lines = []
        for key, acct, kind in fp_lines():
            cells = [{"pull": a.sheet, "row": [acct], "col": ["Actual"]},
                     {"pull": pl.sheet, "row": [acct], "col": ["Target"]}, {"var": [0, 1]}, {"varpct": [0, 1]}]
            if forecast:
                cells += [{"pull": pl.sheet, "row": [acct], "col": [forecast]}, {"var": [0, 4]}]
            lines.append({"key": key, "label": acct, "kind": kind, "cells": cells})
        cols = [{"header": "Actual", "kind": "heltall"}, {"header": "Target (budget)", "kind": "heltall"},
                {"header": "Var. vs target", "kind": "heltall"}, {"header": "Var. %", "kind": "prosent1"}]
        if forecast:
            cols += [{"header": forecast, "kind": "heltall"}, {"header": "Var. vs forecast", "kind": "heltall"}]
        sections = [{"title": "Profit and loss (FinPlan definitions)", "lines": lines}]
        if fm is not None:
            sections.append({"title": "Reconciliation to the ledger (FinMax)", "lines": [
                {"key": "fm_ebitda", "label": f"EBITDA in FinMax ({finmax_entity})", "kind": "line",
                 "cells": [{"pull": fm.sheet, "row": ["ac.EBITDA"], "col": [fm_period]}]},
                {"key": "fm_diff", "label": "Difference FinPlan − FinMax", "kind": "line",
                 "cells": [{"expr": "ebitda-fm_ebitda"}]}]})
        ea, et = _val(a, ["EBITDA"], ["Actual"]), _val(pl, ["EBITDA"], ["Target"])
        ef = _val(pl, ["EBITDA"], [forecast]) if forecast else None
        kp = [f"EBITDA {what}{month} {year}: actual {_fmt(ea)} vs target {_fmt(et)} "
              f"(variance {_fmt(None if ea is None or et is None else ea - et)}, {_pct(ea, et)})."]
        if forecast:
            kp.append(f"Against {forecast}: {_fmt(ef)} (variance {_fmt(None if ea is None or ef is None else ea - ef)}).")
        kp.append(_recon_point(ea, fm, finmax_entity, fm_period, recon_note))
        notes = ["Actual and plan both come from FinPlan so line definitions match (FinPlan groups accounts "
                 "differently from FinMax; e.g. Technology revenue YTD Sep 26 is 59.8 MNOK in FinMax, 67.2 in FinPlan).",
                 "Variance = actual − plan; costs are negative, so a positive variance is better."]
        if recon_note:
            notes.append(recon_note)
        return {"title": f"{entity} — actual vs budget/forecast {what}{month} {year}",
                "subtitle": f"NOK · FinPlan · Actual (Working) vs Target and {forecast or 'no forecast'} (Final) · {bal}",
                "key_points": kp, "columns": cols, "sections": sections,
                "checks": [{"label": l, "expr": e, "columns": [0, 1]} for l, e in fp_checks()],
                "notes": notes,
                "source": f"Source: Oracle EPM Cloud (FinPlan, FinMax) via Smart View, retrieved {_now()}."}

    return Plan(f"{entity} actual vs plan {month} {year}", pulls, answer)
