"""Analysis templates: plan the pulls, then turn the pulled data into an answer spec (see answer.py).

Rules agreed with the user (2026-10-13): actuals from FinMax; plan/forecast from FinPlan; actual-vs-plan entirely in
FinPlan (Actual = version Working, Target/forecast = version Final) with a FinMax EBITDA reconciliation line;
P&L questions compare with the same period last year by default; FinMax signs kept (costs negative), so
variance = current - comparison is positive when better.
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from typing import Callable

from . import atlas, ops
from .grid import Axis
from .host import ExcelHost, HostError

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
    answer: Callable[[dict], dict]            # {pull id: userexcel.Pull} -> answer spec
    followups: list[Callable] = field(default_factory=list)
    summary: Callable[[dict], dict] | None = None


# ------------------------------------------------------------------ periods
def closed_month(h: ExcelHost, connection: str = "FinMax") -> str:
    """'Current Reporting Month' -> e.g. 'Sep 26' (FinMax member ChildMbrName '[CurrRepMonth].[Y26M09]').
    Cached per day on disk: the live lookup needs the hidden server Excel (~10 s + login on a fresh server)."""
    import json
    from . import config
    cache = config.SVKIT / "closed_month.json"
    today = dt.date.today().isoformat()
    try:
        c = json.loads(cache.read_text(encoding="utf-8"))
        if c.get("date") == today and c.get(connection):
            return c[connection]
    except (OSError, ValueError):
        c = {}
    info = h.on_host(lambda: ops.find_member(h, connection, "Current Reporting Month")).get("info") or {}
    m = re.search(r"Y(\d\d)M(\d\d)", str(info.get("ChildMbrName", "")))
    if not m:
        raise HostError("Could not resolve 'Current Reporting Month' in FinMax; pass the period explicitly (e.g. 'Sep 26').")
    val = f"{MONTHS[int(m.group(2)) - 1]} {m.group(1)}"
    try:
        c = c if c.get("date") == today else {"date": today}
        c[connection] = val
        cache.write_text(json.dumps(c), encoding="utf-8")
    except OSError:
        pass
    return val


def split_period(p: str) -> tuple[int, int]:
    m = re.fullmatch(r"([A-Z][a-z]{2}) (\d\d)", p.strip())
    if not m or m.group(1) not in MONTHS:
        raise HostError(f"Period '{p}' is not a month like 'Sep 26'.")
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


def _val(p, row_members, col_members):
    """Parsed value from a pull (for ranking / key points; the sheet itself uses live formulas)."""
    try:
        j = p.layout.col_tuples.index(tuple(col_members))
    except ValueError:
        return None
    for r in p.parsed.rows:
        if tuple(r[:p.layout.nr]) == tuple(row_members):
            v = r[p.layout.nr + j]
            return v if isinstance(v, (int, float)) else None
    return None


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


def _dec(v: float | None) -> str:
    return "n/a" if v is None else f"{v:,.1f}".replace(",", " ").replace(".", ",")


def _now() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M")


# ------------------------------------------------------------------ A. division P&L
PNL = [  # key, FinMax account, kind, num
    ("rev", "ac.C30", "line", None), ("cogs", "ac.C40S", "line", None), ("gm", "ac.C48", "sum", None),
    ("pers", "ac.C50S", "line", None), ("oth", "ac.C60", "line", None), ("opex", "ac.C70", "sum", None),
    ("boi", "ac.EBITDA_BOI", "sum", None), ("c73", "ac.C73", "line", None), ("ebitda", "ac.EBITDA", "grand", None),
    ("da", "ac.C75", "line", None), ("fte", "ac.STA20Y", "line", "tall1"),
]
PNL_CHECKS = [("C30 + C40S = C48", "rev+cogs-gm"), ("C50S + C60 = C70", "pers+oth-opex"),
              ("C48 + C70 = EBITDA_BOI", "gm+opex-boi"), ("EBITDA_BOI + C73 = EBITDA", "boi+c73-ebitda")]


def division_pnl(h: ExcelHost, entity: str, period: str | None = None, basis: str = "ytd", compare: str = "ly",
                 drivers: bool = True, connection: str = "FinMax") -> Plan:
    closed = closed_month(h, "FinMax")
    period = period or closed
    cmp_p = same_period_ly(period) if compare == "ly" else None
    bal = "Accumulated" if basis == "ytd" else "Periodic"
    cols = [period] + ([cmp_p] if cmp_p else [])
    pov = {"All Entities": entity, "All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": bal}
    pulls = [PullSpec("pnl", connection, [Axis("All Accounts", [a for _, a, _, _ in PNL])],
                      [Axis("All Time Periods", cols)], pov)]
    kids = []
    if drivers and cmp_p:
        kids = [k["name"] for k in atlas.children(atlas.connect(), connection, entity)][:40]
        if kids:
            pulls.append(PullSpec("drv", connection, [Axis("All Entities", kids)], [Axis("All Time Periods", cols)],
                                  {"All Accounts": "ac.EBITDA", "All Scenarios": "Actual", "All Currencies": "NOK",
                                   "All Balance Types": bal}, title_suffix="drivers"))
    mo, yy = split_period(period)
    closed_ok = (yy, mo) <= tuple(reversed(split_period(closed)))
    what = "YTD " if basis == "ytd" else ""

    ent_label = member_label(connection, entity)

    def answer(P):
        p = P["pnl"]
        lines = []
        for key, acct, kind, num in PNL:
            cells = [{"pull": p.sheet, "row": [acct], "col": [period]}]
            if cmp_p:
                cells += [{"pull": p.sheet, "row": [acct], "col": [cmp_p]}, {"var": [0, 1]}, {"varpct": [0, 1]}]
            ln = {"key": key, "label": _label_of(p, [acct], acct), "kind": kind, "cells": cells}
            if num:
                ln["num"] = num
            lines.append(ln)
        columns = [{"header": f"{what}{period}", "kind": "heltall"}]
        if cmp_p:
            columns += [{"header": f"{what}{cmp_p}", "kind": "heltall"}, {"header": "Variance", "kind": "heltall"},
                        {"header": "Variance %", "kind": "prosent1"}]
        sections = [{"title": "Profit and loss", "lines": lines}]
        e_cur, e_cmp = _val(p, ["ac.EBITDA"], [period]), _val(p, ["ac.EBITDA"], [cmp_p]) if cmp_p else None
        r_cur, r_cmp = _val(p, ["ac.C30"], [period]), _val(p, ["ac.C30"], [cmp_p]) if cmp_p else None
        o_cur, o_cmp = _val(p, ["ac.C70"], [period]), _val(p, ["ac.C70"], [cmp_p]) if cmp_p else None
        f_cur, f_cmp = _val(p, ["ac.STA20Y"], [period]), _val(p, ["ac.STA20Y"], [cmp_p]) if cmp_p else None
        kp = [f"EBITDA {what}{period}: {_fmt(e_cur)}" + (f" vs {_fmt(e_cmp)} in {cmp_p} "
                                                       f"(variance {_fmt(None if e_cur is None or e_cmp is None else e_cur - e_cmp)}, {_pct(e_cur, e_cmp)})." if cmp_p else ".")]
        if cmp_p:
            kp.append(f"Revenue {_fmt(r_cur)} ({_pct(r_cur, r_cmp)} vs {cmp_p}); total opex {_fmt(o_cur)} "
                      f"({_pct(o_cur, o_cmp)}; costs are negative, so a negative variance means higher costs).")
            if f_cur is not None:
                kp.append(f"FTE at end of {period}: {_dec(f_cur)} vs {_dec(f_cmp)} in {cmp_p}.")
        chart = None
        if "drv" in P:
            d = P["drv"]
            ranked = []
            for k in kids:
                a, b = _val(d, [k], [period]), _val(d, [k], [cmp_p])
                if a is not None or b is not None:
                    ranked.append((abs((a or 0) - (b or 0)), k, (a or 0) - (b or 0)))
            ranked.sort(reverse=True)
            top = ranked[:5]
            dl = []
            for n, (_, k, var) in enumerate(top, 1):
                dl.append({"key": f"d{n}", "label": _label_of(d, [k], k), "kind": "line",
                           "cells": [{"pull": d.sheet, "row": [k], "col": [period]},
                                     {"pull": d.sheet, "row": [k], "col": [cmp_p]}, {"var": [0, 1]}, {"varpct": [0, 1]}]})
            if top:
                rest = "ebitda-" + "-".join(f"d{n}" for n in range(1, len(top) + 1))
                dl.append({"key": "drest", "label": "Other units", "kind": "line",
                           "cells": [{"expr": rest}, {"expr": rest}, {"var": [0, 1]}, {"varpct": [0, 1]}]})
                sections.append({"title": f"Largest EBITDA variance drivers vs {cmp_p} (sub-units)", "lines": dl})
                bridge = [{"key": "b0", "label": f"EBITDA {what}{cmp_p}", "kind": "line",
                           "cells": [{"ref": {"key": "ebitda", "col": 1}}]}]
                for n in range(1, len(top) + 1):
                    bridge.append({"key": f"b{n}", "label": dl[n - 1]["label"], "kind": "line",
                                   "cells": [{"ref": {"key": f"d{n}", "col": 2}}]})
                bridge.append({"key": "brest", "label": "Other units", "kind": "line",
                               "cells": [{"ref": {"key": "drest", "col": 2}}]})
                bridge.append({"key": "bend", "label": f"EBITDA {what}{period}", "kind": "sum",
                               "cells": [{"ref": {"key": "ebitda", "col": 0}}]})
                sections.append({"title": f"EBITDA bridge {cmp_p} → {period}", "lines": bridge})
                chart = {"type": "waterfall", "lines": [b["key"] for b in bridge], "column": 0,
                         "title": f"EBITDA bridge {cmp_p} → {period}"}
                kp.append(f"Largest driver: {dl[0]['label']}, variance {_fmt(top[0][2])} vs {cmp_p}.")
        checks = [{"label": lab, "expr": ex, "columns": [0, 1] if cmp_p else [0]} for lab, ex in PNL_CHECKS]
        checks.append({"label": "Period is closed", "static": "OK" if closed_ok else "DIFF",
                       "detail": f"{period} vs Current Reporting Month {closed}"})
        notes = ["Signs as in FinMax: costs are negative; variance = current − comparison (positive = better).",
                 "History is restated to today's organisation (e.g. the 1 June 2026 reorg); compare old extracts on "
                 "leaf members or the company total."]
        if not closed_ok:
            notes.insert(0, f"{period} is after the Current Reporting Month ({closed}): the period is still open.")
        return {"title": f"{ent_label} — P&L {what}{period}" + (f" vs {cmp_p}" if cmp_p else ""),
                "subtitle": f"NOK · {connection} · Actual · {entity} · {bal}",
                "key_points": kp, "columns": columns, "sections": sections, "checks": checks, "notes": notes,
                "source": f"Source: Oracle EPM Cloud ({connection}) via Smart View, retrieved {_now()}.",
                "chart": chart}

    return Plan(f"{ent_label} P&L {what}{period}", pulls, answer)


# ------------------------------------------------------------------ B. cost base per cost centre
def cost_base(h: ExcelHost, parent: str = "un.SY", accounts: list[str] | None = None, period: str | None = None,
              top: int = 25, connection: str = "FinMax") -> Plan:
    """All descendants are pulled (the data sheet shows the full tree), but only LEAF cost centres are ranked,
    so parents are not double counted in the top list."""
    period = period or closed_month(h, "FinMax")
    accounts = accounts or ["ac.C50S", "ac.C60", "ac.C70", "ac.STA20Y"]
    db = atlas.connect()
    ents = [e["name"] for e in atlas.descendants(db, connection, parent, limit=3000) if e["dimension"] == "All Entities"]
    leaves = {e["name"] for e in atlas.descendants(db, connection, parent, leaves_only=True, limit=3000)}
    seen, members = set(), []
    for e in ents + [parent]:
        if e not in seen:
            seen.add(e)
            members.append(e)
    if len(members) * len(accounts) > 50000:
        raise HostError(f"{len(members)} cost centres × {len(accounts)} accounts is too large; narrow 'parent'.")
    pov = {"All Scenarios": "Actual", "All Currencies": "NOK", "All Time Periods": period,
           "All Balance Types": "Accumulated"}
    pulls = [PullSpec("cb", connection, [Axis("All Entities", members)], [Axis("All Accounts", accounts)], pov)]
    main = "ac.C70" if "ac.C70" in accounts else accounts[0]

    def answer(P):
        p = P["cb"]
        ranked = sorted(((abs(_val(p, [m], [main]) or 0), m) for m in members if m in leaves), reverse=True)
        lines = []
        for n, (_, m) in enumerate(ranked[:top], 1):
            lines.append({"key": f"c{n}", "label": _label_of(p, [m], m), "kind": "line",
                          "cells": [{"pull": p.sheet, "row": [m], "col": [a]} for a in accounts]})
        if ranked[top:]:
            rest = "tot-" + "-".join(f"c{n}" for n in range(1, min(top, len(ranked)) + 1))
            lines.append({"key": "rest", "label": f"Other {len(ranked) - top} cost centres", "kind": "line",
                          "cells": [{"expr": rest} for _ in accounts]})
        lines.append({"key": "tot", "label": f"Total {_label_of(p, [parent], parent)}", "kind": "grand",
                      "cells": [{"pull": p.sheet, "row": [parent], "col": [a]} for a in accounts]})
        cols = [{"header": _clean(p.col_labels.get((a,), (a,))[-1]), "kind": "tall1" if "STA" in a else "heltall"}
                for a in accounts]
        tot = _val(p, [parent], [main])
        topsum = sum(_val(p, [m], [main]) or 0 for _, m in ranked[:top])
        kp = [f"{len(ranked)} cost centres (leaf level) under {member_label(connection, parent)}; total "
              f"{cols[accounts.index(main)]['header']} {_fmt(tot)} YTD {period}.",
              f"The {min(top, len(ranked))} largest account for {_fmt(topsum)} "
              f"({(topsum / tot * 100 if tot else 0):.0f} % of the total). The full tree is on the data sheet '{p.sheet}'."]
        return {"title": f"Cost base per cost centre — {member_label(connection, parent)} — YTD {period}",
                "subtitle": f"NOK · {connection} · Actual · Accumulated {period}",
                "key_points": kp, "columns": cols,
                "sections": [{"title": f"Top {min(top, len(ranked))} cost centres by {cols[accounts.index(main)]['header']}",
                              "lines": lines}],
                "notes": ["FTE (STA20Y) is the end-of-period headcount.",
                          "Parent nodes are restated to today's organisation."],
                "source": f"Source: Oracle EPM Cloud ({connection}) via Smart View, retrieved {_now()}."}

    return Plan(f"{member_label(connection, parent)} cost base {period}", pulls, answer)


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
FP_PNL = [("rev", "Revenue", "line"), ("cogs", "COGS", "line"), ("gp", "Gross_Profit", "sum"),
          ("opex", "OPEX", "line"), ("boi", "EBITDA before Other Items", "sum"), ("oth", "Total Other Items", "line"),
          ("ebitda", "EBITDA", "grand")]
FP_CHECKS = [("Revenue + COGS = Gross profit", "rev+cogs-gp"), ("Gross profit + OPEX = EBITDA before other items", "gp+opex-boi"),
             ("EBITDA before other items + Other items = EBITDA", "boi+oth-ebitda")]


def latest_forecast(year_yy: int) -> str | None:
    db = atlas.connect()
    names = [r[0] for r in db.execute("SELECT DISTINCT name FROM member WHERE connection='FinPlan' AND dimension='Scenario' "
                                      "AND name LIKE 'Forecast Q%'")]
    cands = []
    for n in names:
        m = re.fullmatch(r"Forecast Q(\d)(\d\d)", n)
        if m and int(m.group(2)) == year_yy:
            cands.append((int(m.group(1)), n))
    return max(cands)[1] if cands else None


def actual_vs_plan(h: ExcelHost, entity: str, month: str | None = None, year: str | None = None, basis: str = "ytd",
                   forecast: str | None = None, finmax_entity: str | None = None) -> Plan:
    closed = closed_month(h, "FinMax")
    mo, yy = split_period(closed)
    month = month or MONTHS[mo - 1]
    year = year or f"FY{yy:02d}"
    fyy = int(year[2:])
    forecast = forecast or latest_forecast(fyy)
    bal = "Accumulated" if basis == "ytd" else "Periodic"
    accts = [a for _, a, _ in FP_PNL]
    base = {"Entity": entity, "Years": year, "Period": month, "All Balance Types": bal}
    pulls = [PullSpec("act", "FinPlan", [Axis("Account", accts)], [Axis("Scenario", ["Actual"])],
                      {**base, "Version": "Working"}, title_suffix="actual"),
             PullSpec("plan", "FinPlan", [Axis("Account", accts)],
                      [Axis("Scenario", ["Target"] + ([forecast] if forecast else []))], {**base, "Version": "Final"},
                      title_suffix="plan")]
    code = entity.split(" ")[0]
    finmax_entity = finmax_entity or f"un.SY.{code}"
    fm_period = f"{month} {fyy:02d}"
    pulls.append(PullSpec("fm", "FinMax", [Axis("All Accounts", ["ac.EBITDA"])], [Axis("All Time Periods", [fm_period])],
                          {"All Entities": finmax_entity, "All Scenarios": "Actual", "All Currencies": "NOK",
                           "All Balance Types": bal}, title_suffix="recon"))
    what = "YTD " if basis == "ytd" else ""

    def answer(P):
        a, pl, fm = P["act"], P["plan"], P["fm"]
        lines = []
        for key, acct, kind in FP_PNL:
            cells = [{"pull": a.sheet, "row": [acct], "col": ["Actual"]},
                     {"pull": pl.sheet, "row": [acct], "col": ["Target"]}, {"var": [0, 1]}, {"varpct": [0, 1]}]
            if forecast:
                cells += [{"pull": pl.sheet, "row": [acct], "col": [forecast]}, {"var": [0, 4]}]
            lines.append({"key": key, "label": acct, "kind": kind, "cells": cells})
        cols = [{"header": "Actual", "kind": "heltall"}, {"header": "Target (budget)", "kind": "heltall"},
                {"header": "Var. vs target", "kind": "heltall"}, {"header": "Var. %", "kind": "prosent1"}]
        if forecast:
            cols += [{"header": forecast, "kind": "heltall"}, {"header": "Var. vs forecast", "kind": "heltall"}]
        recon = [{"key": "fm_ebitda", "label": f"EBITDA in FinMax ({finmax_entity})", "kind": "line",
                  "cells": [{"pull": fm.sheet, "row": ["ac.EBITDA"], "col": [fm_period]}]},
                 {"key": "fm_diff", "label": "Difference FinPlan − FinMax", "kind": "line",
                  "cells": [{"expr": "ebitda-fm_ebitda"}]}]
        ea, et = _val(a, ["EBITDA"], ["Actual"]), _val(pl, ["EBITDA"], ["Target"])
        ef = _val(pl, ["EBITDA"], [forecast]) if forecast else None
        efm = _val(fm, ["ac.EBITDA"], [fm_period])
        kp = [f"EBITDA {what}{month} {year}: actual {_fmt(ea)} vs target {_fmt(et)} "
              f"(variance {_fmt(None if ea is None or et is None else ea - et)}, {_pct(ea, et)})."]
        if forecast:
            kp.append(f"Against {forecast}: {_fmt(ef)} (variance {_fmt(None if ea is None or ef is None else ea - ef)}).")
        kp.append(f"FinPlan Actual reconciles to FinMax EBITDA within {_fmt(None if ea is None or efm is None else ea - efm)}.")
        return {"title": f"{entity} — actual vs budget/forecast {what}{month} {year}",
                "subtitle": f"NOK · FinPlan · Actual (Working) vs Target and {forecast or 'no forecast'} (Final) · {bal}",
                "key_points": kp, "columns": cols,
                "sections": [{"title": "Profit and loss (FinPlan definitions)", "lines": lines},
                             {"title": "Reconciliation to the ledger (FinMax)", "lines": recon}],
                "checks": [{"label": l, "expr": e, "columns": [0, 1]} for l, e in FP_CHECKS],
                "notes": ["Actual and plan both come from FinPlan so line definitions match (FinPlan groups accounts "
                          "differently from FinMax; e.g. Technology revenue YTD Sep 26 is 59.8 MNOK in FinMax, 67.2 in FinPlan).",
                          "Variance = actual − plan; costs are negative, so a positive variance is better."],
                "source": f"Source: Oracle EPM Cloud (FinPlan, FinMax) via Smart View, retrieved {_now()}."}

    return Plan(f"{entity} actual vs plan {month} {year}", pulls, answer)


TEMPLATES = {"division_pnl": division_pnl, "cost_base": cost_base, "monthly_trend": monthly_trend,
             "actual_vs_plan": actual_vs_plan}
