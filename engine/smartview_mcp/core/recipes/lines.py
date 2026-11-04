"""Recipe lines_ranking (plan 2026-11-02 Batch B, D9/D17/D20): who the costs on a few accounts went to, from the
FinLite ledger lines. drills() names the cells to drill (d1 this year, d2 last year); lines_ranking() ranks the
invoiced amounts per supplier (top 10 + Other), shows the group companies as their own block and the lines without
a supplier (accruals, reversals, journals) as a reconciling line, and ties the total to the cube. A difference
between the lines and the cube is its own line ("Unexplained gap") and a value finding (no automatic re-drill)."""
from __future__ import annotations

import re
from dataclasses import dataclass

from ... import templates
from smartass.epm.grid import Axis
from ...host import InputError
from . import Draft, RecipeContext, pull_dict

TOP = 10
SUPPLIER = "AP_CUSTOMER_NAME"
CONN = "FinLite"
NO_SUPPLIER = "Accruals / reversals / journals without supplier"
GAP = "Unexplained gap (ledger lines vs cube)"
# commentary: only Total (cube) and the 3 largest supplier changes need a driver; subtotals/reconciliation never
TOP_DRIVERS = {"top": 3, "rest": "supplier changes"}
DEFAULTS = {"top": (lambda v: str(v).strip() == str(TOP), "the standard ranking is the top 10 + Other"),
            "compare": (lambda v: list(v or []) == ["ly"], "a ranking is compared with the same months last year"),
            "group_separate": (True, "group companies are internal costs, not suppliers to rank")}
FOLLOWUPS = ["Show the invoice lines for {top1}",
             "Monthly trend of {measure} this year vs last year",
             "Top suppliers on contractors (C60110)"]


@dataclass(frozen=True)
class Span:
    kind: str           # ytd | month | quarter
    first: int          # first month (1-12)
    last: int           # last month
    yy: int

    def head(self, yy: int) -> str:
        m = templates.MONTHS[self.last - 1]
        return {"ytd": f"YTD {m} {yy:02d}", "month": f"{m} {yy:02d}", "quarter": f"Q{self.last // 3} {yy:02d}"}[
            self.kind]

    def member(self, yy: int) -> str:
        return f"{templates.MONTHS[self.last - 1]} {yy:02d}"

    def opening(self, yy: int) -> str | None:
        """A quarter after Q1 = Accumulated at its end minus Accumulated before its first month."""
        return f"{templates.MONTHS[self.first - 2]} {yy:02d}" if self.kind == "quarter" and self.first > 1 else None

    def months(self, yy: int) -> list[str]:
        return [f"{templates.MONTHS[m - 1]}-{yy:02d}" for m in range(self.first, self.last + 1)]

    @property
    def balance(self) -> str:
        return "Periodic" if self.kind == "month" else "Accumulated"


def span(period: str | None, closed: str) -> Span:
    """The months a gate period label covers ('YTD Sep 26', 'Sep 26', 'Q3 26', 'FY 26', 'Yr 26', '2026', None).
    A full year not yet closed is YTD to the closed month."""
    cm, cy = templates.split_period(closed)
    p = (period or "").strip()
    if not p:
        return Span("ytd", 1, cm, cy)
    if p.upper().startswith("YTD "):
        mo, yy = templates.split_period(p[4:])
        return Span("ytd", 1, mo, yy)
    if m := re.fullmatch(r"Q([1-4]) (\d\d)", p, re.I):
        n = int(m.group(1))
        return Span("quarter", 3 * n - 2, 3 * n, int(m.group(2)))
    if m := re.fullmatch(r"(?:FY|Yr) ?(\d\d)|20(\d\d)", p, re.I):
        yy = int(m.group(1) or m.group(2))
        if yy > cy:
            raise InputError(f"{p} has no actuals yet (the last closed month is {closed}).")
        return Span("ytd", 1, 12 if yy < cy else cm, yy)
    mo, yy = templates.split_period(p)
    return Span("month", mo, mo, yy)


def drills(params: dict, ctx: RecipeContext) -> list[dict]:
    """The cells to drill: d1 = this year's ledger lines, d2 = last year's (unless compare drops 'ly'). One cell per
    account; the runner concatenates the rows of an alias."""
    sp = span(params.get("period"), ctx.closed)
    accounts, entity = _accounts(params), params.get("entity") or "un.SY"
    years = [("d1", sp.yy)] + ([("d2", sp.yy - 1)] if "ly" in (params.get("compare") or ["ly"]) else [])
    return [{"alias": a, "year": yy, "cells": [{"All Accounts": ac, "All Entities": entity,
                                                "All Time Periods": f"Yr {yy:02d}"} for ac in accounts]}
            for a, yy in years]


def is_group(name: str, rows: list[dict], gc) -> bool:
    """A group company: the counterparty column (if the catalog names one) says so, or the name matches a pattern."""
    if gc.column and any(str(r.get(gc.column) or "").strip() and "external" not in str(r.get(gc.column)).lower()
                         for r in rows):
        return True
    return any(re.search(p, name or "", re.I) for p in gc.patterns)


INITIATIVE, DERIVED = "INITIATIVE", "INITIATIVE_DERIVED"
NO_DESC = "(no description)"


def mark(params: dict, rows: list[dict], columns: list[str]) -> list[str]:
    """Before the Lines sheet is written: the initiative columns when the question asks for projects."""
    return add_initiative(rows, columns) if params.get("projects") else list(columns)


def add_initiative(rows: list[dict], columns: list[str] | None = None) -> list[str]:
    """Mark each line's initiative (in place): INITIATIVE = the booked project (PROJECT_DESC, else PROJECT) when
    PROJECT is filled and not 000000; otherwise INITIATIVE_DERIVED = the cost centre (CC_DESC), else the first 40
    characters of the line text without number-only words. Exactly one of the two is filled. Returns the columns
    with both added."""
    for r in rows:
        proj = str(r.get("PROJECT") or "").strip()
        if proj and proj.strip("0"):
            r[INITIATIVE], r[DERIVED] = str(r.get("PROJECT_DESC") or "").strip() or proj, ""
        else:
            text = " ".join(w for w in str(r.get("GL_DESCRIPTION_LINE") or "").split() if not w.isdigit())[:40]
            r[INITIATIVE], r[DERIVED] = "", str(r.get("CC_DESC") or "").strip() or text.strip() or NO_DESC
    cols = list(columns or [])
    return cols + [c for c in (INITIATIVE, DERIVED) if c not in cols]


def _by_project(rows: list[dict], now: list[str], head: str, top: int) -> dict:
    """The 'By project' topic sheet: booked projects and derived initiatives (top + Other each), tied to the lines."""
    def block(col: str, title: str, key: str, other: str, total: str) -> dict:
        sums: dict[str, list] = {}
        for r in rows:
            v = str(r.get(col) or "")
            if v:
                e = sums.setdefault(v.casefold(), [v, 0.0])
                e[1] += r["AMOUNT"] if isinstance(r.get("AMOUNT"), (int, float)) else 0.0
        ranked = sorted(sums.values(), key=lambda e: e[1])[:top]
        share = {"share": {"col": 0, "of": "plines"}}
        out = [{"key": f"{key}{i}", "label": n, "kind": "line",
                "cells": [{"drill": "d1", "where": {col: n, "Period": now}}, share]} for i, (n, _) in enumerate(ranked, 1)]
        rest = f"{key}" + "".join(f"-{key}{i}" for i in range(1, len(ranked) + 1))
        out += [{"key": f"{key}_other", "label": other, "kind": "line", "cells": [{"expr": rest}, share]},
                {"key": key, "label": total, "kind": "sum",
                 "cells": [{"drill": "d1", "where": {INITIATIVE if col == DERIVED else DERIVED: "", "Period": now}},
                           share]}]
        return {"title": title, "lines": out}
    tot = {"title": "Total", "lines": [{"key": "plines", "label": "Total (ledger lines)", "kind": "grand",
                                        "cells": [{"drill": "d1", "where": {"Period": now}}, {"value": 1}]}]}
    return {"topic": "By project", "layout": {
        "title": f"By project — {head}", "subtitle": "NOK · FinLite ledger lines · Actual",
        "columns": [{"header": head, "kind": "heltall"}, {"header": "Share %", "kind": "prosent1"}],
        "sections": [block(INITIATIVE, "Booked projects", "bp", "Other booked projects", "Total booked projects"),
                     block(DERIVED, "Initiative (derived, not booked)", "dp", "Other derived initiatives",
                           "Total derived (not booked)"), tot],
        "checks": [{"label": "By project = lines total", "expr": "bp+dp-plines", "columns": [0]}],
        "notes": ["Booked = the line carries a project (PROJECT not 000000). Derived = no project: the initiative is "
                  "the cost centre, else the start of the line text; it is an estimate, not a booking."]}}


def _accounts(params: dict) -> list[str]:
    acc = params.get("accounts") or []
    if not acc:
        raise InputError("lines_ranking needs 'accounts' (the FinLite accounts to rank, e.g. ['ac.C60080']).")
    return list(acc)


def _suppliers(rows: list[dict]) -> list[tuple[str, float, list[dict]]]:
    """(name as first seen, sum AMOUNT, rows) per non-blank supplier (case-insensitive), biggest cost first."""
    acc: dict[str, list] = {}
    for r in rows:
        name = str(r.get(SUPPLIER) or "").strip()
        if name:
            e = acc.setdefault(name.casefold(), [name, 0.0, []])
            e[1] += r["AMOUNT"] if isinstance(r.get("AMOUNT"), (int, float)) else 0.0
            e[2].append(r)
    return sorted((tuple(e) for e in acc.values()), key=lambda e: e[1])


def _title(accounts: list[str], head: str) -> str:
    """A short English title (the question has its own row): 'Consultancy (C60080 + C60090) — invoiced by
    supplier, YTD Oct 26'; the measure's label from capabilities.toml when it names exactly these accounts."""
    from .. import semantic
    m = next((x for x in semantic.all_measures() if x.label and set(x.accounts) == set(accounts)), None)
    codes = [a.removeprefix("ac.") for a in accounts]
    what = " + ".join(codes) if len(codes) <= 3 else f"{len(codes)} accounts"
    lead = f"{m.label} ({what})" if m else f"Accounts {what}"
    return f"{lead} — invoiced by supplier, {head}"


def lines_ranking(params: dict, ctx: RecipeContext) -> Draft:
    sp = span(params.get("period"), ctx.closed)
    accounts, entity = _accounts(params), params.get("entity") or "un.SY"
    top = int(params.get("top", TOP))
    yy, ly = sp.yy, sp.yy - 1
    now, before = sp.months(yy), sp.months(ly)
    head, head_ly = sp.head(yy), sp.head(ly)
    got = ctx.lines or {}
    d1 = [r for r in got.get("d1") or [] if r.get("Period") in now] if "d1" in got else None
    has_ly = "d2" in got and d1 is not None
    labels = {a: ctx.label(CONN, a) for a in accounts}
    measure = ", ".join(labels.values())
    var, pct = {"var": [0, 1]}, {"varpct": [0, 1]}

    def cells(where_now: dict, where_ly: dict, of: str = "cube") -> list:
        c0 = {"drill": "d1", "where": {**where_now, "Period": now}}
        if not has_ly:
            return [c0, None, None, None, {"share": {"col": 0, "of": of}}]
        return [c0, {"drill": "d2", "where": {**where_ly, "Period": before}}, var, pct, {"share": {"col": 0, "of": of}}]

    def calc(expr: str) -> list:
        return [{"expr": expr}, {"expr": expr} if has_ly else None, var if has_ly else None, pct if has_ly else None,
                {"share": {"col": 0, "of": "cube"}}]

    # the cube: per account (always when quarters need an opening balance or several accounts), then the total
    opening = sp.opening(yy)
    cols = [sp.member(yy), sp.member(ly)] + ([opening, sp.opening(ly)] if opening else [])
    pull = templates.PullSpec("cube", CONN, [Axis("All Accounts", accounts)], [Axis("All Time Periods", cols)],
                              {"All Entities": entity, "All Scenarios": "Actual", "All Currencies": "NOK",
                               "All Balance Types": sp.balance})

    def pulled(a: str, c: list[str]) -> list:
        return [{"pull": "cube", "row": [a], "col": [c[0]]}, {"pull": "cube", "row": [a], "col": [c[1]]}, var, pct,
                None]
    per_account = opening is not None or len(accounts) > 1 or d1 is None
    cube_lines = []
    if per_account:
        for i, a in enumerate(accounts, 1):
            cube_lines.append({"key": f"cube_a{i}", "label": labels[a], "aux": True, "cells": pulled(a, cols)})
            if opening:
                cube_lines.append({"key": f"open_a{i}", "label": f"{labels[a]} before {head}", "aux": True,
                                   "cells": pulled(a, cols[2:])})
        expr = "+".join(f"cube_a{i}" for i in range(1, len(accounts) + 1)) + \
            ("".join(f"-open_a{i}" for i in range(1, len(accounts) + 1)) if opening else "")
        cube_lines.append({"key": "cube", "label": "Total (cube)", "kind": "grand",
                           "cells": [{"expr": expr}, {"expr": expr}, var, pct, None]})
    else:
        cube_lines.append({"key": "cube", "label": "Total (cube)", "kind": "grand",
                           "cells": pulled(accounts[0], cols)})

    columns = [{"header": head, "kind": "heltall"}, {"header": head_ly, "kind": "heltall"},
               {"header": "Change", "kind": "heltall"}, {"header": "Change %", "kind": "prosent1"},
               {"header": "Share %", "kind": "prosent1"}]
    name = ctx.label(CONN, entity)
    lay = {"title": _title(accounts, head),
           "subtitle": f"NOK · {name} · {CONN} ledger lines and cube · Actual · {head} vs {head_ly}",
           "columns": columns, "sections": [],
           "notes": ["Budget per supplier: FinPlan has no supplier grain, so the plan is not compared here.",
                     "Share % = the line's amount / Total (cube) this year."],
           "source": f"Source: Oracle EPM Cloud ({CONN}) via Smart View, ledger lines drilled {ctx.now}."}
    assumptions = [f"Accounts {', '.join(accounts)}; entity {entity}; {head}",
                   "Ranking = invoiced amounts per supplier (AP lines); accruals, reversals and spreadsheet journals "
                   "have no supplier and form a reconciling line",
                   "Group companies shown separately", "NOK, EPM signs (costs negative)",
                   f"Last year = the same months of 20{ly:02d}"]
    warnings, ncols = [], [0, 1] if has_ly else [0]
    if d1 is None:
        lay["sections"] = [{"title": f"Cube total — {head}", "lines": cube_lines}]
        lay["checks"] = [{"label": "Accounts = total", "expr": expr + "-cube" if per_account else "cube-cube",
                          "columns": [0, 1]}]
        warnings.append("The ledger lines could not be drilled, so this shows the cube total only; ask again to "
                        "rank the suppliers.")
        return Draft([pull_dict(pull)], lay, assumptions, warnings, {"period": head, "fill": {"measure": measure}})
    if not has_ly:
        warnings.append("Last year's ledger lines could not be drilled: the last-year columns show the cube total "
                        "only.")

    ranked = _suppliers(d1)
    gc = ctx.cat.group_companies
    group = [s for s in ranked if is_group(s[0], s[2], gc)] if params.get("group_separate", True) else []
    gnames = {s[0] for s in group}
    external = [s for s in ranked if s[0] not in gnames][:top]
    sup = [{"key": f"s{i}", "label": n, "kind": "line", "driver": TOP_DRIVERS,
            "cells": cells({SUPPLIER: n}, {SUPPLIER: n})} for i, (n, _, _) in enumerate(external, 1)]
    others = "ext" + "".join(f"-s{i}" for i in range(1, len(external) + 1))
    sup += [{"key": "other", "label": "Other external suppliers", "kind": "line", "driver": "never",
             "cells": calc(others)},
            {"key": "ext", "label": "Total external suppliers", "kind": "sum", "driver": "never",
             "cells": calc("lines-grp-nosupp")}]
    grp = [{"key": f"grp{i}", "label": n, "kind": "line", "driver": TOP_DRIVERS,
            "cells": cells({SUPPLIER: n}, {SUPPLIER: n})} for i, (n, _, _) in enumerate(group, 1)]
    gkeys = [g["key"] for g in grp]
    gtot = [{"sum": gkeys}, {"sum": gkeys} if has_ly else None, var if has_ly else None, pct if has_ly else None,
            {"share": {"col": 0, "of": "cube"}}] if gkeys else [{"value": 0}, {"value": 0} if has_ly else None]
    grp.append({"key": "grp", "label": "Total group companies", "kind": "sum", "driver": "never", "cells": gtot})
    recon = [{"key": "nosupp", "label": NO_SUPPLIER, "kind": "line", "driver": "never",
              "cells": cells({SUPPLIER: ""}, {SUPPLIER: ""})},
             {"key": "lines", "label": "Total (ledger lines)", "kind": "sum", "driver": "never", "cells": cells({}, {})},
             {"key": "gap", "label": GAP, "kind": "line", "driver": "never", "cells": calc("cube-lines")}] + cube_lines
    lay["sections"] = [{"title": f"Invoiced by supplier — {head}", "lines": sup},
                       {"title": "Group companies", "lines": grp},
                       {"title": "Reconciliation", "lines": recon}]
    lay["checks"] = [{"label": "Blocks = lines total", "expr": "ext+grp+nosupp-lines", "columns": ncols},
                     {"label": "Total = cube", "expr": "lines+gap-cube", "columns": ncols}]
    finding = {"severity": "Warning", "line": "gap", "col": 0, "above": 1.0,
               "text": "Ledger lines differ from the cube by {value:,.0f} NOK (" + str(len(d1)) +
                       " lines); shown as 'Unexplained gap'. Ask again to re-drill if needed."}
    result = {"period": head, "fill": {"top1": external[0][0] if external else "the largest supplier",
                                       "measure": measure},
              "findings": [finding], "lines": {k: len(v or []) for k, v in got.items()}}
    extras = []
    if params.get("projects"):
        add_initiative(d1)                   # idempotent; the runner adds it before writing the Lines sheet
        extras.append(_by_project(d1, now, head, top))
    return Draft([pull_dict(pull)], lay, assumptions, warnings, result, extras)
