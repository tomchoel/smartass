"""Golden report (live): sv_drill (FinLite un.SY.2WHS x ac.C40S x Q3 26) + sv_ws_answer with {drill: ...} cells
reproduces the approved reference example (spikes\\drill\\build_example.py): 1 reconciliation, 2 JE source, 3 journal
category, 4 month (FILTER form), 5 account, 6 top 12 suppliers + no supplier + other, 7 top 10 products + other,
8 ten largest invoices. Member lists come from the drill's own lines (drill.db), not hard-coded.
Leaves the new analysis workbook open in the user's Excel. Opens NO browser window.
Usage: PY tests/live_drill_report.py"""
import asyncio
import collections
import datetime as dt
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mcp import Client  # noqa: E402

from smartview_mcp import drill_store, server  # noqa: E402

CELL = {"All Accounts": "ac.C40S", "All Entities": "un.SY.2WHS", "All Time Periods": "Q3 26"}
EXPECT = {"lines": 1231, "sum": -24706615.55, "debit": -369759826.95, "credit": 345053211.40,
          "Payables": (-123132549.19, 359), "months": [354, 459, 418], "SUPPLIER 037 AS": -60646654.33,
          "SUPPLIER 020 AS": -20187679.96, "no_supplier": 872}
SOURCE_DESC = {"Payables": "purchase invoices (Coupa/AP)", "Spreadsheet": "manual journals and accruals",
               "DATABRONN": "Databrønn traffic feeds (roaming, interconnect)", "Projects": "project costs",
               "Cost Accounting": "cost accounting transactions"}


async def _call(c, name: str, args: dict) -> dict:
    r = await c.call_tool(name, args)
    if r.is_error:
        raise AssertionError(f"{name} failed: {r.content}")
    body = r.structured_content or json.loads(r.content[0].text)   # a plain -> dict tool has no output schema
    return body.get("result", body)


def nok(v: float) -> str:
    """Full NOK in text: thousands with a space, minus sign '−'."""
    s = f"{abs(v):,.0f}".replace(",", " ")
    return ("−" if v < -0.5 else "+" if v > 0.5 else "") + s


def groups(rows, col):
    """value -> [amount, lines], ordered like the reference: most negative (largest cost) first."""
    d = collections.defaultdict(lambda: [0.0, 0])
    for x in rows:
        e = d[x.get(col) or ""]
        e[0] += x["AMOUNT"]
        e[1] += 1
    return dict(sorted(d.items(), key=lambda kv: kv[1][0]))


def first(rows, col, val, want):
    return next((x.get(want) for x in rows if x.get(col) == val and x.get(want)), "")


def dl(key, label, where):
    return {"key": key, "label": label, "cells": [{"drill": "d", "where": where},
                                                  {"drill": "d", "where": where, "agg": "count"}]}


def block(prefix, title, col, keys, label=lambda k: k, none_label=None, other_label=None, total=True):
    lines = [dl(f"{prefix}{i}", label(k), {col: k}) for i, k in enumerate(keys)]
    parts = [f"{prefix}*"]
    if none_label or other_label:            # top-N subtotal, then (none) and the remainder
        lines.append({"key": f"top{prefix}", "label": f"Top {len(keys)}", "kind": "sum",
                      "cells": [{"sum": f"{prefix}*"}, {"sum": f"{prefix}*"}]})
        parts = [f"top{prefix}"]
    if none_label:
        lines.append(dl(f"none{prefix}", none_label, {col: ""}))
        parts.append(f"none{prefix}")
    if other_label:
        rest = f"sum-top{prefix}" + (f"-none{prefix}" if none_label else "")
        lines.append({"key": f"oth{prefix}", "label": other_label, "cells": [{"expr": rest}, {"expr": rest}]})
        parts.append(f"oth{prefix}")
    if total:
        cell = {"sum": parts[0]} if len(parts) == 1 else {"sum": parts}
        lines.append({"key": f"tot{prefix}", "label": "Total", "kind": "sum", "cells": [cell, dict(cell)]})
    return {"title": title, "lines": lines}


def month_order(keys):
    return sorted(keys, key=lambda k: dt.datetime.strptime(k, "%b-%y"))


def build_spec(d: dict, rows: list[dict]) -> tuple[dict, dict]:
    pull, rec = d["cube_pull"], d["reconciliation"]
    src, cat, acc = groups(rows, "JE_SOURCE"), groups(rows, "JE_CATEGORY_NAME"), groups(rows, "ACCOUNT")
    months = month_order([k for k in groups(rows, "Period") if k])
    sup = [k for k in groups(rows, "AP_CUSTOMER_NAME") if k][:12]
    prd = [k for k in groups(rows, "Products") if k][:10]
    inv = [k for k in groups(rows, "AP_INVOICE_NUMBER") if k][:10]
    acc_name = {k: first(rows, "ACCOUNT", k, "ACCOUNT_NAME") for k in acc}
    sections = [
        {"title": "1. Reconciliation to the cube", "lines": [
            {"key": "cube", "label": "Cube value, FinLite (live Smart View grid)",
             "cells": [{"pull": pull["sheet"], "row": pull["row_keys"][0], "col": pull["col_keys"][0]}, None]},
            dl("sum", "Sum of drill-through lines", {}),
            {"key": "gap", "label": "Gap (cube − lines)", "kind": "grand", "cells": [{"expr": "cube-sum"}, None]},
            {"key": "debit", "label": "Gross debit (cost lines, negative)",
             "cells": [{"drill": "d", "where": {}, "agg": "debit"}, None]},
            {"key": "credit", "label": "Gross credit (allocations out, reversals, positive)",
             "cells": [{"drill": "d", "where": {}, "agg": "credit"}, None]}]},
        block("src", "2. By posting source (JE_SOURCE)", "JE_SOURCE", list(src),
              lambda k: f"{k} – {SOURCE_DESC[k]}" if k in SOURCE_DESC else k),
        block("cat", "3. By journal category (JE_CATEGORY_NAME)", "JE_CATEGORY_NAME", list(cat)),
        block("mon", "4. By month", "Period", months),
        block("acc", "5. By account", "ACCOUNT", list(acc), lambda k: f"{k} {acc_name[k]}".strip()),
        block("sup", "6. Top 12 suppliers (AP_CUSTOMER_NAME)", "AP_CUSTOMER_NAME", sup,
              none_label="(no supplier: accruals, feeds, journals)", other_label="Other suppliers"),
        block("prd", "7. Top 10 products (pd.SY tag)", "Products", prd,
              lambda k: f"{k} {first(rows, 'Products', k, 'PRODUCT_DESC')}".strip(), other_label="Other products"),
        block("inv", "8. Ten largest invoices", "AP_INVOICE_NUMBER", inv,
              lambda k: f"{k} – {first(rows, 'AP_INVOICE_NUMBER', k, 'AP_CUSTOMER_NAME')}", total=False),
    ]
    sec_checks = [(f"tot{p}", t) for p, t in (("src", "Posting sources"), ("cat", "Journal categories"),
                                             ("mon", "Months"), ("acc", "Accounts"), ("sup", "Suppliers"),
                                             ("prd", "Products"))]
    checks = [{"label": "Gap is zero (lines reconcile to the cube)", "expr": "gap", "columns": [0]},
              {"label": "Debit + credit = sum of lines", "expr": "debit+credit-sum", "columns": [0]}] + \
             [{"label": f"{t} total = sum of lines", "expr": f"{k}-sum", "columns": [0, 1]} for k, t in sec_checks]

    # key points: short and interpretive, numbers taken from the lines
    pay = src.get("Payables", [0.0, 0])
    man = [x for x in rows if x.get("JE_CATEGORY_NAME") == "Manual"]
    jn = groups(man, "JOURNAL_NAME")
    ic = sum(v[0] for k, v in jn.items() if "allokering IC VK" in k)
    alloc = [(k.replace(" Manual SY", ""), v[0]) for k, v in jn.items() if "allokering" in k and "IC VK" not in k]
    sup_g = groups(rows, "AP_CUSTOMER_NAME")
    ic_acc = max(groups(man, "ACCOUNT").items(), key=lambda kv: kv[1][0])[0]
    top_acc = next(iter(acc))
    n = lambda i: f"{i:,}".replace(",", " ")
    short = lambda name: " ".join(name.split()[:2]) if name.lower().startswith("telenor") else name.split()[0].title()
    acc2 = list(acc)[1]
    pts = [
        f"Q3 COGS for Wholesale is {nok(rec['lines_sum'])} on {n(rec['lines'])} ledger lines and reconciles "
        "exactly to the FinLite cell.",
        f"Purchase invoices are {nok(pay[0])} gross ({pay[1]} lines), mostly interconnect from "
        + ", ".join(f"{short(s)} ({nok(sup_g[s][0])})" for s in sup[:2]) + f" and {short(sup[2])} "
        f"({nok(sup_g[sup[2]][0])}).",
        f"Manual journals credit {nok(cat['Manual'][0])} on {ic_acc}: the interconnect allocation "
        f"'SF26xx allokering IC VK' ({nok(ic)})" + "".join(f" and '{k}' ({nok(v)})" for k, v in alloc)
        + f" pass the traffic cost on, so {ic_acc} nets to {nok(acc[ic_acc][0])}.",
        f"What stays is mainly energy for production ({top_acc}, Telenor Towers, {nok(acc[top_acc][0])}) and "
        f"accruals on {acc2} {acc_name[acc2]} ({nok(acc[acc2][0])}). {n(sup_g[''][1])} of {n(rec['lines'])} "
        "lines have no supplier (accruals, Databrønn feeds, journals).",
    ]
    spec = {"title": "Wholesale COGS Q3 2026 – what is behind the number",
            "subtitle": "FinLite · un.SY.2WHS Wholesale · C40S Total COGS · Q3 26 (Jul–Sep) · Actual · NOK · Periodic "
                        "· ledger lines from drill-through",
            "auto_points": False, "key_points": pts,
            "columns": [{"header": "NOK"}, {"header": "Lines"}],
            "sections": sections, "checks": checks,
            "notes": ["Signs as in FinMax: costs negative; positive amounts are credits (allocations out, reversed "
                      "accruals, roaming discounts). Mixed signs, so no share-of-total column.",
                      f"All figures are SUMIFS over table {d['table']} on sheet '{d['lines_sheet']}' "
                      f"({rec['lines']} lines); the month lines use an exact-match FILTER ('Jul-26' would be "
                      "read as a date)."],
            "source": "Source: FinLite drill-through (Data Integration) via sv_drill, and a live Smart View grid."}
    inv_lines = next(x for x in sections if x["title"].startswith("8."))["lines"]
    expected = {"mon": [groups(rows, "Period")[m] for m in months], "months": months,
                "inv": [groups(rows, "AP_INVOICE_NUMBER")[k] for k in inv],
                "inv_labels": [x["label"] for x in inv_lines]}
    return spec, expected


def read_answer(workbook: str, sheet: str) -> dict[str, tuple]:
    """label -> (NOK, Lines) from the answer sheet in the user's Excel (read only; attached through the Excel window
    like the server does - never GetObject(path), which can spawn a hidden Excel)."""
    import pythoncom
    from smartview_mcp import userexcel
    pythoncom.CoInitialize()
    name = Path(workbook).name.lower()
    for _pid, _h, app in userexcel.interactive_excels():
        for wb in app.Workbooks:
            if str(wb.Name).lower() == name:
                out = {}
                for r in wb.Worksheets(sheet).UsedRange.Value:
                    if len(r) >= 4 and isinstance(r[1], str) and r[1]:
                        out.setdefault(r[1], (r[2], r[3]))
                return out
    raise AssertionError(f"{name} is not open in the user's Excel")


def verify(workbook: str, sheet: str, exp: dict) -> None:
    """The answer sheet as Excel calculated it: months (FILTER form), Payables, suppliers, invoices."""
    got = read_answer(workbook, sheet)
    for m, (amt, n) in zip(exp["months"], exp["mon"]):
        print("month", m, got[m], flush=True)
        assert round(got[m][0], 2) == round(amt, 2) and got[m][1] == n, (m, got[m], amt, n)
    assert [got[m][1] for m in exp["months"]] == EXPECT["months"]
    pay = next(v for k, v in got.items() if k.startswith("Payables"))
    assert (round(pay[0], 2), pay[1]) == EXPECT["Payables"], pay
    for name in ("SUPPLIER 037 AS", "SUPPLIER 020 AS"):
        assert round(got[name][0], 2) == EXPECT[name], (name, got[name])
    assert got["(no supplier: accruals, feeds, journals)"][1] == EXPECT["no_supplier"]
    inv = [got[k] for k in exp["inv_labels"]]
    assert [(round(x[0], 2), x[1]) for x in inv] == [(round(e[0], 2), e[1]) for e in exp["inv"]], (inv, exp["inv"])
    print("verified in Excel: months (FILTER form), Payables, suppliers, invoices", flush=True)


async def main():
    async with Client(server.mcp) as c:
        d = await _call(c, "sv_drill", {"question": "Wholesale COGS Q3 26 – drill-through report", "cell": CELL,
                                        "target": "new"})
        rec = d["reconciliation"]
        print("drill:", d["workbook"], d["q"], d["lines_sheet"], d["table"], rec, d["timings"], flush=True)
        assert (rec["lines"], rec["lines_sum"], rec["status"]) == (EXPECT["lines"], EXPECT["sum"], "exact"), rec
        assert (rec["debit"], rec["credit"]) == (EXPECT["debit"], EXPECT["credit"]), rec
        db = drill_store.connect()
        try:
            _, rows = drill_store.load(db, d["drill_id"])
        finally:
            db.close()
        spec, exp = build_spec(d, rows)
        print("\n".join(spec["key_points"]), flush=True)
        a = await _call(c, "sv_ws_answer", {"q": d["q"], "spec": spec, "target": d["workbook"]})
        print("checks:", a["checks"], flush=True)
        print("formula_style:", {k: a["formula_style"].get(k) for k in ("ok", "checked", "counts")}, flush=True)
        print("warnings:", a["warnings"], flush=True)
        print(a["table"], flush=True)
        assert a["checks"] and all(v == "OK" for v in a["checks"].values()), a["checks"]
        assert a["formula_style"]["ok"], a["formula_style"]
        assert not a["warnings"], a["warnings"]
    verify(d["workbook"], a["answer_sheet"], exp)
    print("workbook:", d["workbook"])


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    asyncio.run(main())
