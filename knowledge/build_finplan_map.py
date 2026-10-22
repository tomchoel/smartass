"""Forretningskart, område 4: FinPlan – plan, prognose og KPI-er, og koblingen til FinMax.

Kilder: atlas (FinPlan Account: P&L, Ratios, Statistics til nivå 3), uttrekk map_finplan_actual (Actual/Working,
FY25 og FY26 × Dec/Sep) og map_finplan_plan (Target og Forecast Q326, FY26 Dec, Final), accounts_map.json (FinMax-tall
for avstemming). Lager «Forretningskart 4 - FinPlan og KPI.xlsx» og finplan_map.json.
"""
import datetime as dt
import json
import os
import sqlite3
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from build_accounts_map import ATLAS  # noqa: E402

XLSX = os.path.join(HERE, "Forretningskart 4 - FinPlan og KPI.xlsx")
JSON_OUT = os.path.join(HERE, "finplan_map.json")
sys.path.insert(0, os.path.join(os.path.expanduser("~"), ".claude", "skills", "tn-arbeidsbok"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import review  # noqa: E402

# FinPlan P&L-linje -> FinMax-konto (samme begrep, men ulik gruppering: avstem alltid)
FINMAX_EQUIV = {"Revenue": "C30", "COGS": "C40S", "Gross_Profit": "C48", "OPEX": "C70",
                "EBITDA before Other Items": "EBITDA_BOI", "Total Other Items": "C73", "EBITDA": "EBITDA"}
MEANING = {
    "P&L": "FinPlan resultatregnskap (plan, prognose og faktisk i FinPlans egne definisjoner).",
    "Revenue": "Inntekter i FinPlan-definisjon (gruppert annerledes enn FinMax C30 – avstem).",
    "COGS": "Varekost i FinPlan-definisjon.",
    "Gross_Profit": "Bruttofortjeneste = Revenue + COGS.",
    "OPEX": "Driftskostnader (lønn + andre driftskostnader).",
    "EBITDA before Other Items": "EBITDA før andre poster = Gross_Profit + OPEX.",
    "Total Other Items": "Andre poster (engangseffekter, gevinst/tap, nedbemanning) – Target har ingen.",
    "EBITDA": "EBITDA = EBITDA før andre poster + andre poster.",
    "Statistics": "Statistikk og KPI-er: årsverk, abonnement, ARPU, volumer.",
    "Ratios": "Nøkkeltall i prosent eller forholdstall (marginer, OPEX/inntekter, CAPEX/salg).",
}


def nodes():
    db = sqlite3.connect(ATLAS)
    q = "select name from member where connection='FinPlan' and dimension='Account' and parent=? order by ord"
    out = []

    def walk(n, parent, depth, root):
        kids = [k[0] for k in db.execute(q, (n,)).fetchall()]
        keep = root in ("P&L", "Ratios") or depth <= 3
        if keep and "Slettes" not in n and n != "Comments":
            out.append({"name": n, "parent": parent, "depth": depth, "root": root, "leaf": not kids, "n_children": len(kids)})
        for k in kids:
            walk(k, n, depth + 1, root)
    for r in ("P&L", "Statistics", "Ratios"):
        walk(r, "Account", 1, r)
    seen, uniq = set(), []
    for o in out:
        if o["name"] not in seen:
            seen.add(o["name"])
            uniq.append(o)
    return uniq


def read(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)["Data"]
    rows = list(ws.iter_rows(values_only=True))
    head = rows[0]
    return {r[0]: {head[i]: (r[i] if isinstance(r[i], (int, float)) else None) for i in range(1, len(head) - 1)}
            for r in rows[1:]}, os.path.basename(path)


def read_all(prefix):
    """Merge the latest extract of each chunk (map_finplan_actual_pl, _ratios, _stats1 ...)."""
    data, names = {}, []
    for path in _latest_chunks(prefix):
        d, n = read(path)
        data.update(d)
        names.append(n)
    return data, ", ".join(names)


# Members Smart View rejects (rc=1000) - found by the engine's bisection on 2026-10-14
REJECTED = {"Other": "Navnet finnes to ganger i Account (under «Other cash items» og «Statistics») – Smart View avviser det.",
            "SDx Revenue": "Smart View avviser medlemmet (rc=1000) – bruk «Fixed Other Retail Revenue: SDx» i P&L."}
# KPIs whose Accumulated value is a sum of the months (not the stock): read them Periodic
SUMS_MONTHS = ("Subscriptions", "ARPU", "%", "APPM", "CAPEX / Sales")
VOLUME_WORDS = ("Subscription", "Subscripton", "Volume", "ARPU", "APPM", "Customers", "users", "FTE", "Minutes", "Churn",
                "Gross Add", "Net Add", "Homes Passed", "share", "%", "CAPEX / Sales", "Cost per", "KPI", "antall",
                "sikkerhetssenter")


def main():
    act, f_act = read_all("map_finplan_actual")
    plan, f_plan = read_all("map_finplan_plan")
    fm = {a["code"]: a for a in json.load(open(os.path.join(HERE, "accounts_map.json"), encoding="utf-8"))["accounts"]}
    rows = []
    for nd in nodes():
        n = nd["name"]
        a, p = act.get(n, {}), plan.get(n, {})
        fy25 = a.get("FY25 / Dec")
        ytd26 = a.get("FY26 / Sep")
        tgt, fc = p.get("Target"), p.get("Forecast Q326")
        kind = {"P&L": "Resultat", "Statistics": "KPI / statistikk", "Ratios": "Nøkkeltall"}[nd["root"]]
        eq = FINMAX_EQUIV.get(n)
        fm25 = fm.get(eq, {}).get("fy2025_nok") if eq else None
        pit = []
        if eq and fy25 is not None and fm25:
            diff = fy25 - fm25
            pit.append(f"Avstemt mot FinMax {eq} 2025: differanse {diff / 1e6:,.1f} MNOK ({diff / abs(fm25):+.1%}) – "
                       "ulik gruppering av konti.".replace(",", " "))
        if n in ("Total Other Items",) or (nd["root"] == "P&L" and tgt in (None, 0) and fy25):
            pit.append("Target (budsjett) har ingen verdi her – sammenlign mot budsjett bare til EBITDA før/etter andre poster.")
        volume = any(s.lower() in n.lower() for s in VOLUME_WORDS)
        if nd["root"] in ("Statistics", "Ratios") and volume:
            pit.append("Volum/KPI – finnes ikke i FinMax (unntatt årsverk: FinMax STA20Y).")
        elif nd["root"] == "Statistics" and n not in REJECTED:
            pit.append("Beløp i Statistics (FinPlans egen gruppering) – ikke en del av P&L-summen.")
        if nd["root"] in ("Statistics", "Ratios") and fy25 is None and ytd26 is None and (tgt or fc):
            pit.append("Bare i plan/prognose: faktiske volumer ligger ikke i FinPlan for BANORWAY (Actual er tom).")
        if nd["root"] in ("Statistics", "Ratios") and any(s in n for s in SUMS_MONTHS):
            pit.append("Les som Periodic per måned: Accumulated summerer månedene for denne typen KPI "
                       "(f.eks. Mobile Subscriptions hittil sep = 78 mill. mot 8,7 mill. i beholdning).")
        if n in REJECTED:
            pit.append(REJECTED[n])
        elif all(v in (None, 0) for v in (fy25, ytd26, tgt, fc)):
            pit.append("Ingen verdi for BANORWAY i disse periodene.")
        meaning = MEANING.get(n) or f"{kind} under {nd['parent']}: {n}."
        rows.append({"member": n, "parent": nd["parent"], "depth": nd["depth"], "root": nd["root"], "kind": kind,
                     "leaf": nd["leaf"], "actual_fy25": fy25, "actual_ytd_sep26": ytd26, "target_fy26": tgt,
                     "forecast_q326_fy26": fc, "finmax_equivalent": f"ac.{eq}" if eq else None,
                     "finmax_fy25": fm25, "meaning": meaning, "pitfalls": " ".join(pit)})
    status, review_label = review.apply("finplan", rows)
    meta = {"built": dt.datetime.now().isoformat(timespec="seconds"), "status": status, "review_label": review_label,
            "sources": [f_act, f_plan, "atlas.db (FinPlan Account)", "accounts_map.json (FinMax-avstemming)"],
            "notes": ["FinPlan-entitet BANORWAY = Telenor Norge. Faktisk = Scenario Actual, versjon Working; "
                      "Target og Forecast Q<q><yy> = versjon Final.",
                      "Faktisk mot plan gjøres helt i FinPlan (samme definisjoner), med en avstemmingslinje mot FinMax EBITDA.",
                      "Beløp i NOK; Accumulated (hittil i år / helår ved Dec). Volumer, ARPU og prosenter: les Periodic "
                      "per måned – Accumulated summerer månedene for flere av dem.",
                      "Faktiske volumer (abonnement, kunder) finnes ikke i FinPlan for BANORWAY – bare i Target/prognose.",
                      "FinPlan BANORWAY og FinMax un.SY grupperer ulikt: EBITDA 2025 er 14,16 mrd i FinPlan mot "
                      "12,40 mrd i FinMax. Sammenlign aldri linjer på tvers uten avstemmingslinje.",
                      "Smart View avviser enkelte medlemmer (rc=1000), f.eks. «Other» (navnet finnes to ganger). "
                      "Motoren finner og utelater dem nå selv og varsler i warnings."]}
    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "accounts": rows}, f, ensure_ascii=False, indent=1)
    print(len(rows), "FinPlan-konti ->", JSON_OUT)
    build_xlsx(rows, meta)


def _latest_chunks(prefix):
    """Newest file per chunk title: <stamp>_FinPlan_<prefix>[_<chunk>].xlsx."""
    import glob
    import re
    from build_accounts_map import OUT_DIR as out
    latest = {}
    for f in sorted(glob.glob(os.path.join(out, f"*_FinPlan_{prefix}*.xlsx"))):
        m = re.search(rf"_FinPlan_({re.escape(prefix)}(?:_\w+)?)\.xlsx$", f)
        if m:
            latest[m.group(1)] = f
    if not latest:
        sys.exit(f"Mangler FinPlan-uttrekk '{prefix}' i {out}")
    return list(latest.values())


def build_xlsx(rows, meta):
    import pythoncom
    import win32com.client as wc
    import tn_arbeidsbok as T
    pythoncom.CoInitialize()
    xl = wc.DispatchEx("Excel.Application")
    pid = T.excel_pid(xl)
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.ScreenUpdating = False
        wb = xl.Workbooks.Add()
        T.ensure_styles(xl, wb)
        T.ensure_table_style(wb)
        info = wb.Worksheets(1)
        info.Name = "Les meg"
        ws = wb.Worksheets.Add(After=info)
        ws.Name = "FinPlan"
        lst = wb.Worksheets.Add(After=ws)
        lst.Name = "Lister"
        lst.Range("B4:B6").Value = (("OK",), ("Endre",), ("Usikker",))
        T.name(wb, "Liste_Vurdering", "=Lister!$B$4:$B$6")
        lst.Visible = 0
        T.setup_sheet(info, "Forretningskart 4 – FinPlan: plan, prognose og KPI",
                      f"FinPlan · BANORWAY · {meta['status']} · bygget {meta['built'][:16].replace('T', ' ')}")
        lines = [
            "Hva er dette: FinPlans resultatlinjer, nøkkeltall og KPI-er (årsverk, abonnement, ARPU, volumer) med faktisk 2025,",
            "hittil i 2026, budsjett (Target) og siste prognose (Forecast Q326), og hvilke linjer som tilsvarer FinMax.",
            "",
            "Slik går du gjennom: arket «FinPlan». Sjekk «Betydning», «FinMax-tilsvarende» og «Fallgruver», velg",
            "OK / Endre / Usikker i «Vurdering» og skriv rettelser i «Kommentar».",
            "",
            "Kilder: " + "; ".join(meta["sources"]),
            "Bygget av smartass\\knowledge\\build_finplan_map.py.",
        ]
        for i, t in enumerate(lines):
            c = info.Cells(4 + i, 2)
            c.Value = "'" + t if t else None
            c.Style = "TN Tekst"
        info.Columns("B").ColumnWidth = 115
        T.setup_sheet(ws, "FinPlan – konti, KPI-er og kobling til FinMax", f"{meta['review_label']} · BANORWAY (Telenor Norge)",
                      freeze_row=4, freeze_col=2)
        headers = ["Konto (FinPlan)", "Overordnet", "Nivå", "Type", "Faktisk 2025", "Faktisk hittil sep 26",
                   "Target 2026", "Forecast Q326 2026", "FinMax-tilsvarende", "FinMax 2025", "Betydning",
                   "Fallgruver", "Vurdering", "Kommentar"]
        data = [["'" + r["member"], "'" + str(r["parent"]), r["depth"], r["kind"], r["actual_fy25"], r["actual_ytd_sep26"],
                 r["target_fy26"], r["forecast_q326_fy26"], r["finmax_equivalent"], r["finmax_fy25"], r["meaning"],
                 r["pitfalls"] or None, None, None] for r in rows]
        xl.AutoCorrect.AutoFillFormulasInLists = False
        lo = T.add_table(ws, "B4", headers, data, "tblFinPlan")
        for nm, st, wdt in (("Konto (FinPlan)", "TN Tabell tekst", 36), ("Overordnet", "TN Tabell kort", 22),
                            ("Nivå", "TN Tabell heltall", 6), ("Type", "TN Tabell kort", 14),
                            ("Faktisk 2025", "TN Tabell heltall", 16), ("Faktisk hittil sep 26", "TN Tabell heltall", 16),
                            ("Target 2026", "TN Tabell heltall", 16), ("Forecast Q326 2026", "TN Tabell heltall", 16),
                            ("FinMax-tilsvarende", "TN Tabell kort", 14), ("FinMax 2025", "TN Tabell heltall", 16),
                            ("Betydning", "TN Tabell tekst", 44), ("Fallgruver", "TN Tabell tekst", 44),
                            ("Vurdering", "TN Status", 11), ("Kommentar", "TN Tabell input", 34)):
            T.style_col(lo, nm, st, wdt)
        T.list_validation(T.col(lo, "Vurdering"), "Liste_Vurdering")
        T.status_colors(T.col(lo, "Vurdering"), {"OK": "ferdig", "Endre": "venter", "Usikker": "stopp"})
        T.page_setup(ws)
        info.Activate()
        T.label_internal(wb)
        if os.path.exists(XLSX):
            os.remove(XLSX)
        wb.SaveAs(XLSX, 51)
        wb.Close(False)
    finally:
        T.shutdown(xl, pid)
    T.freeze_in_file(XLSX)
    T.gridlines_off_in_file(XLSX)
    print("skrev", XLSX)


if __name__ == "__main__":
    main()
