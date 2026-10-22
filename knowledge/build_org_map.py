"""Forretningskart, område 3: organisasjon og kostnadsbase (FinMax un.SY, 2026-strukturen).

Kilder: atlas (FinMax All Entities under un.SY), uttrekk map_org_costbase (sv_query save_xlsx: C30, C40S, C50S, C60,
EBITDA_BOI, STA20Y × Dec 25 / Sep 26), og reorg-kartleggingen 2025→2026
(prodreg2026\\01 Struktur og mapping\\2026_Reorg_Mapping_Ansvarssted_Verdistrom.xlsx, ark Src_GLCC_2025).
Lager «Forretningskart 3 - Organisasjon og kostnader.xlsx» og org_map.json.
"""
import datetime as dt
import json
import os
import sqlite3
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from build_accounts_map import ADJUST, ATLAS, DIVISIONS, latest, read_pull  # noqa: E402

XLSX = os.path.join(HERE, "Forretningskart 3 - Organisasjon og kostnader.xlsx")
JSON_OUT = os.path.join(HERE, "org_map.json")
REORG = os.path.join(os.path.dirname(os.path.dirname(HERE)), "01 Struktur og mapping",
                     "2026_Reorg_Mapping_Ansvarssted_Verdistrom.xlsx")
sys.path.insert(0, os.path.join(os.path.expanduser("~"), ".claude", "skills", "tn-arbeidsbok"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import review  # noqa: E402
CODE_2026 = {"un.SY.2SSM": "SSM", "un.SY.2EI": "EI", "un.SY.2BUS": "BUS", "un.SY.2WHS": "WHS", "un.SY.2TEC": "TEC",
             "un.SY.2IT": "IT", "un.SY.2ADM": "ADM", "un.SY.2NRD_N": "NRD", "un.SY.2CWK": "CSW", "un.SY.3KYS_C": "KYS",
             "un.SY.3FIB": "FIB", "un.SY.5OMS_A": "ADM", "un.SY.2CON": "CON", "un.SY.2IG": "IG"}


def tree():
    db = sqlite3.connect(ATLAS)
    q = "select name from member where connection='FinMax' and dimension='All Entities' and parent=? order by ord"
    lab = lambda n: (db.execute("select label from member where connection='FinMax' and dimension='All Entities' and "
                                "name=? limit 1", (n,)).fetchone() or [n])[0].replace(" (un.SY)", "")
    out = []

    def walk(n, parent, depth, division):
        kids = [k[0] for k in db.execute(q, (n,)).fetchall()]
        out.append({"name": n, "label": lab(n), "parent": parent, "depth": depth, "division": division,
                    "leaf": not kids, "n_children": len(kids)})
        for k in kids:
            walk(k, n, depth + 1, division or k)
    walk("un.SY", None, 0, None)
    return out


def reorg_2025():
    import openpyxl
    if not os.path.exists(REORG):
        return {}
    ws = openpyxl.load_workbook(REORG, read_only=True, data_only=True)["Src_GLCC_2025"]
    return {r[1]: {"div": r[4], "div_name": r[5], "parent": r[6]} for r in ws.iter_rows(min_row=2, values_only=True) if r[1]}


def main():
    cb, f_cb = read_pull(latest("map_org_costbase"))
    r25 = reorg_2025()
    rows = []
    for e in tree():
        n = e["name"]
        v = cb.get(n, {})
        g = lambda a, p="Dec 25": v.get(f"ac.{a} / {p}", 0.0)
        c30, c40, c50, c60, boi, fte = g("C30"), g("C40S"), g("C50S"), g("C60"), g("EBITDA_BOI"), g("STA20Y")
        opex26 = g("C50S", "Sep 26") + g("C60", "Sep 26")
        div = e["division"]
        div_name = DIVISIONS.get(div, div.split(".")[-1] if div else "Telenor Norge AS")
        kind = ("Selskap" if e["depth"] == 0 else "Divisjon" if e["depth"] == 1 else
                "Koststed" if e["leaf"] else "Avdeling/gruppe")
        old = r25.get(n)
        moved = None
        if old and div in CODE_2026 and old["div"] and old["div"] != CODE_2026[div]:
            moved = f"{old['div']} → {CODE_2026[div]}"
        pit = []
        if moved:
            pit.append(f"Flyttet i reorg 1. juni 2026 ({moved}); historikken er omarbeidet til ny organisasjon.")
        if not old and e["depth"] >= 2 and r25:
            pit.append("Fantes ikke i 2025-strukturen (ny i 2026 eller ikke i GL-CC 2025).")
        if div in ADJUST or n in ADJUST:
            pit.append("Justerings-/korreksjonsenhet (IFRS, eliminering, vurdering) – ikke en driftsenhet.")
        if div in ("un.SY.2CON", "un.SY.2IG", "un.SY.OUTDATED"):
            pit.append("Nedlagt/utgått node – tall omarbeidet til nye divisjoner.")
        if all(abs(x) < 1 for x in (c30, c40, c50, c60, fte)):
            pit.append("Ingen bevegelse i 2025.")
        if abs(c30) >= 1e6 and e["leaf"]:
            pit.append("Inntektsbærende koststed (inntekt bokført på koststedet).")
        meaning = f"{kind} i {div_name}" + (f" (2025: {old['div_name']})" if moved and old else "") + f": {e['label']}"
        rows.append({"member": n, "code": n.split(".")[-1] if n != "un.SY" else "un.SY", "label": e["label"],
                     "parent": (e["parent"] or "").split(".")[-1], "depth": e["depth"], "kind": kind,
                     "division_2026": div_name, "division_2025": old["div_name"] if old else None, "moved": moved,
                     "fy2025_c30": round(c30, 2), "fy2025_c40s": round(c40, 2), "fy2025_c50s": round(c50, 2),
                     "fy2025_c60": round(c60, 2), "fy2025_ebitda_boi": round(boi, 2), "fte_dec25": round(fte, 1),
                     "ytd_sep26_opex": round(opex26, 2), "meaning": meaning, "pitfalls": " ".join(pit)})
    status, review_label = review.apply("org", rows)
    meta = {"built": dt.datetime.now().isoformat(timespec="seconds"), "status": status, "review_label": review_label,
            "sources": [f_cb, "atlas.db (FinMax All Entities)", "2026_Reorg_Mapping_Ansvarssted_Verdistrom.xlsx (Src_GLCC_2025)"],
            "notes": ["Strukturen er FinMax sin i dag (2026-organisasjonen); 2025-divisjon fra GL-CC 2025.",
                      "Årsverk = STA20Y Accumulated Dec 25 (beholdning ved årsslutt).",
                      "Driftskostnad = C50S lønn + C60 andre driftskostnader; kostnader er negative."]}
    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "entities": rows}, f, ensure_ascii=False, indent=1)
    print(len(rows), "enheter ->", JSON_OUT, "| flyttet:", sum(1 for r in rows if r["moved"]))
    build_xlsx(rows, meta)


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
        ws.Name = "Organisasjon"
        lst = wb.Worksheets.Add(After=ws)
        lst.Name = "Lister"
        lst.Range("B4:B6").Value = (("OK",), ("Endre",), ("Usikker",))
        T.name(wb, "Liste_Vurdering", "=Lister!$B$4:$B$6")
        lst.Visible = 0
        T.setup_sheet(info, "Forretningskart 3 – Organisasjon og kostnadsbase",
                      f"FinMax un.SY · 2026-organisasjonen · FY2025 · {meta['status']} · bygget {meta['built'][:16].replace('T', ' ')}")
        lines = [
            "Hva er dette: alle enheter i Telenor Norge (un.SY) i FinMax – divisjoner, avdelinger og koststeder – med hva de",
            "koster og tjener, årsverk, hvilken divisjon de tilhørte i 2025 (reorg 1. juni 2026) og fallgruver.",
            "",
            "Slik går du gjennom: arket «Organisasjon». Filtrer på divisjon eller «Flyttet», sjekk «Betydning» og «Fallgruver»,",
            "velg OK / Endre / Usikker i «Vurdering» og skriv rettelser i «Kommentar». Tallene er fakta fra FinMax.",
            "",
            "Kilder: " + "; ".join(meta["sources"]),
            "Bygget av smartass\\knowledge\\build_org_map.py.",
        ]
        for i, t in enumerate(lines):
            c = info.Cells(4 + i, 2)
            c.Value = "'" + t if t else None
            c.Style = "TN Tekst"
        info.Columns("B").ColumnWidth = 115
        T.setup_sheet(ws, "Organisasjon og kostnadsbase – FinMax", f"{meta['review_label']} · beløp i NOK, FY2025",
                      freeze_row=4, freeze_col=3)
        headers = ["Kode", "Navn (FinMax)", "Overordnet", "Nivå", "Type", "Divisjon 2026", "Divisjon 2025", "Flyttet",
                   "Inntekt 2025", "Varekost 2025", "Lønn 2025", "Andre driftskost. 2025", "EBITDA f. a.p. 2025",
                   "Årsverk des 25", "Driftskost. hittil sep 26", "Betydning", "Fallgruver", "Vurdering", "Kommentar"]
        data = [["'" + r["code"], "'" + r["label"], "'" + r["parent"], r["depth"], r["kind"], r["division_2026"],
                 r["division_2025"], r["moved"], r["fy2025_c30"], r["fy2025_c40s"], r["fy2025_c50s"], r["fy2025_c60"],
                 r["fy2025_ebitda_boi"], r["fte_dec25"], r["ytd_sep26_opex"], r["meaning"], r["pitfalls"] or None,
                 None, None] for r in rows]
        xl.AutoCorrect.AutoFillFormulasInLists = False
        lo = T.add_table(ws, "B4", headers, data, "tblOrganisasjon")
        for nm, st, wdt in (("Kode", "TN Tabell ID", 11), ("Navn (FinMax)", "TN Tabell tekst", 34),
                            ("Overordnet", "TN Tabell kort", 11), ("Nivå", "TN Tabell heltall", 6), ("Type", "TN Tabell kort", 13),
                            ("Divisjon 2026", "TN Tabell kort", 22), ("Divisjon 2025", "TN Tabell kort", 18),
                            ("Flyttet", "TN Tabell kort", 11), ("Inntekt 2025", "TN Tabell heltall", 15),
                            ("Varekost 2025", "TN Tabell heltall", 15), ("Lønn 2025", "TN Tabell heltall", 15),
                            ("Andre driftskost. 2025", "TN Tabell heltall", 15), ("EBITDA f. a.p. 2025", "TN Tabell heltall", 15),
                            ("Årsverk des 25", "TN Tabell tall", 10), ("Driftskost. hittil sep 26", "TN Tabell heltall", 15),
                            ("Betydning", "TN Tabell tekst", 46), ("Fallgruver", "TN Tabell tekst", 40),
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
