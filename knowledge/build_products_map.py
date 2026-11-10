"""Forretningskart, område 2: produkter og inntekt/varekost (FinMax pd.SY <-> Produktregnskap).

Kilder:
  atlas (FinMax All Products, pd.SY.Mobile + pd.SY.FIXED), uttrekk map_products_pnl og map_products_revenue_by_division
  (smartview sv_query save_xlsx), source_pd_to_prodreg_code_2025.json (FinMax-produkt -> produktregnskapskode per
  divisjon, fra 2025-filene i 2 Input\\20 Inntekt-varekost), source_product_list_2025.json (offisiell produktliste),
  source_pd_unresolved_2025.json (produkter uten fast kode), concepts.json (hvilke begrep produktgruppen inngår i).
Lager «Forretningskart 2 - Produkter.xlsx» (gjennomgang) og products_map.json (for skill og server).
"""
import datetime as dt
import glob
import json
import os
import sqlite3
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "engine"))
from smartass.core import config as _cfg  # noqa: E402

OUT_DIR = str(_cfg.OUT_DIR)
ATLAS = str(_cfg.ATLAS_DB)
DIVISIONS = {
    "un.SY.2SSM": "SSM Security Solutions & Mobile", "un.SY.2EI": "E&I Entertainment & Internet",
    "un.SY.2BUS": "Business", "un.SY.2WHS": "Wholesale", "un.SY.2TEC": "Technology", "un.SY.2IT": "IT",
    "un.SY.2ADM": "Administrasjon", "un.SY.2NRD_N": "Nordics", "un.SY.2CWK": "CA, WS & Kystradio",
    "un.SY.3KYS_C": "Kystradio", "un.SY.3FIB": "TnFiber IC", "un.SY.5OMS_A": "Omstilling",
    "un.SY.2I15": "IFRS15-justering", "un.SY.2I16": "IFRS16-justering", "un.SY.E171": "IFRS16 elim. korr.",
    "un.SY.DE11": "STAR", "un.SY.G916": "IMEI korr.", "un.SY.SYi": "Vurdering", "un.SY.2CON": "Consumer (nedlagt)",
    "un.SY.2IG": "Infrastructure Growth (nedlagt)", "un.SY.OUTDATED": "Utgåtte noder",
}


def latest(title):
    files = sorted(glob.glob(os.path.join(OUT_DIR, f"*_FinMax_{title}.xlsx")))
    if not files:
        sys.exit(f"Mangler uttrekk '{title}' i {OUT_DIR} - kjør sv_query med save_xlsx og den tittelen først.")
    return files[-1]


def read_pull(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)["Data"]
    rows = list(ws.iter_rows(values_only=True))
    head = rows[0]
    out = {}
    for r in rows[1:]:
        out[r[0]] = {head[i]: (r[i] if isinstance(r[i], (int, float)) else 0.0) for i in range(1, len(head) - 1)}
    return out, os.path.basename(path)

XLSX = os.path.join(HERE, "Forretningskart 2 - Produkter.xlsx")
JSON_OUT = os.path.join(HERE, "products_map.json")
sys.path.insert(0, os.path.join(os.path.expanduser("~"), ".claude", "skills", "tn-arbeidsbok"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import review  # noqa: E402


PACK = os.path.join(os.path.dirname(HERE), "packs", "produktregnskap", "knowledge")


def load(name):
    folder = PACK if name.startswith("source_") else HERE        # Produktregnskap inputs live in the pack
    with open(os.path.join(folder, name), encoding="utf-8") as f:
        return json.load(f)


def groups():
    db = sqlite3.connect(ATLAS)
    q = "select name,label from member where connection='FinMax' and dimension='All Products' and parent=? order by ord"
    out = []

    def walk(n, parent, depth):
        kids = db.execute(q, (n,)).fetchall()
        if not kids:
            return [n]
        lab = db.execute("select label from member where connection='FinMax' and dimension='All Products' and name=? "
                         "limit 1", (n,)).fetchone()
        rec = {"name": n, "label": (lab[0] if lab else n).replace(" (pd.SY)", ""), "parent": parent, "depth": depth}
        out.append(rec)
        leaves = []
        for k, _ in kids:
            leaves += walk(k, n, depth + 1)
        rec["leaves"] = leaves
        return leaves

    for root in ("pd.SY.Mobile", "pd.SY.FIXED"):
        walk(root, "pd.SY", 0)
    return out


def main():
    pnl, f_pnl = read_pull(latest("map_products_pnl"))
    rdiv, f_div = read_pull(latest("map_products_revenue_by_division"))
    mapping = load("source_pd_to_prodreg_code_2025.json")
    plist = {}
    for r in load("source_product_list_2025.json"):        # SAS codes (C2x_) and CostPerform codes (L50x_) alike
        for k in ("code_sas", "code_costperform"):
            if r.get(k):
                plist.setdefault(r[k], r)
    KEYED = "Consumer fast – fordeles med FinPlan-nøkler"
    unresolved = defaultdict(list)
    for u in load("source_pd_unresolved_2025.json"):
        if u.get("pd_code"):
            unresolved[u["pd_code"]].append(u)
    concept_members = defaultdict(set)
    for c in load("concepts.json").get("concepts", []):
        for b in c.get("blocks", []):
            for it in b.get("items", []):
                for src in (it, it.get("base") or {}, it.get("key_den") or {}):
                    for p in src.get("products") or []:
                        concept_members[p].add(c["id"])
    by_pd = defaultdict(list)
    for m in mapping:
        by_pd[m["pd_code"]].append(m)

    rows = []
    for g in groups():
        n = g["name"]
        p = pnl.get(n, {})
        c30, c40, c70, boi = (p.get(f"ac.{a} / Dec 25", 0.0) for a in ("C30", "C40S", "C70", "EBITDA_BOI"))
        c30_26 = p.get("ac.C30 / Sep 26", 0.0)
        d = rdiv.get(n, {})
        parts = {k: v for k, v in d.items() if abs(v) >= 1}
        tot = sum(abs(v) for v in parts.values())
        top = max(parts, key=lambda k: abs(parts[k])) if parts else None
        # produktregnskapskoder: vekt = |C30|+|C40S| på bladnivå (fra divisjonsfilene)
        w = defaultdict(float)
        divs_by_code = defaultdict(set)
        unres_w = mapped_w = 0.0
        for leaf in g["leaves"]:
            code = leaf.split(".")[-1]
            for m in by_pd.get(code, []):
                x = abs(m.get("c30") or 0) + abs(m.get("c40s") or 0)
                w[m["prodreg_code"]] += x
                divs_by_code[m["prodreg_code"]].add(m["division"])
                mapped_w += x
            for u in unresolved.get(code, []):
                x = abs(u.get("c30") or 0) + abs(u.get("c40s") or 0)
                label = KEYED if str(u.get("prodreg_code") or "").startswith("skal ha egen") else "uten kode i 2025-filene"
                w[label] += x
                divs_by_code[label].add(u["division"])
                unres_w += x
        tw = sum(w.values())
        codes = sorted(w.items(), key=lambda kv: -kv[1])[:4]
        code_txt = "; ".join(f"{c} ({v / tw:.0%})" for c, v in codes if tw) if tw else ""
        top_code = codes[0][0] if codes else None
        fam = plist.get(top_code) if top_code else None
        PREFIX_FAMILY = {"C2B": "Business", "C2M": "Mobil / Consumer", "C2C": "Fixed / Consumer", "C1F": "Fixed / Consumer",
                         "C2W": "Fixed / Wholesale", "L50W": "Wholesale", "C2A": "Administrasjon / annet",
                         "L50A": "Administrasjon / annet"}
        if fam:
            family = f"{fam['value_chain']} / {fam['market']} / {fam['product_group']}"
            vstream = fam.get("valuestream") or ""
        elif top_code == KEYED:
            family, vstream = "Fixed / Consumer / Internett og TV (fordeles med nøkler)", "Internet @home / Entertainment"
        elif top_code and top_code.split("_")[0] in PREFIX_FAMILY:
            family, vstream = f"{PREFIX_FAMILY[top_code.split('_')[0]]} (kode utenfor produktlisten)", ""
        else:
            family, vstream = "", ""
        product_names = "; ".join(sorted({plist[c]['product'] for c, _ in codes if c in plist}))
        pit = []
        if tw and w.get(KEYED, 0) >= 0.2 * tw:
            pit.append(f"{w[KEYED] / tw:.0%} har ingen fast produktkode: Consumer fast (HFC/fiber/TV) fordeles til "
                       "C2C_AKSESS_/C2C_TV_/C2C_FTB med FinPlan-nøkler etter årsslutt.")
        if tw and w.get("uten kode i 2025-filene", 0) >= 0.2 * tw:
            pit.append("En stor del mangler produktkode i 2025-filene.")
        prefixes = {c.split("_")[0] for c in w if c in plist or c.startswith(("C1", "C2", "L50"))}
        if len(prefixes) > 1:
            pit.append(f"Får ulik produktkode per divisjon ({', '.join(sorted(prefixes))}): Consumer C2C/C2M, Business C2B, "
                       "Wholesale C2W – slå opp på divisjon + produkt.")
        if abs(c30) < 1 and abs(c40) >= 1:
            pit.append("Kostnadsprodukt: varekost uten inntekt på produktet.")
        if abs(c70) >= 1e6:
            pit.append(f"Har også driftskostnader bokført på produktet (C70 {c70 / 1e6:,.1f} MNOK i 2025).".replace(",", " "))
        hit_concepts = set()
        for leaf in [n] + g["leaves"]:
            hit_concepts |= concept_members.get(leaf, set())
        if hit_concepts:
            pit.append("Inngår i begrep: " + ", ".join(sorted(hit_concepts)) + " (bruk sv_concept).")
        if abs(c30) < 1 and abs(c40) < 1 and abs(c70) < 1:
            pit.append("Ingen bevegelse i 2025.")
        meaning = g["label"]
        if family:
            meaning += f" – {family}"
        if product_names:
            meaning += f" (produktregnskap: {product_names})"
        rows.append({"member": n, "code": n.split(".")[-1], "label": g["label"], "parent": g["parent"].split(".")[-1],
                     "depth": g["depth"], "n_products": len(g["leaves"]), "fy2025_c30": round(c30, 2),
                     "fy2025_c40s": round(c40, 2), "fy2025_c70": round(c70, 2), "fy2025_ebitda_boi": round(boi, 2),
                     "ytd_sep26_c30": round(c30_26, 2), "top_division": DIVISIONS.get(top) if top else None,
                     "top_division_share": abs(parts[top]) / tot if top and tot else None,
                     "prodreg_codes": code_txt, "family": family, "value_stream": vstream, "meaning": meaning,
                     "pitfalls": " ".join(pit),
                     "divisions_c30_fy2025": {DIVISIONS.get(k, k): round(v, 2) for k, v in parts.items()}})
    status, review_label = review.apply("products", rows)
    meta = {"built": dt.datetime.now().isoformat(timespec="seconds"), "status": status, "review_label": review_label,
            "sources": [f_pnl, f_div, "source_pd_to_prodreg_code_2025.json (2025 Inntekt-varekost per divisjon)",
                        "source_product_list_2025.json (Produkter & nettverk_2025)", "atlas.db (FinMax All Products)"],
            "notes": ["FinMax-produkter (pd.SY) er Telenor Norge sine. Mobile = M-koder, Fixed = P/DP/PR-koder.",
                      "Produktregnskapskode og familie kommer fra 2025-filene; vekt = |C30|+|C40S| på bladnivå.",
                      "Beløp: FinMax Actual NOK, un.SY, FY2025 = Dec 25 Accumulated."]}
    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "products": rows}, f, ensure_ascii=False, indent=1)
    print(len(rows), "produktgrupper ->", JSON_OUT)
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
        ws.Name = "Produkter"
        lst = wb.Worksheets.Add(After=ws)
        lst.Name = "Lister"
        lst.Range("B4:B6").Value = (("OK",), ("Endre",), ("Usikker",))
        T.name(wb, "Liste_Vurdering", "=Lister!$B$4:$B$6")
        lst.Visible = 0
        T.setup_sheet(info, "Forretningskart 2 – Produkter og inntekt/varekost",
                      f"FinMax pd.SY · Telenor Norge · FY2025 · {meta['status']} · bygget {meta['built'][:16].replace('T', ' ')}")
        lines = [
            "Hva er dette: alle produktgrupper i FinMax (Telenor Norge, pd.SY) med inntekt, varekost og driftskostnad bokført",
            "på produktet, hvilken divisjon som selger det, og koblingen til Produktregnskapets produktkoder og -familier.",
            "",
            "Slik går du gjennom: arket «Produkter». Sjekk «Familie», «Produktregnskapskoder», «Betydning» og «Fallgruver»,",
            "velg OK / Endre / Usikker i «Vurdering» og skriv rettelser i «Kommentar». Tallene er fakta fra FinMax.",
            "",
            "Produktregnskapskoder: koder fra 2025-filene i 2 Input\\20 Inntekt-varekost, vektet med |inntekt|+|varekost|.",
            "Consumer fast bredbånd/TV har ingen fast kode per produkt – den fordeles med FinPlan-nøkler etter årsslutt.",
            "",
            "Kilder: " + "; ".join(meta["sources"]),
            "Bygget av smartass\\knowledge\\build_products_map.py.",
        ]
        for i, t in enumerate(lines):
            c = info.Cells(4 + i, 2)
            c.Value = "'" + t if t else None
            c.Style = "TN Tekst"
        info.Columns("B").ColumnWidth = 115
        T.setup_sheet(ws, "Produkter – FinMax ↔ Produktregnskap", f"{meta['review_label']} · beløp i NOK, Telenor Norge (un.SY)",
                      freeze_row=4, freeze_col=3)
        headers = ["Kode", "Navn (FinMax)", "Overordnet", "Nivå", "Antall produkter", "Inntekt 2025", "Varekost 2025",
                   "Driftskost. 2025", "EBITDA f. a.p. 2025", "Inntekt hittil sep 26", "Største divisjon", "Andel",
                   "Familie (produktliste)", "Verdistrøm", "Produktregnskapskoder", "Betydning",
                   "Fallgruver", "Vurdering", "Kommentar"]
        data = [["'" + r["code"], "'" + r["label"], "'" + r["parent"], r["depth"], r["n_products"], r["fy2025_c30"],
                 r["fy2025_c40s"], r["fy2025_c70"], r["fy2025_ebitda_boi"], r["ytd_sep26_c30"], r["top_division"],
                 r["top_division_share"], r["family"] or None, r["value_stream"] or None, r["prodreg_codes"] or None,
                 r["meaning"], r["pitfalls"] or None, None, None] for r in rows]
        xl.AutoCorrect.AutoFillFormulasInLists = False
        lo = T.add_table(ws, "B4", headers, data, "tblProdukter")
        for nm, st, wdt in (("Kode", "TN Tabell ID", 10), ("Navn (FinMax)", "TN Tabell tekst", 34),
                            ("Overordnet", "TN Tabell kort", 10), ("Nivå", "TN Tabell heltall", 6),
                            ("Antall produkter", "TN Tabell heltall", 9), ("Inntekt 2025", "TN Tabell heltall", 15),
                            ("Varekost 2025", "TN Tabell heltall", 15), ("Driftskost. 2025", "TN Tabell heltall", 14),
                            ("EBITDA f. a.p. 2025", "TN Tabell heltall", 15), ("Inntekt hittil sep 26", "TN Tabell heltall", 15),
                            ("Største divisjon", "TN Tabell kort", 20), ("Andel", "TN Tabell prosent", 8),
                            ("Familie (produktliste)", "TN Tabell tekst", 26), ("Verdistrøm", "TN Tabell kort", 18),
                            ("Produktregnskapskoder", "TN Tabell tekst", 34), ("Betydning", "TN Tabell tekst", 46),
                            ("Fallgruver", "TN Tabell tekst", 40), ("Vurdering", "TN Status", 11),
                            ("Kommentar", "TN Tabell input", 34)):
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
