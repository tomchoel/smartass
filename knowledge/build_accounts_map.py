"""Forretningskart, område 1: konti og bokføring (FinMax).

Leser kontotreet fra atlaset og fire FinMax-uttrekk (laget med smartview sv_query, save_xlsx):
  map_accounts_by_division_FY25, map_accounts_product_tag, map_accounts_reserve_tag (+ function_tag: ubrukt i un.SY)
og bygger:
  knowledge\\Forretningskart 1 - Konti og bokføring.xlsx   (gjennomgangsark i TN arbeidsbok-stil, Vurdering OK/Endre/Usikker)
  knowledge\\accounts_map.json                              (maskinlesbar kopi for skill og server)
Kjør: python build_accounts_map.py   (med svkit-venv; bygger i en egen skjult Excel)
"""
import datetime as dt
import glob
import json
import os
import sqlite3
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "engine"))
from smartass.core import config as _cfg  # noqa: E402

OUT_DIR = str(_cfg.OUT_DIR)
XLSX = os.path.join(HERE, "Forretningskart 1 - Konti og bokføring.xlsx")
JSON_OUT = os.path.join(HERE, "accounts_map.json")
ATLAS = str(_cfg.ATLAS_DB)
sys.path.insert(0, os.path.join(os.path.expanduser("~"), ".claude", "skills", "tn-arbeidsbok"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import review  # noqa: E402

DIVISIONS = {
    "un.SY.2SSM": "SSM Security Solutions & Mobile", "un.SY.2EI": "E&I Entertainment & Internet",
    "un.SY.2BUS": "Business", "un.SY.2WHS": "Wholesale", "un.SY.2TEC": "Technology", "un.SY.2IT": "IT",
    "un.SY.2ADM": "Administrasjon", "un.SY.2NRD_N": "Nordics", "un.SY.2CWK": "CA, WS & Kystradio",
    "un.SY.3KYS_C": "Kystradio", "un.SY.3FIB": "TnFiber IC", "un.SY.5OMS_A": "Omstilling",
    "un.SY.2I15": "IFRS15-justering", "un.SY.2I16": "IFRS16-justering", "un.SY.E171": "IFRS16 elim. korr.",
    "un.SY.DE11": "STAR", "un.SY.G916": "IMEI korr.", "un.SY.SYi": "Vurdering", "un.SY.2CON": "Consumer (nedlagt)",
    "un.SY.2IG": "Infrastructure Growth (nedlagt)", "un.SY.OUTDATED": "Utgåtte noder",
}
ADJUST = {"un.SY.2I15", "un.SY.2I16", "un.SY.E171", "un.SY.DE11", "un.SY.G916", "un.SY.SYi", "un.SY.5OMS_A"}

# ------------------------------------------------------------------ betydning (utkast, til gjennomgang)
M = {
    "C85": "Resultat før skatt for Telenor Norge.",
    "C84S": "Netto finansposter: renter, valuta, verdiendringer og gevinst/tap på finansielle eiendeler.",
    "C84": "Gevinst/tap ved salg og nedskrivning av finansielle eiendeler (aksjer, fordringer).",
    "C83": "Verdiendring på finansielle instrumenter og sikringsposter (hedge accounting).",
    "C82": "Netto valutagevinst/-tap (realisert og urealisert), mot konsernbank og eksternt.",
    "C81": "Sum finanskostnader.",
    "C815": "Aktiverte renter (reduserer finanskostnaden; rente henført til anlegg under utførelse).",
    "C813": "Finanskostnader: rentekostnader, leierenter (IFRS 16), ARO-renter, pensjonsrente og tap på kunders avbetalingsfordringer.",
    "C81315": "Konstaterte tap på kunders avbetalingsfordringer (delbetaling av utstyr, f.eks. Splitt).",
    "C81314": "Endring i avsetning for tap på kunders avbetalingsplaner.",
    "C81313": "Inngått på tidligere avskrevne avbetalingsfordringer.",
    "C81340": "Rentekostnad på leieforpliktelser (IFRS 16).",
    "C81300": "Rentekostnader på lån og annen rentebærende gjeld.",
    "C81370": "Renteeffekt på forpliktelser for fjerning/opprydding av anlegg (ARO).",
    "C81385": "Netto rentekostnad på ytelsesbasert pensjon.",
    "C81390": "Andre finanskostnader (gebyrer, provisjoner o.l.).",
    "C80": "Sum finansinntekter: renter, utbytte og forsinkelsesrenter.",
    "C80311": "Renteinntekt på kunders avbetalingsplaner (delbetaling av utstyr, f.eks. Splitt).",
    "C80300": "Renteinntekter på bankinnskudd og plasseringer.",
    "C80350": "Forsinkelsesrenter fra kunder.",
    "C80330": "Renteinntekt på forfalte kundefordringer.",
    "C80100": "Utbytte fra datterselskaper.",
    "C77": "Andel resultat i tilknyttede selskaper og felleskontrollert virksomhet (egenkapitalmetoden).",
    "EBIT": "Driftsresultat (EBITDA minus av- og nedskrivninger).",
    "C76S": "Sum av- og nedskrivninger.",
    "C76": "Nedskrivninger og reverseringer (varige, immaterielle, bruksrett, goodwill).",
    "C75": "Ordinære avskrivninger: varige driftsmidler, immaterielle eiendeler, bruksrett (IFRS 16) og merverdier.",
    "C75110": "Avskrivning varige driftsmidler (nett, utstyr, bygg).",
    "C75120": "Avskrivning immaterielle eiendeler (programvare, lisenser, frekvenser).",
    "C75130": "Avskrivning bruksrettseiendeler (leie etter IFRS 16).",
    "EBITDA": "Resultat før av- og nedskrivninger, finans og skatt (inkl. andre poster C73).",
    "C73": "Andre inntekter og kostnader («other items»): gevinst/tap på salg og utrangering av anlegg, nedbemanning, tapskontrakter, engangseffekter.",
    "C731": "Gevinst/tap på salg og utrangering av varige/immaterielle eiendeler og bruksrett.",
    "C733": "Gevinst/tap på salg av aksjer i datterselskaper.",
    "C735": "Nedbemanning og tapsbringende kontrakter.",
    "C736": "Andre engangseffekter (bl.a. pensjon).",
    "C737": "Andre inntekter utenfor ordinær drift.",
    "EBITDA_BOI": "EBITDA før andre poster: den operasjonelle EBITDA (C48 + C70).",
    "C70": "Sum driftskostnader: lønn og personal (C50S) + andre driftskostnader (C60).",
    "C60": "Andre driftskostnader: salg og markedsføring, drift og vedlikehold, konsulenter, lokaler, energi og øvrig.",
    "C609": "Øvrige driftskostnader (konsulenter og innleie, lokaler, energi, reise, frakt, tap på fordringer, interne tjenester).",
    "C693": "Konsulenter og innleid arbeidskraft (IT- og andre konsulenter).",
    "C694": "Andre kostnader: lokaler, energi, reise, porto/frakt, lisenser, tap på fordringer, management fee, tjenester fra delt tjenestesenter.",
    "C60300": "Avsetning for tap på kundefordringer.",
    "C60190": "Konstaterte tap på kundefordringer.",
    "C60200": "Inngått på tidligere avskrevne fordringer.",
    "C60180": "Lisensavgifter og gebyrer for tjenester.",
    "C60170": "Management fee (konserninterne tjenester).",
    "C60910": "Tjenester kjøpt fra delt tjenestesenter og andre forretningsenheter i Telenor.",
    "C60920": "Investeringer/avskrivninger fakturert fra delt tjenestesenter og andre forretningsenheter.",
    "C60930": "Margin på tjenester fra delt tjenestesenter og andre forretningsenheter.",
    "C60700": "Abonnementer, gaver, forsikringer o.l.",
    "C60120": "Porto, frakt og telekommunikasjon.",
    "C60130": "Reise og diett.",
    "C60025": "Energikostnad (strøm til nett og lokaler).",
    "C60020": "Andre kostnader til bygg og eiendom (drift, vedlikehold).",
    "C60010": "Leie av lokaler og grunn (utenfor IFRS 16).",
    "C601": "Salg, markedsføring og provisjoner.",
    "C60160": "Reklame.",
    "C60150": "Salgs-, markedsførings- og representasjonskostnader.",
    "C60140": "Provisjoner (forhandler- og salgsprovisjoner).",
    "C60145": "Amortisering av kostnader for å skaffe og oppfylle kundekontrakter (IFRS 15), driftskostnad.",
    "C600": "Drift og vedlikehold.",
    "C60050": "Drift og vedlikehold IS/IT.",
    "C60060": "Drift og vedlikehold fra entreprenører (nett).",
    "C60065": "Drift og vedlikehold fra delt tjenestesenter og andre forretningsenheter.",
    "C60070": "Annen drift og vedlikehold.",
    "C50S": "Sum lønn og personalkostnader, fratrukket aktivert eget arbeid.",
    "C49200": "Aktivert eget arbeid (lønn) – reduserer personalkostnad.",
    "C50": "Lønn og personalkostnader.",
    "C50100": "Lønn og feriepenger.",
    "C50200": "Bonus.",
    "C50300": "Arbeidsgiveravgift.",
    "C50500": "Innskuddspensjon inkl. arbeidsgiveravgift.",
    "C50600": "Ytelsespensjon inkl. arbeidsgiveravgift.",
    "C50700": "Aksjebasert avlønning.",
    "C50900": "Andre personalkostnader.",
    "C48": "Bidrag: inntekter (C30) + vare- og trafikkostnader (C40S).",
    "C40S": "Sum vare- og trafikkostnader (varekost), fratrukket aktivert eget arbeid.",
    "C40": "Vare- og trafikkostnader: varer for videresalg, leid og kjøpt nettkapasitet, trafikkostnader, lagerendring.",
    "C40510": "Forbruk av varer for videresalg (håndsett, utstyr, CPE) og råvarer.",
    "C40110": "Leid nettkapasitet nasjonalt, stamnett.",
    "C40130": "Leid nettkapasitet nasjonalt, aksessnett / siste mil.",
    "C40150": "Leid nettkapasitet internasjonalt.",
    "C40310": "Trafikkostnader nasjonalt (samtrafikk).",
    "C40330": "Trafikkostnader internasjonalt (samtrafikk, roaming).",
    "C40610": "Endring i beholdning av varer i arbeid og ferdigvarer.",
    "C40685": "Amortisering av kontraktskostnader (IFRS 15), varekost.",
    "C40700": "Annen varekost.",
    "C30": "Sum inntekter.",
    "C30001": "Inntekter fra kundekontrakter (IFRS 15).",
    "30": "Varesalg (håndsett, utstyr, tilbehør).",
    "31": "Abonnement og andre avtaleinntekter (faste månedsavgifter).",
    "32": "Trafikkinntekter (bruk: tale, data, SMS, roaming).",
    "35": "Innholdstjenester (TV-innhold, strømming, tredjepartstjenester).",
    "36": "Inntektsformidling (benyttes kun av Networks).",
    "37": "Leieinntekter (utleie av utstyr, lokaler, infrastruktur).",
    "39": "Annen driftsrelatert inntekt.",
    "C30002": "Leieinntekter etter IFRS 16 (utleier).",
    "C30003": "Inntekter fra finansielle instrumenter.",
    "C30004": "Andre inntekter.",
}
P = {   # fallgruver / hva som ikke ligger i FinMax
    "C84S": "Finansposter ligger hovedsakelig i Administrasjon/Finance og uten produkt. Avbetalingsrenter (C80311) ligger i SSM/5SAL_C.",
    "C813": "Tap på avbetalingsfordringer (C81313–C81315) og finansieringskostnad for håndsett var ikke synlige i un.SY i 2025 – kan ligge i annen juridisk enhet.",
    "C80311": "Bare delvis produktmerket. Splitt-definisjonen henfører hele beløpet til Splitt.",
    "C76S": "Mest på koststed (Technology/IT); bare delvis produktmerket. Full produktfordeling krever Produktregnskapet.",
    "C75": "Mest på koststed (Technology/IT), bare delvis produktmerket. FinPlan Target har ikke avskrivninger – sammenlign mot budsjett bare til EBITDA.",
    "C73": "Holdes utenfor EBITDA før andre poster. FinPlan Target har ikke andre poster.",
    "C40510": "Håndsett-varekost ligger på produkt (bl.a. M5008/M5010), men ikke per tilbud (f.eks. Splitt) – fordel med nøkkel.",
    "C60140": "Kan ligge på produkt, f.eks. Splitt-finanskomponenten FV081.",
    "C60145": "IFRS 15-kontraktskostnader, amortiseres. Delvis i IFRS15-enheten 2I15.",
    "C40685": "IFRS 15-kontraktskostnader, amortiseres. Delvis i IFRS15-enheten 2I15.",
    "C75130": "IFRS 16 – delvis i justeringsenhetene 2I16/E171.",
    "C81340": "IFRS 16 – delvis i justeringsenhetene 2I16/E171.",
    "C30": "Historikken er omarbeidet til 2026-organisasjonen (reorg 1. juni 2026).",
    "C60910": "Interne kjøp fra andre Telenor-enheter – ikke ekstern kostnad.",
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


def groups():
    db = sqlite3.connect(ATLAS)
    q = "select name from member where connection='FinMax' and dimension='All Accounts' and parent=? order by ord"
    lab = lambda n: (db.execute("select label from member where connection='FinMax' and dimension='All Accounts' "
                                "and name=? limit 1", (n,)).fetchone() or [n])[0].replace(" (ac)", "")
    out = []

    def leaves(n):
        k = db.execute(q, (n,)).fetchall()
        return 1 if not k else sum(leaves(x[0]) for x in k)

    def walk(n, parent, depth):
        k = db.execute(q, (n,)).fetchall()
        if not k:
            return
        label = lab(n)
        mirror = n.endswith("i") and any(ch.isdigit() for ch in n[-3:])
        if not mirror and "plug" not in label.lower() and "elimination" not in label.lower():
            leaf_list = []
            if all(not db.execute(q, (x[0],)).fetchall() for x in k):
                leaf_list = [x[0] for x in k]
            out.append({"name": n, "label": label, "parent": parent, "depth": depth, "n_leaves": leaves(n),
                        "leaves": leaf_list[:60]})
        for x in k:
            walk(x[0], n, depth + 1)
    walk("ac.C85", None, 0)
    return out


def pct(a, b):
    t = abs(a) + abs(b)
    return None if t < 1 else abs(a) / t


def main():
    div, f_div = read_pull(latest("map_accounts_by_division_FY25"))
    prod, f_prod = read_pull(latest("map_accounts_product_tag"))
    res, f_res = read_pull(latest("map_accounts_reserve_tag"))
    rows, js = [], []
    for g in groups():
        n, code = g["name"], g["name"][3:]
        d = div.get(n, {})
        total = d.get("un.SY", 0.0)
        ytd = prod.get(n, {})
        ytd26 = ytd.get("pd.All_Mapped_Produc / Sep 26", 0.0) + ytd.get("pd.No_Product / Sep 26", 0.0)
        p_share = pct(ytd.get("pd.All_Mapped_Produc / Dec 25", 0.0), ytd.get("pd.No_Product / Dec 25", 0.0))
        r = res.get(n, {})
        r_share = pct(r.get("re.All_Mapped_Reserv", 0.0), r.get("re.No_Reserve", 0.0))
        parts = {k: d.get(k, 0.0) for k in DIVISIONS if abs(d.get(k, 0.0)) >= 1}
        tot_abs = sum(abs(v) for v in parts.values())
        top = max(parts, key=lambda k: abs(parts[k])) if parts else None
        top_share = abs(parts[top]) / tot_abs if top and tot_abs else None
        adj_share = sum(abs(v) for k, v in parts.items() if k in ADJUST) / tot_abs if tot_abs else 0
        pattern = []
        if abs(total) < 1 and tot_abs < 1:
            pattern.append("Ingen bevegelse i Telenor Norge (un.SY) i 2025")
        else:
            if p_share is not None:
                pattern.append("Produktmerket" if p_share >= 0.9 else
                               f"Delvis produktmerket ({p_share:.0%})" if p_share >= 0.1 else "Uten produkt (koststed)")
            if r_share and r_share >= 0.1:
                pattern.append(f"reserve-merket {r_share:.0%}")
            if top:
                pattern.append(f"mest i {DIVISIONS[top]} ({top_share:.0%})")
            if adj_share > 0.3:
                pattern.append(f"{adj_share:.0%} i justeringsenheter")
        pit = [P[code]] if code in P else []
        if abs(total) < 1 and tot_abs < 1:
            pit.append("Ikke bokført i un.SY i 2025 – kan ligge i annen enhet eller være ubrukt.")
        meaning = M.get(code) or f"Del av {g['parent'][3:] if g['parent'] else ''}: {g['label'].split(' ', 1)[-1]}."
        parent = (g["parent"] or "")[3:]
        rec = {"code": code, "member": n, "label": g["label"], "parent": parent, "depth": g["depth"],
               "n_posting_accounts": g["n_leaves"], "fy2025_nok": round(total, 2), "ytd_sep26_nok": round(ytd26, 2),
               "top_division": DIVISIONS.get(top) if top else None, "top_division_share": top_share,
               "product_tagged_share": p_share, "reserve_tagged_share": r_share,
               "booking_pattern": "; ".join(pattern), "meaning": meaning, "pitfalls": " ".join(pit),
               "divisions_fy2025": {DIVISIONS[k]: round(v, 2) for k, v in parts.items()},
               "posting_accounts": g["leaves"]}
        js.append(rec)
        rows.append(rec)
    status, review_label = review.apply("accounts", rows)
    meta = {"built": dt.datetime.now().isoformat(timespec="seconds"), "status": status, "review_label": review_label,
            "sources": [f_div, f_prod, f_res, "atlas.db (FinMax All Accounts)"],
            "notes": ["Funksjonsdimensjonen (All Functions) er ubrukt i un.SY: alt ligger på No_Function.",
                      "Beløp: FinMax Actual NOK, un.SY = Telenor Norge, FY2025 = Dec 25 Accumulated.",
                      "Andel produktmerket = |pd.All_Mapped_Produc| / (|mapped| + |pd.No_Product|), FY2025."]}
    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "accounts": js}, f, ensure_ascii=False, indent=1)
    print(f"{len(rows)} kontogrupper -> {JSON_OUT}")
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
        ws_info = wb.Worksheets(1)
        ws_info.Name = "Les meg"
        ws = wb.Worksheets.Add(After=ws_info)
        ws.Name = "Konti"
        ws_l = wb.Worksheets.Add(After=ws)
        ws_l.Name = "Lister"

        # Lister
        ws_l.Range("B4:B6").Value = (("OK",), ("Endre",), ("Usikker",))
        T.name(wb, "Liste_Vurdering", "=Lister!$B$4:$B$6")
        ws_l.Visible = 0

        # Les meg
        T.setup_sheet(ws_info, "Forretningskart 1 – Konti og bokføring",
                      f"FinMax · Telenor Norge (un.SY) · FY2025 · {meta['status']} · bygget {meta['built'][:16].replace('T', ' ')}")
        lines = [
            "Hva er dette: hva hver kontogruppe i FinMax betyr, hvor den normalt bokføres (produkt, koststed, reserve,",
            "divisjon) og fallgruver – slik at smartview vet hvilke poster et spørsmål trenger.",
            "",
            "Slik går du gjennom: arket «Konti». Les «Betydning» og «Fallgruver», velg OK / Endre / Usikker i",
            "«Vurdering», og skriv rettelser i «Kommentar». Tall er fakta fra FinMax og skal ikke endres.",
            "",
            "Kolonner: Bokføringsmønster = produktmerket andel, reserve-merket andel og største divisjon (FY2025).",
            "Funksjonsdimensjonen er ubrukt i Telenor Norge (alt på No_Function).",
            "",
            "Kilder: " + "; ".join(meta["sources"]),
            "Bygget av smartass\\knowledge\\build_accounts_map.py – kjøres på nytt etter månedlig atlas-oppdatering.",
        ]
        for i, t in enumerate(lines):
            c = ws_info.Cells(4 + i, 2)
            c.Value = "'" + t if t else None
            c.Style = "TN Tekst"
        ws_info.Columns("B").ColumnWidth = 110

        # Konti
        T.setup_sheet(ws, "Konti og bokføring – FinMax", f"{meta['review_label']} · beløp i NOK, Telenor Norge (un.SY)",
                      freeze_row=4, freeze_col=3)
        headers = ["Kode", "Navn (FinMax)", "Overordnet", "Nivå", "Antall posteringskonti", "FY2025 (NOK)",
                   "Hittil sep 26 (NOK)", "Produktmerket andel", "Reserve-merket andel", "Største divisjon",
                   "Andel største divisjon", "Bokføringsmønster", "Betydning", "Fallgruver / ikke i FinMax",
                   "Vurdering", "Kommentar"]
        data = []
        for r in rows:
            data.append(["'" + r["code"], "'" + r["label"], "'" + r["parent"], r["depth"], r["n_posting_accounts"],
                         r["fy2025_nok"], r["ytd_sep26_nok"], r["product_tagged_share"], r["reserve_tagged_share"],
                         r["top_division"], r["top_division_share"], r["booking_pattern"], r["meaning"],
                         r["pitfalls"] or None, None, None])
        xl.AutoCorrect.AutoFillFormulasInLists = False
        lo = T.add_table(ws, "B4", headers, data, "tblKonti")
        for name_, style, w in (("Kode", "TN Tabell ID", 11), ("Navn (FinMax)", "TN Tabell tekst", 38),
                                ("Overordnet", "TN Tabell kort", 11), ("Nivå", "TN Tabell heltall", 6),
                                ("Antall posteringskonti", "TN Tabell heltall", 11),
                                ("FY2025 (NOK)", "TN Tabell heltall", 16), ("Hittil sep 26 (NOK)", "TN Tabell heltall", 16),
                                ("Produktmerket andel", "TN Tabell prosent", 11),
                                ("Reserve-merket andel", "TN Tabell prosent", 11),
                                ("Største divisjon", "TN Tabell kort", 22), ("Andel største divisjon", "TN Tabell prosent", 11),
                                ("Bokføringsmønster", "TN Tabell tekst", 34), ("Betydning", "TN Tabell tekst", 50),
                                ("Fallgruver / ikke i FinMax", "TN Tabell tekst", 40),
                                ("Vurdering", "TN Status", 11), ("Kommentar", "TN Tabell input", 36)):
            T.style_col(lo, name_, style, w)
        T.list_validation(T.col(lo, "Vurdering"), "Liste_Vurdering")
        T.status_colors(T.col(lo, "Vurdering"), {"OK": "ferdig", "Endre": "venter", "Usikker": "stopp"})
        T.page_setup(ws)
        ws_info.Activate()
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
