"""One-off converter (kept for reruns): v1's owner-approved catalogue (catalog.toml, capabilities.toml) and business
maps (accounts_map.json, org_map.json, finplan_map.json, all approved by Tom on 2026-10-15) into the semantic layer's
TOML files. `seed(knowledge_dir, out_dir)` (re)writes the files it generates and leaves every other file alone.

Metrics join the catalogue's [pnl] lines across cubes on `concept` (never on `key`), as interim.py did. Facts carry
the maps' pitfalls in English (fact text is sheet and reply text)."""
from __future__ import annotations

import json
import re
import tomllib
from collections import defaultdict
from datetime import date
from pathlib import Path

from .semantic import Concept, Fact, Metric, to_toml

SIGN = {"revenue": "income", "gross_margin": "income", "ebitda_boi": "income", "ebitda": "income",
        "cogs": "cost", "personnel": "cost", "other_opex": "cost", "opex": "cost", "da": "cost"}
APPROVED = {"by": "tom", "on": "2026-10-15", "via": "catalog.toml"}
REORG = "2026-06"
DIFFERS = ("FinPlan groups accounts differently from FinMax (e.g. EBITDA 2025: 14.16 bn FinPlan vs 12.40 bn FinMax); "
           "compare across cubes only with a reconciliation line.")
BOILERPLATE = "Ikke bokført i un.SY i 2025"

# accounts_map.json pitfalls (Norwegian, approved 2026-10-15) -> English fact text. A new pitfall stops the seed
# until it has a translation here.
ACCOUNT_PITFALLS = {
    "Finansposter ligger hovedsakelig i Administrasjon/Finance og uten produkt. Avbetalingsrenter (C80311) ligger i "
    "SSM/5SAL_C.":
        "Financial items sit mainly in Administration/Finance, with no product. Instalment interest (C80311) sits in "
        "SSM/5SAL_C.",
    "Tap på avbetalingsfordringer (C81313–C81315) og finansieringskostnad for håndsett var ikke synlige i un.SY i "
    "2025 – kan ligge i annen juridisk enhet.":
        "Losses on instalment receivables (C81313–C81315) and the handset financing cost were not visible in un.SY "
        "in 2025; they may sit in another legal entity.",
    "IFRS 16 – delvis i justeringsenhetene 2I16/E171.":
        "IFRS 16: partly booked in the adjustment units 2I16/E171.",
    "Bare delvis produktmerket. Splitt-definisjonen henfører hele beløpet til Splitt.":
        "Only partly product-tagged. The Splitt definition attributes the whole amount to Splitt.",
    "Mest på koststed (Technology/IT); bare delvis produktmerket. Full produktfordeling krever Produktregnskapet.":
        "Mostly booked on cost centres (Technology/IT); only partly product-tagged. A full product split needs the "
        "product cost accounts (Produktregnskapet).",
    "Mest på koststed (Technology/IT), bare delvis produktmerket. FinPlan Target har ikke avskrivninger – "
    "sammenlign mot budsjett bare til EBITDA.":
        "Mostly booked on cost centres (Technology/IT); only partly product-tagged. FinPlan Target has no "
        "depreciation: compare against budget only down to EBITDA.",
    "Holdes utenfor EBITDA før andre poster. FinPlan Target har ikke andre poster.":
        "Kept outside EBITDA before other items. FinPlan Target has no other items.",
    "Interne kjøp fra andre Telenor-enheter – ikke ekstern kostnad.":
        "Internal purchases from other Telenor units, not an external cost.",
    "Kan ligge på produkt, f.eks. Splitt-finanskomponenten FV081.":
        "May be booked on a product, e.g. the Splitt finance component FV081.",
    "IFRS 15-kontraktskostnader, amortiseres. Delvis i IFRS15-enheten 2I15.":
        "IFRS 15 contract costs, amortised. Partly booked in the IFRS 15 unit 2I15.",
    "Håndsett-varekost ligger på produkt (bl.a. M5008/M5010), men ikke per tilbud (f.eks. Splitt) – fordel med "
    "nøkkel.":
        "Handset COGS is booked on products (e.g. M5008/M5010) but not per offer (e.g. Splitt): allocate it with a "
        "key.",
    "Historikken er omarbeidet til 2026-organisasjonen (reorg 1. juni 2026).":
        "History is restated to the 2026 organisation (reorg 1 Jun 2026).",
}

# org_map.json pitfall patterns -> (fact id, English text)
ORG_PATTERNS = {
    "Inntektsbærende koststed": (
        "org_revenue_bearing_cost_centres",
        "Revenue-bearing cost centre: revenue is booked on the cost centre itself, so its net figures mix revenue "
        "and cost."),
    "Justerings-/korreksjonsenhet": (
        "org_adjustment_units",
        "Adjustment or correction unit (IFRS, elimination, valuation), not an operating unit."),
}

# finplan_map.json pitfall patterns (from its meta notes) -> (fact id, English text, severity)
FINPLAN_PATTERNS = {
    "Les som Periodic": (
        "finplan_kpis_periodic",
        "FinPlan volumes, ARPU and percentages read Periodic per month: Accumulated sums the months for these KPIs "
        "(e.g. Mobile Subscriptions YTD Sep = 78 million against 8.7 million in stock).", "warn"),
    "Bare i plan/prognose": (
        "finplan_no_actual_volumes",
        "Actual volumes (subscriptions, customers) are not in FinPlan for BANORWAY: only Target and forecasts carry "
        "them (Actual is empty).", "warn"),
    "Target (budsjett) har ingen verdi her": (
        "finplan_target_stops_at_ebitda",
        "FinPlan Target carries no depreciation and no other items, so budget comparisons stop at EBITDA (before or "
        "after other items).", "warn"),
}
MAPS_SOURCE = "{} (approved by Tom 2026-10-15)"


def _read(knowledge_dir: Path, name: str):
    p = knowledge_dir / name
    text = p.read_text(encoding="utf-8")
    return tomllib.loads(text) if name.endswith(".toml") else json.loads(text)


def _dedupe(words) -> list[str]:
    seen, out = set(), []
    for w in words:
        if w and w.casefold() not in seen:
            seen.add(w.casefold())
            out.append(w)
    return out


def _slug(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")


def seed_metrics(catalog: dict, capabilities: dict) -> list[Metric]:
    mat = catalog["materiality"]
    materiality = (float(mat["abs_nok"]), float(mat["pct"]))
    by_concept: dict[str, dict] = {}
    for cube in ("FinMax", "FinPlan"):
        for line in catalog["pnl"][cube]:
            concept = line.get("concept")
            if not concept:
                continue
            d = by_concept.setdefault(concept, {"label": line["label"], "members": {}, "stock": False})
            d["members"][cube] = [line["account"]]
            d["stock"] = d["stock"] or line.get("measure") == "stock"
            if cube == "FinMax":
                d["label"] = line["label"]
    measures = capabilities.get("measure", [])
    out = []
    for concept, d in by_concept.items():
        stock = d["stock"]
        members = d["members"]
        if "FinMax" in members and not stock:      # FinLite = the same accounts as FinMax; stock (FTE) is not in it
            members["FinLite"] = members["FinMax"]
        mine = [m for m in measures if m.get("concept") == concept]
        own = next((m for m in mine if m["key"] == concept), None)
        aliases = _dedupe([s for m in mine for s in m.get("synonyms", [])]
                          + [a.removeprefix("ac.") for a in members.get("FinMax", [])]
                          + members.get("FinPlan", []))
        out.append(Metric(
            id=concept, name=d["label"], aliases=aliases,
            description=(own or {}).get("note") or _identity(catalog, members.get("FinMax"), d["label"]),
            unit="FTE" if stock else "NOK", sign=SIGN.get(concept, "neutral"),
            balance="Accumulated" if stock else "Periodic", aggregation="last" if stock else "sum",
            materiality=materiality, status="approved", approved=dict(APPROVED),
            finmax={"members": members["FinMax"]} if "FinMax" in members else None,
            finlite={"members": members["FinLite"]} if "FinLite" in members else None,
            finplan={"members": members["FinPlan"]} if "FinPlan" in members else None,
            differs=DIFFERS if "FinPlan" in members else ""))
    return out


def _day(iso: str) -> str:
    """'2026-06-01' -> '1 Jun 2026'."""
    d = date.fromisoformat(iso)
    return f"{d.day} {d:%b %Y}"


def _identity(catalog: dict, finmax: list[str] | None, label: str) -> str:
    """The label, plus the FinMax identity that defines a sum line (catalog [checks], e.g. 'C50S + C60 = C70')."""
    code = finmax[0].removeprefix("ac.") if finmax else None
    for check in catalog.get("checks", {}).get("FinMax", []):
        if code and check["label"].endswith(f"= {code}"):
            return f"{label}: {check['label']}"
    return label


def _finmax_scope(members: list[str]) -> dict[str, list[str]]:
    """FinLite shares FinMax's entity and account members."""
    return {"FinMax": list(members), "FinLite": list(members)}


def seed_facts(catalog: dict, accounts_map: dict, org_map: dict, finplan_map: dict) -> list[Fact]:
    facts: list[Fact] = []
    entities = {e["key"]: e for e in catalog.get("entity", [])}

    # 1. abolished business areas
    for e in entities.values():
        if not e.get("abolished"):
            continue
        succ = ", ".join(f"{entities[s]['label']} ({entities[s]['finmax']})" for s in e.get("successors", []))
        when = e["abolished"]
        scope = _finmax_scope([e["finmax"]])
        if e.get("finplan"):
            scope["FinPlan"] = [e["finplan"]]
        facts.append(Fact(
            id=f"abolished_{e['key']}",
            text=f"{e['label']} was abolished on {_day(when)}; successors {succ}; history is restated to today's "
                 "organisation.",
            scope_members=scope, valid_from=when[:7], severity="warn",
            source=f"catalog.toml [[entity]] {e['key']} (approved by Tom 2026-10-15)", status="approved"))

    # 2. reorg moves, one fact per move pair
    moves: dict[str, list[str]] = defaultdict(list)
    for row in org_map["entities"]:
        if row.get("moved"):
            moves[row["moved"]].append(row["member"])
    for pair, members in sorted(moves.items()):
        a, b = (s.strip() for s in pair.split("→"))
        facts.append(Fact(
            id=f"reorg_{_slug(a)}_{_slug(b)}",
            text=f"Moved in the 1 Jun 2026 reorg ({a} → {b}); history is restated to the new organisation.",
            scope_members=_finmax_scope(members), valid_from=REORG, severity="info",
            source=MAPS_SOURCE.format(f"org_map.json moved '{pair}'"), status="approved"))

    # 3. account traps
    for acc in accounts_map["accounts"]:
        pit = acc.get("pitfalls") or ""
        if not pit or pit.startswith(BOILERPLATE):
            continue
        if pit not in ACCOUNT_PITFALLS:
            raise ValueError(f"accounts_map.json {acc['member']}: no English text for the pitfall {pit!r}; add it to "
                             "seed.ACCOUNT_PITFALLS")
        facts.append(Fact(
            id=f"account_{_slug(acc['code'])}", text=f"{acc['label']}: {ACCOUNT_PITFALLS[pit]}",
            scope_members=_finmax_scope([acc["member"]]), severity="warn",
            source=MAPS_SOURCE.format(f"accounts_map.json {acc['member']} pitfalls"), status="approved"))

    # 4. revenue-bearing cost centres and adjustment units
    for pattern, (fid, text) in ORG_PATTERNS.items():
        members = [r["member"] for r in org_map["entities"] if pattern in (r.get("pitfalls") or "")]
        if members:
            facts.append(Fact(id=fid, text=text, scope_members=_finmax_scope(members), severity="info",
                              source=MAPS_SOURCE.format(f"org_map.json pitfalls '{pattern}'"), status="approved"))

    # 5. FinPlan facts (finplan_map meta notes)
    facts.append(Fact(
        id="finplan_banorway",
        text="FinPlan entity BANORWAY is Telenor Norge (FinMax un.SY). Actual is Scenario Actual, version Working; "
             "Target and Forecast Q<q><yy> are version Final.",
        scope_members={"FinPlan": ["BANORWAY"]}, severity="info",
        source=MAPS_SOURCE.format("finplan_map.json meta notes"), status="approved"))
    for pattern, (fid, text, severity) in FINPLAN_PATTERNS.items():
        members = [r["member"] for r in finplan_map["accounts"] if pattern in (r.get("pitfalls") or "")]
        if members:
            facts.append(Fact(id=fid, text=text, scope_members={"FinPlan": members}, severity=severity,
                              source=MAPS_SOURCE.format(f"finplan_map.json meta notes + pitfalls '{pattern}'"),
                              status="approved"))
    return facts


# ---- product concepts (Plan 2C): knowledge/concepts.json -> concepts/<id>.toml

# concepts.json texts (Norwegian; splitt agreed with Tom 14.10, the business map with the families approved 15.10,
# streamix 22.10) -> English sheet text. A new text stops the seed until it has a translation here.
CONCEPT_TEXT = {
    "Telenor Consumers tilbud om å kjøpe håndsett/utstyr på delbetaling, med valgfri Plussforsikring og egne "
    "Splitt-abonnement. Lønnsomhet = direkte bidrag fra FinMax.":
        "Telenor Consumer's offer to buy handsets and equipment on instalments, with optional Pluss insurance "
        "(Plussforsikring) and dedicated Splitt subscriptions. Profitability = direct contribution from FinMax.",
    "Lønnsomhet = direkte bidrag fra FinMax (ikke fullfordelt Produktregnskap).":
        "Profitability = direct contribution from FinMax (not the fully allocated product cost accounts, "
        "Produktregnskapet).",
    "Merking: (bokført) = merket på Splitt-produkt/koststed; (estimert) = delt beløp × nøkkel; (henført 100 %) = "
    "ikke merket, men tilhører Splitt.":
        "Tagging: (booked) = tagged on a Splitt product or cost centre; (estimated) = shared amount × key; "
        "(attributed) = not tagged, but belongs to Splitt in full.",
    "Øvrige kontoer bokført på Splitt-produktene (tap på fordringer, frakt, markedsføring, rabatter, gebyrer) kommer "
    "automatisk med i C40S/C70-linjene.":
        "Other accounts booked on the Splitt products (losses on receivables, freight, marketing, discounts, fees) "
        "are included automatically in the C40S/C70 lines.",
    "Sammenligning mot i fjor: provisjon på FV081 (C60140) og renteinntekt C80311 ble i 2025 først bokført i 4. "
    "kvartal – hittil-tall mot hittil i fjor gir derfor store endringer.":
        "Against last year: commission on FV081 (C60140) and interest income C80311 were first booked in Q4 2025, "
        "so year to date against last year's year to date shows large changes.",
    "Plussforsikring (F4046) økte kraftig fra 2025 til 2026 – nøkkelen for forsikringsinnholdskostnad følger med "
    "(andel av forsikringsinntekt).":
        "Pluss insurance (F4046) grew strongly from 2025 to 2026; the key for the insurance content cost follows it "
        "(share of insurance revenue).",
    "Innbytte fordeles med håndsettnøkkelen (Splitts andel av håndsettinntekt), fordi innbytte gjelder alt "
    "håndsettsalg. Valgt 15.10.2026 uten bedre kilde; uvesentlig: innbytte totalt er −3,5 MNOK hittil sep 26, så 0 "
    "% eller 100 % flytter bidraget med under 3 %.":
        "Trade-in is allocated with the handset key (Splitt's share of handset revenue), because trade-in concerns "
        "all handset sales. Chosen 15.10.2026 for want of a better source; immaterial: trade-in in total is "
        "-3.5 MNOK YTD Sep 26, so 0 % or 100 % moves the contribution by less than 3 %.",
    "Direkte bidrag Splitt (før finansieringskostnad og tap)":
        "Direct contribution Splitt (before financing cost and losses)",
    "Maskinvarelinjene = M4517 Handset Splitt Revenue": "Hardware lines = M4517 Handset Splitt Revenue",
    "Maskinvare (Splitt-salg av utstyr)": "Hardware (Splitt equipment sales)",
    "Splitt-inntekt Apple": "Splitt revenue Apple",
    "Splitt-inntekt Samsung": "Splitt revenue Samsung",
    "Splitt-inntekt øvrige merker": "Splitt revenue other brands",
    "God Som Ny (brukte enheter)": "God Som Ny (used devices)",
    "Finanskomponent (renteandel skilt ut av inntekten)": "Finance component (interest share split out of revenue)",
    "Maskinvareinntekt Splitt": "Hardware revenue Splitt",
    "Varekost maskinvare": "Hardware COGS",
    "Splitts andel av all håndsettinntekt": "Splitt's share of all handset revenue",
    "Håndsett-varekost totalt (M5008 + M5010)": "Handset COGS in total (M5008 + M5010)",
    "Splitt maskinvareinntekt ekskl. finanskomponent": "Splitt hardware revenue excl. the finance component",
    "Håndsettinntekt totalt (M411)": "Handset revenue in total (M411)",
    "Andre kostnader bokført på Splitt-produktene": "Other costs booked on the Splitt products",
    "Varekost bokført på Splitt (C40S)": "COGS booked on Splitt (C40S)",
    "Driftskostnader bokført på Splitt (C70, bl.a. provisjon C60140)":
        "Opex booked on Splitt (C70, incl. commission C60140)",
    "Finansiering": "Financing",
    "Renteinntekt på avbetalingsplaner (C80311)": "Interest income on instalment plans (C80311)",
    "Ikke produktmerket; henført 100 % til Splitt.": "Not product-tagged; attributed 100 % to Splitt.",
    "Finansieringskostnad og tap på avbetalingsfordringer": "Financing cost and losses on instalment receivables",
    "Ikke synlig i FinMax un.SY (C81313–C81315, 81502) – ikke med i bidraget.":
        "Not visible in FinMax un.SY (C81313–C81315, 81502); not in the contribution.",
    "Forsikring": "Insurance",
    "Splitt Plussforsikring": "Splitt Pluss insurance",
    "Forsikringsinnholdskostnad": "Insurance content cost",
    "Splitts andel av all forsikringsinntekt": "Splitt's share of all insurance revenue",
    "Forsikringsinnholdskostnad totalt (F4010)": "Insurance content cost in total (F4010)",
    "Forsikringsinntekt totalt (M4124)": "Insurance revenue in total (M4124)",
    "Innbytte": "Trade-in",
    "Innbytte (innbytteverdi og -bonus, C120–C134)": "Trade-in (trade-in value and bonus, C120–C134)",
    "samme nøkkel som maskinvaren (andel av håndsettinntekt) – innbytte gjelder alt håndsettsalg":
        "the same key as the hardware (share of handset revenue): trade-in concerns all handset sales",
    "Innbytte totalt (EBITDA før andre poster, C120–C134)":
        "Trade-in in total (EBITDA before other items, C120–C134)",
    "Splitt-abonnement og datapakker": "Splitt subscriptions and data packs",
    "Splitt-tariffer og datapakker (bidrag C30 + C40S)": "Splitt tariffs and data packs (contribution C30 + C40S)",
    "Splitt – direkte lønnsomhet": "Splitt - direct profitability",
    "Streamix – direkte bidrag siden lansering": "Streamix - direct contribution since launch",
    "Streamix (pd.SY.P255) bokført i FinMax: abonnement, oppkjøpsprodukter og COGS, samt driftskostnader merket på "
    "produktet. Streamix Mobil (GF068/GF069 under M4115) vises som memo utenfor totalen.":
        "Streamix (pd.SY.P255) as booked in FinMax: subscriptions, purchase products and COGS, plus opex tagged on "
        "the product. Streamix Mobil (GF068/GF069 under M4115) is shown as a memo outside the total.",
    "Forbrukerens fast-produkter har ingen fast Produktregnskap-kode; Streamix P255 fordeles med FinPlan-nøkler "
    "etter årsslutt.":
        "Consumer fixed products have no fixed product-cost-accounts code; Streamix P255 is allocated with FinPlan "
        "keys after year end.",
    "Avtalt med Tom 22.10.2026: booked + Streamix Mobil som memo.":
        "Agreed with Tom 22.10.2026: booked items + Streamix Mobil as a memo.",
    "Direkte bidrag Streamix": "Direct contribution Streamix",
    "Direkte bidrag = EBITDA før andre poster på produktet":
        "Direct contribution = EBITDA before other items on the product",
    "Direkte bidrag = EBITDA før andre poster på produktene":
        "Direct contribution = EBITDA before other items on the products",
    "Inntekter og varekost": "Revenue and COGS",
    "Inntekter (C30)": "Revenue (C30)",
    "Varekost (C40S)": "COGS (C40S)",
    "Bruttofortjeneste": "Gross profit",
    "Driftskostnader bokført på produktet": "Opex booked on the product",
    "Driftskostnader bokført på produktene": "Opex booked on the products",
    "Driftskostnader (C70)": "Opex (C70)",
    "Memo: Streamix Mobil (ikke i total)": "Memo: Streamix Mobil (not in the total)",
    "Memo: inntekter Streamix Mobil (C30)": "Memo: revenue Streamix Mobil (C30)",
    "Memo: varekost Streamix Mobil (C40S)": "Memo: COGS Streamix Mobil (C40S)",
    "Ikke med": "Not included",
    "Lønn og kostnader på koststed": "Salaries and costs on cost centres",
    "lønn (C50S) er aldri produktmerket, C60 og C75 bare delvis.":
        "salaries (C50S) are never product-tagged, C60 and C75 only partly.",
    "lønn (C50S) er aldri produktmerket, C60 og C75 bare delvis; fullt fordelt i Produktregnskapet.":
        "salaries (C50S) are never product-tagged, C60 and C75 only partly; fully allocated in the product cost "
        "accounts (Produktregnskapet).",
    "Familie fra produktlisten 2025 via forretningskart 2 (godkjent 15.10.2026); bare bokførte tall.":
        "Family from the 2025 product list via business map 2 (approved 15.10.2026); booked figures only.",
    # splitt decisions
    "håndsettnøkkel (key_same_as hw_cogs)": "handset key (key_same_as hw_cogs)",
    "Claude (brukeren hadde ingen preferanse)": "Claude (the user had no preference)",
    "innbytte totalt −3,5 MNOK YTD sep 26": "trade-in in total -3.5 MNOK YTD Sep 26",
}
# family texts with the family's own name in them
FAMILY_TEXT = [
    (re.compile(r"(.+) – direkte bidrag"), r"\1 - direct contribution"),
    (re.compile(r"Direkte bidrag (.+)"), r"Direct contribution \1"),
    (re.compile(r"Produktfamilien «(.+)» i FinMax \(pd\.SY\): (\d+) produktgruppe\(r\), bl\.a\. (.+)\. Direkte bidrag "
                r"= inntekter \+ varekost \+ driftskostnader bokført på produktene\."),
     r"Product family «\1» in FinMax (pd.SY): \2 product group(s), e.g. \3. Direct contribution = revenue + COGS + "
     r"opex booked on the products."),
]
KIND_SUFFIX = {"booked": " (booked)", "attributed": " (attributed 100 %)", "estimated": " (estimated)"}
CONCEPT_SOURCE = {"manual": "manual", "auto-familie": "auto family (2025 product list, business map 2)"}
CONCEPT_APPROVED_ON = {"splitt": "2026-10-14", "streamix": "2026-10-22"}     # the rest with the business map, 15.10


def _tr(text: str, where: str) -> str:
    if not text:
        return ""
    if text in CONCEPT_TEXT:
        return CONCEPT_TEXT[text]
    for pattern, repl in FAMILY_TEXT:
        if pattern.fullmatch(text):
            return pattern.sub(repl, text)
    raise ValueError(f"concepts.json {where}: no English text for {text!r}; add it to seed.CONCEPT_TEXT")


def _cpick(d: dict | None, where: str) -> dict | None:
    """A concepts.json amount (account + products|entities), expression ('expr') or reference ('ref': one item)."""
    if d is None:
        return None
    out = {"label": _tr(d.get("label", ""), where)}
    if "expr" in d or "ref" in d:
        out["expr"] = d.get("expr") or d["ref"]
    else:
        out |= {"account": d["account"], "products": d.get("products", []), "entities": d.get("entities", [])}
    return out


def _citem(it: dict, cid: str) -> dict:
    where = f"{cid}.{it['key']}"
    kind = it["kind"]
    out = {"key": it["key"], "label": _tr(it.get("label", it["key"]), where) + KIND_SUFFIX.get(kind, ""),
           "kind": kind, "note": _tr(it.get("note", ""), where)}
    if kind in ("booked", "attributed"):
        out |= {"account": it["account"], "products": it.get("products", []), "entities": it.get("entities", [])}
    elif kind == "sum":
        out["of"] = it["of"]
    elif kind == "estimated":
        out["base"] = _cpick(it["base"], where)
        share = {"text": _tr(it.get("key_text", ""), where)}
        if it.get("key_same_as"):
            share["same_as"] = it["key_same_as"]
        else:
            share |= {"num": _cpick(it["key_num"], where), "den": _cpick(it["key_den"], where)}
        out["key_share"] = share
    return out


def _translate_values(d, where: str):
    """Decisions: translate the texts, keep dates and codes."""
    if isinstance(d, dict):
        return {k: _translate_values(v, where) for k, v in d.items()}
    return CONCEPT_TEXT.get(d, d) if isinstance(d, str) else d


def seed_concepts(concepts_json: dict) -> list[Concept]:
    """Every concept of knowledge/concepts.json as a semantic Concept (approved: the user approved the business map
    on 15.10; the families are automatic but booked-only). Keys are mapped as they stand: an estimated item's
    key_num/key_den (an amount, an 'expr' over item keys, or a 'ref' to one item) become key_share num/den,
    key_same_as becomes key_share.same_as."""
    out = []
    for c in concepts_json["concepts"]:
        cid = c["id"]
        out.append(Concept(
            id=cid, title=_tr(c["title"], cid), names=_dedupe(c.get("names", [])), kind=c["kind"],
            description=_tr(c["description"], cid), entity=c.get("entity", "un.SY"),
            source=CONCEPT_SOURCE.get(c.get("source") or "manual", c.get("source") or "manual"),
            status="approved",
            approved={"by": "tom", "on": CONCEPT_APPROVED_ON.get(cid, "2026-10-15"), "via": "concepts.json"},
            notes=[_tr(n, cid) for n in c.get("notes", [])],
            decisions=_translate_values(c.get("decisions", {}), cid),
            total={"label": _tr(c["total"]["label"], cid), "of": c["total"]["of"]},
            check=[{"label": _tr(x["label"], cid), "expr": x["expr"], "source": _cpick(x.get("source"), cid)}
                   for x in c.get("checks", [])],
            block=[{"title": _tr(b["title"], cid), "item": [_citem(it, cid) for it in b["items"]]}
                   for b in c["blocks"]]))
    return out


def seed(knowledge_dir: Path, out_dir: Path) -> tuple[int, int, int]:
    """Write metrics/<id>.toml, facts/<id>.toml and concepts/<id>.toml under out_dir; returns (metrics, facts,
    concepts) written."""
    knowledge_dir, out_dir = Path(knowledge_dir), Path(out_dir)
    catalog = _read(knowledge_dir, "catalog.toml")
    metrics = seed_metrics(catalog, _read(knowledge_dir, "capabilities.toml"))
    facts = seed_facts(catalog, _read(knowledge_dir, "accounts_map.json"), _read(knowledge_dir, "org_map.json"),
                       _read(knowledge_dir, "finplan_map.json"))
    concepts = seed_concepts(_read(knowledge_dir, "concepts.json"))
    for kind, items in (("metrics", metrics), ("facts", facts), ("concepts", concepts)):
        d = out_dir / kind
        d.mkdir(parents=True, exist_ok=True)
        for it in items:
            (d / f"{it.id}.toml").write_text(to_toml(it), encoding="utf-8", newline="\n")
    return len(metrics), len(facts), len(concepts)


if __name__ == "__main__":                              # uv run python -m smartass.knowledge.seed < /dev/null
    from ..core import config
    print(seed(config.KNOWLEDGE_DIR, config.KNOWLEDGE_DIR / "semantic"))
