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

from .semantic import Fact, Metric, to_toml

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


def seed(knowledge_dir: Path, out_dir: Path) -> tuple[int, int]:
    """Write metrics/<id>.toml and facts/<id>.toml under out_dir; returns (metrics, facts) written."""
    knowledge_dir, out_dir = Path(knowledge_dir), Path(out_dir)
    catalog = _read(knowledge_dir, "catalog.toml")
    metrics = seed_metrics(catalog, _read(knowledge_dir, "capabilities.toml"))
    facts = seed_facts(catalog, _read(knowledge_dir, "accounts_map.json"), _read(knowledge_dir, "org_map.json"),
                       _read(knowledge_dir, "finplan_map.json"))
    for kind, items in (("metrics", metrics), ("facts", facts)):
        d = out_dir / kind
        d.mkdir(parents=True, exist_ok=True)
        for it in items:
            (d / f"{it.id}.toml").write_text(to_toml(it), encoding="utf-8", newline="\n")
    return len(metrics), len(facts)


if __name__ == "__main__":                              # uv run python -m smartass.knowledge.seed < /dev/null
    from ..core import config
    print(seed(config.KNOWLEDGE_DIR, config.KNOWLEDGE_DIR / "semantic"))
