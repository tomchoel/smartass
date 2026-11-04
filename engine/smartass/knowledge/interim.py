"""Interim metrics until the semantic layer (sub-spec 2): v1's knowledge/catalog.toml P&L lines, joined across
cubes on `concept` (never on `key`). The catalogue is owner-approved, so every metric is 'approved'. Materiality is
the catalogue's owner rule: NOK 5 million, or 10 % from NOK 500k (= a tenth of the absolute threshold, which is how
model.findings.material reads it)."""
from __future__ import annotations

import tomllib

from ..core import config as v1config

from ..model.resolver import MetricDef

SIGN = {"revenue": "income", "gross_margin": "income", "ebitda_boi": "income", "ebitda": "income",
        "cogs": "cost", "personnel": "cost", "other_opex": "cost", "opex": "cost", "da": "cost"}


def catalog_metrics(path=None) -> list[MetricDef]:
    data = tomllib.loads((path or v1config.KNOWLEDGE_DIR / "catalog.toml").read_text(encoding="utf-8"))
    mat = data["materiality"]
    materiality = (float(mat["abs_nok"]), float(mat["pct"]))
    by_concept: dict[str, dict] = {}
    for cube in ("FinMax", "FinPlan"):
        for line in data["pnl"][cube]:
            concept = line.get("concept")
            if not concept:
                continue
            d = by_concept.setdefault(concept, {"label": line["label"], "members": {}, "stock": False})
            d["members"][cube] = (line["account"],)
            d["stock"] = d["stock"] or line.get("measure") == "stock"
            if cube == "FinMax":
                d["label"] = line["label"]
    out = []
    for concept, d in by_concept.items():
        stock = d["stock"]
        if "FinMax" in d["members"] and not stock:     # FinLite = the same accounts as FinMax; stock (FTE) is not in it
            d["members"]["FinLite"] = d["members"]["FinMax"]
        out.append(MetricDef(d["label"], d["members"], sign=SIGN.get(concept, "neutral"), status="approved",
                             materiality=materiality, balance="Accumulated" if stock else "Periodic",
                             aggregation="last" if stock else "sum"))
    return out
