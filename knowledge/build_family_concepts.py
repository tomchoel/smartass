"""Begreper per produktfamilie (direkte bidrag) fra forretningskart 2 -> concepts.json (+ oversiktsark).

En node i pd.SY-treet hører til familie F når alle undergrupper med tall er F; begrepet bruker de høyeste slike nodene,
så ingenting telles to ganger. Hvert begrep: inntekter (C30), varekost (C40S), driftskostnader bokført på produktene
(C70) = direkte bidrag, kontroll mot EBITDA før andre poster på de samme produktene, og det som ikke ligger på produkt.
Håndlagde begreper (uten "source": "auto-familie") røres ikke.
    python build_family_concepts.py
"""
import json
import os
import re
import sqlite3
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from build_accounts_map import ATLAS  # noqa: E402

CONCEPTS = os.path.join(HERE, "concepts.json")
SOURCE = "auto-familie"
MIN_NOK = 5e6                     # families smaller than this (|C30| and |EBITDA_BOI| FY25) get no concept
NEUTRAL_NOK = 1e6                 # unlabelled groups below this do not break a family


def slug(s: str) -> str:
    s = s.lower().replace("æ", "ae").replace("ø", "o").replace("å", "a")
    return re.sub(r"[^a-z0-9]+", "_", s).strip("_")


def size(r: dict) -> float:
    return sum(abs(r.get(k) or 0) for k in ("fy2025_c30", "fy2025_c40s", "fy2025_c70"))


def family_roots(rows: list[dict]) -> dict[str, list[dict]]:
    db = sqlite3.connect(ATLAS)
    names = {r["member"] for r in rows}
    parent = {n: p for n, p in db.execute("SELECT name, parent FROM member WHERE connection='FinMax' "
                                          "AND dimension LIKE '%Product%'") if n in names}
    kids = defaultdict(list)
    for n, p in parent.items():
        if p in names:
            kids[p].append(n)
    by = {r["member"]: r for r in rows}
    fams: dict[str, set] = {}

    def walk(n):                  # set of families below n (None = neutral)
        if n in fams:
            return fams[n]
        if not kids[n]:
            r = by[n]
            f = r.get("family")
            fams[n] = {f} if f else ({"-"} if size(r) >= NEUTRAL_NOK else set())
        else:
            fams[n] = set().union(*(walk(k) for k in kids[n]))
        return fams[n]
    for n in names:
        walk(n)
    roots = defaultdict(list)
    for n in names:
        fs = fams[n]
        if len(fs) == 1 and "-" not in fs:
            f = next(iter(fs))
            if fams.get(parent.get(n)) != {f}:
                roots[f].append(by[n])
    return roots


def concept(fam: str, roots: list[dict]) -> dict:
    parts = [p.strip() for p in fam.split("/")]
    short = parts[-1]
    prods = sorted(r["member"] for r in roots)
    names = list(dict.fromkeys([short, short.lower(), fam, " ".join(parts[:2] + [short])]))
    ex = ", ".join(f"{r['code']} {r['label'].split(' ', 1)[-1]}" for r in sorted(roots, key=lambda r: -size(r))[:4])
    return {
        "id": "fam_" + slug(short + ("_" + parts[1] if len(parts) > 2 and parts[1].lower() not in short.lower() else "")),
        "title": f"{short} – direkte bidrag",
        "names": names,
        "kind": "product_family",
        "source": SOURCE,
        "description": f"Produktfamilien «{fam}» i FinMax (pd.SY): {len(prods)} produktgruppe(r), bl.a. {ex}. "
                       "Direkte bidrag = inntekter + varekost + driftskostnader bokført på produktene.",
        "entity": "un.SY",
        "blocks": [
            {"title": "Inntekter og varekost", "items": [
                {"key": "rev", "label": "Inntekter (C30)", "kind": "booked", "account": "ac.C30", "products": prods},
                {"key": "cogs", "label": "Varekost (C40S)", "kind": "booked", "account": "ac.C40S", "products": prods},
                {"key": "gp", "label": "Bruttofortjeneste", "kind": "sum", "of": ["rev", "cogs"]}]},
            {"title": "Driftskostnader bokført på produktene", "items": [
                {"key": "opex", "label": "Driftskostnader (C70)", "kind": "booked", "account": "ac.C70",
                 "products": prods}]},
            {"title": "Ikke med", "items": [
                {"key": "not_tagged", "kind": "missing", "label": "Lønn og kostnader på koststed",
                 "note": "lønn (C50S) er aldri produktmerket, C60 og C75 bare delvis; fullt fordelt i Produktregnskapet."}]}],
        "checks": [{"label": "Direkte bidrag = EBITDA før andre poster på produktene",
                    "source": {"account": "ac.EBITDA_BOI", "products": prods}, "expr": "rev+cogs+opex-$src"}],
        "total": {"label": f"Direkte bidrag {short}", "of": ["rev", "cogs", "opex"]},
        "notes": ["Familie fra produktlisten 2025 via forretningskart 2 (godkjent 15.10.2026); bare bokførte tall."],
        "fy2025": {k: round(sum(r.get(k) or 0 for r in roots), 2)
                   for k in ("fy2025_c30", "fy2025_c40s", "fy2025_c70", "fy2025_ebitda_boi")},
    }


def main():
    rows = json.load(open(os.path.join(HERE, "products_map.json"), encoding="utf-8"))["products"]
    roots = family_roots(rows)
    made, skipped = [], []
    for fam, rs in sorted(roots.items(), key=lambda x: -sum(r.get("fy2025_c30") or 0 for r in x[1])):
        c30 = sum(r.get("fy2025_c30") or 0 for r in rs)
        boi = sum(r.get("fy2025_ebitda_boi") or 0 for r in rs)
        if "kode utenfor produktlisten" in fam or max(abs(c30), abs(boi)) < MIN_NOK:
            skipped.append((fam, c30))
            continue
        made.append(concept(fam, rs))
    ids = [c["id"] for c in made]
    assert len(ids) == len(set(ids)), "duplicate concept ids"
    lib = json.load(open(CONCEPTS, encoding="utf-8"))
    keep = [c for c in lib["concepts"] if c.get("source") != SOURCE]
    clash = {c["id"] for c in keep} & set(ids)
    assert not clash, f"id clash with hand-made concepts: {clash}"
    lib["concepts"] = keep + made
    with open(CONCEPTS, "w", encoding="utf-8") as f:
        json.dump(lib, f, ensure_ascii=False, indent=1)
    tot = sum(c["fy2025"]["fy2025_c30"] for c in made)
    allc30 = sum(r.get("fy2025_c30") or 0 for r in rows if r["member"] in ("pd.SY.Mobile", "pd.SY.FIXED"))
    print(f"{len(made)} familiebegreper; C30 FY25 dekket {tot / 1e6:,.1f} av {allc30 / 1e6:,.1f} MNOK "
          f"({tot / allc30:.1%}); hoppet over {len(skipped)}: " + "; ".join(f"{f} ({v / 1e6:.1f})" for f, v in skipped))
    for c in made:
        print(f"  {c['id']:42} {c['fy2025']['fy2025_c30'] / 1e6:10.1f} {len(c['blocks'][0]['items'][0]['products']):3}  {c['title']}")
    return made


if __name__ == "__main__":
    main()
