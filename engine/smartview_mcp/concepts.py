"""Business concepts (knowledge\\concepts.json) -> FinMax pulls + answer spec, and scoping lookups.

A concept is blocks of items. Item kinds:
  booked      account x products|entities, tagged to the concept in FinMax       label suffix "(booked)"
  attributed  not tagged, but belongs to the concept in full                      "(attributed 100 %)"
  estimated   shared base x key (num/den, or the key of another item)             "(estimated)"
  sum         sum of other items
  missing     not visible in FinMax - shown as a text line, not in the total
The answer shows every item for the period YTD vs the same period last year, then the total, the keys (base, numerator,
denominator, share) and the FinMax basis (one line per member x account), so every number is traceable and live."""
from __future__ import annotations

import json
import re

from . import atlas, config

SUFFIX = {"booked": " (booked)", "attributed": " (attributed 100 %)", "estimated": " (estimated)"}


def _load(name: str) -> dict:
    p = config.knowledge_file(name)
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def library() -> list[dict]:
    return _load("concepts.json").get("concepts", [])


STOP = {"for", "and", "the", "med", "til", "som", "hva", "har", "kan", "vis", "hvor", "mye", "per", "fra", "ved", "der",
        "det", "den", "denne", "eller", "også", "av", "fra", "mot", "with", "what", "how", "show", "from", "all", "alle",
        "total", "totalt", "sum", "del", "andre", "other", "annet", "annen", "inn", "ut",
        # profitability words say WHAT to compute, not WHICH business - they must not pick the concept
        "lønnsomhet", "lønnsom", "bidrag", "direkte", "tjener", "tjene", "koster", "profitability", "contribution",
        "resultat", "earn", "earns", "cost", "costs"}


# question words (Norwegian/English) -> words used in the maps (member names are English, meanings Norwegian)
SYN = {"abonnement": "subscriptions", "abonnementer": "subscriptions", "kunder": "customers", "inntekt": "revenue", "inntekter": "revenue", "omsetning": "revenue", "varekost": "cogs",
       "lønn": "salaries", "lønnskostnader": "salaries", "årsverk": "fte", "ansatte": "fte", "avskrivninger": "depreciation",
       "håndsett": "handset", "mobil": "mobile", "fastnett": "fixed", "bredbånd": "broadband", "forsikring": "insurance",
       "strøm": "electricity", "kraft": "power", "markedsføring": "marketing", "provisjon": "commissions",
       "konsulenter": "consultants", "grossist": "wholesale", "salg": "sales", "margin": "margin",
       "subscriptions": "abonnement", "revenue": "inntekter", "cogs": "varekost", "fte": "årsverk",
       "mobilabonnement": "mobile subscriptions", "bredbåndsabonnement": "broadband subscriptions",
       "tv-abonnement": "subscriptions", "mobilinntekter": "mobile revenue", "fastnettinntekter": "fixed revenue",
       "driftskostnader": "opex", "bruttofortjeneste": "gross profit", "bruttomargin": "gross margin",
       "kundetap": "losses receivables", "bedrift": "business", "bedriftsmarked": "business",
       "privat": "consumer", "privatmarked": "consumer", "forbruker": "consumer", "telefoni": "phone",
       "telefon": "phone", "fasttelefoni": "fixed phone", "fiberbredbånd": "fiber broadband", "husleie": "lease buildings", "energi": "energy electricity"}


def _raw_words(s: str) -> set[str]:
    ws = {w for w in re.findall(r"[\wæøå-]+", s.lower()) if len(w) > 2 and w not in STOP and not w.isdigit()}
    return ws | {p for w in ws if "-" in w for p in w.split("-") if len(p) > 2}


def _words(s: str) -> set[str]:
    ws = _raw_words(s)
    return ws | {x for w in ws if w in SYN for x in SYN[w].split()}


def _score(question: str, hay_words: set[str]) -> float:
    """One point per question word found in the hay (itself, or all words of its synonym; part of a multi-word
    synonym gives part of a point), so a word and its translation never count twice."""
    s = 0.0
    for w in _raw_words(question):
        if w in hay_words:
            s += 1
        elif w in SYN:
            exp = set(SYN[w].split())
            s += len(exp & hay_words) / len(exp)
    return s


def scored(text: str) -> list[tuple[float, dict]]:
    """(score, concept) for every concept whose id/title/names match the text, best first. A name contained in the
    text scores 10; otherwise one point per matching question word (_score)."""
    t = text.lower()
    hits = []
    for c in library():
        names = [c.get("id", "")] + [c.get("title", "")] + list(c.get("names", []))
        score = max([10 if n and len(n) > 3 and n.lower() in t else _score(text, _words(n)) for n in names] + [0])
        if score:
            hits.append((score, c))
    return sorted(hits, key=lambda x: -x[0])


def find(text: str) -> list[dict]:
    """Concepts whose id/title/names match the text (best first)."""
    return [c for _, c in scored(text)]


def get(concept: str) -> dict:
    for c in library():
        if c.get("id") == concept:
            return c
    found = find(concept)
    if not found:
        known = ", ".join(c.get("id", "") for c in library())
        raise ValueError(f"Unknown concept '{concept}'. Known: {known}. Define it in knowledge\\concepts.json "
                         "(research -> propose -> user confirms).")
    return found[0]


def _label(dim_member: str) -> str:
    db = atlas.connect()
    r = db.execute("SELECT label FROM member WHERE connection='FinMax' AND name=? LIMIT 1", (dim_member,)).fetchone()
    if not r:
        return dim_member
    return re.sub(r"\s*\((pd|un|ac)[^)]*\)\s*$", "", r[0]).strip()


def plan(c: dict, period: str, last_year: str, entity: str | None = None) -> tuple[list[dict], dict]:
    """(pulls, answer spec) for concept c, YTD `period` vs `last_year`."""
    entity = entity or c.get("entity", "un.SY")
    details: dict[tuple, str] = {}             # (kind, member, account) -> detail line key
    detail_lines: list[dict] = []
    prods, p_accts, ents, e_accts = [], [], [], []

    def add(lst, x):
        if x not in lst:
            lst.append(x)

    def src_keys(src: dict) -> list[str]:
        acct = src["account"]
        keys = []
        for kind, members in (("p", src.get("products") or []), ("e", src.get("entities") or [])):
            for m in members:
                k = (kind, m, acct)
                if k not in details:
                    details[k] = f"d{len(details) + 1}"
                    pull = "p" if kind == "p" else "e"
                    if kind == "p":
                        add(prods, m), add(p_accts, acct)
                    else:
                        add(ents, m), add(e_accts, acct)
                    detail_lines.append({"key": details[k], "aux": True,
                                         "label": f"{_label(m)} · {acct.removeprefix('ac.')}",
                                         "cells": [{"pull": pull, "row": [m], "col": [acct, period]},
                                                   {"pull": pull, "row": [m], "col": [acct, last_year]},
                                                   {"var": [0, 1]}, {"varpct": [0, 1]}]})
                keys.append(details[k])
        return keys

    def two(cell):
        return [cell(0), cell(1), {"var": [0, 1]}, {"varpct": [0, 1]}]

    sections, key_lines, key_of, missing_notes = [], [], {}, []
    for blk in c.get("blocks", []):
        lines = []
        for it in blk.get("items", []):
            k, kind = it["key"], it.get("kind", "booked")
            label = it.get("label", k) + SUFFIX.get(kind, "")
            if kind in ("booked", "attributed"):
                ks = src_keys(it)
                lines.append({"key": k, "label": label, "cells": two(lambda j, ks=ks: {"sum": ks})})
            elif kind == "sum":
                lines.append({"key": k, "label": label, "kind": "sum", "cells": two(lambda j, o=it["of"]: {"sum": o})})
            elif kind == "estimated":
                base = f"{k}_base"
                key_lines.append({"key": base, "aux": True, "label": f"{it['label']}: base – {it['base']['label']}",
                                  "cells": two(lambda j, ks=src_keys(it["base"]): {"sum": ks})})
                if it.get("key_same_as"):
                    kk = key_of[it["key_same_as"]]
                else:
                    num, den, kk = f"{k}_num", f"{k}_den", f"{k}_key"
                    for part, lk in (("key_num", num), ("key_den", den)):
                        s = it[part]
                        if "expr" in s:
                            cells = two(lambda j, e=s["expr"]: {"expr": e})
                        elif "ref" in s:
                            cells = two(lambda j, r=s["ref"]: {"ref": {"key": r, "col": j}})
                        else:
                            cells = two(lambda j, ks=src_keys(s): {"sum": ks})
                        key_lines.append({"key": lk, "aux": True, "label": f"{it['label']}: key – {s.get('label', lk)}",
                                          "cells": cells})
                    key_lines.append({"key": kk, "aux": True, "num": "prosent1",
                                      "label": f"{it['label']}: key = {it.get('key_text', 'share')}",
                                      "cells": [{"expr": f"{num}/{den}"}, {"expr": f"{num}/{den}"}, None, None]})
                key_of[k] = kk
                lines.append({"key": k, "label": label, "cells": two(lambda j, b=base, kk=kk: {"expr": f"{b}*{kk}"})})
            elif kind == "missing":
                lines.append({"key": k, "label": f"{it.get('label', k)} (not in FinMax)",
                              "cells": [{"text": "not included"}, None, None, None]})
                missing_notes.append(f"{it.get('label', k)}: {it.get('note', 'not in FinMax')}")
            else:
                raise ValueError(f"unknown item kind '{kind}' in concept {c.get('id')}")
        sections.append({"title": blk["title"], "lines": lines})
    checks = []
    for i, ch in enumerate(c.get("checks", [])):
        ref = f"chk{i}"
        key_lines.append({"key": ref, "aux": True, "label": f"Check: {ch['label']}",
                          "cells": two(lambda j, ks=src_keys(ch["source"]): {"sum": ks})})
        checks.append({"label": ch["label"], "expr": ch["expr"].replace("$src", ref), "columns": [0, 1]})
    tot = c.get("total")
    if tot:
        sections.append({"title": "Sum", "lines": [{"key": "total", "label": tot["label"], "kind": "grand",
                                                     "cells": two(lambda j: {"sum": tot["of"]})}]})
    if key_lines:
        sections.append({"title": "Allocation keys and checks", "lines": key_lines})
    sections.append({"title": "Basis from FinMax (member · account)", "lines": detail_lines})

    pulls = []
    pov = {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Accumulated"}
    if prods:
        pulls.append({"id": "p", "connection": "FinMax", "rows": [{"dimension": "All Products", "members": prods}],
                      "columns": [{"dimension": "All Accounts", "members": p_accts},
                                  {"dimension": "All Time Periods", "members": [period, last_year]}],
                      "pov": {**pov, "All Entities": entity}})
    if ents:
        pulls.append({"id": "e", "connection": "FinMax", "rows": [{"dimension": "All Entities", "members": ents}],
                      "columns": [{"dimension": "All Accounts", "members": e_accts},
                                  {"dimension": "All Time Periods", "members": [period, last_year]}],
                      "pov": pov, "title": None})
    est = [it for b in c.get("blocks", []) for it in b.get("items", []) if it.get("kind") == "estimated"]
    notes = list(c.get("notes", [])) + missing_notes + [
        f"{it['label']}: estimated = base ({it['base']['label']}) × key ({it.get('key_text', '')})." for it in est]
    spec = {"title": f"{c.get('title', c.get('id'))} — YTD {period} vs {last_year}",
            "subtitle": f"NOK · FinMax · Actual · {entity} · Accumulated",
            "columns": [{"header": f"YTD {period}"}, {"header": f"YTD {last_year}"}, {"header": "Change"},
                        {"header": "Change %", "kind": "prosent1"}],
            "sections": sections, "checks": checks, "notes": notes,
            "source": "Source: Oracle EPM Cloud (FinMax) via Smart View; definition from knowledge\\concepts.json."}
    return pulls, spec


