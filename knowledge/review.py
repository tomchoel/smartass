"""Gjennomgangsstatus for forretningskartet (review_status.json), felles for alle build_*_map.py.

Når brukeren godkjenner et område, lagres en liste over medlemmene som var med. Ved senere ombygging (etter månedlig
atlas-oppdatering) markeres bare NYE medlemmer som ikke gjennomgått, i stedet for at hele området blir utkast igjen.
    python review.py approve <område> "<kommentar>"     # område: accounts | products | org | finplan
"""
import datetime as dt
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
STATUS = os.path.join(HERE, "review_status.json")
JSON_OF = {"accounts": ("accounts_map.json", "accounts", "member"), "products": ("products_map.json", "products", "member"),
           "org": ("org_map.json", "entities", "member"), "finplan": ("finplan_map.json", "accounts", "member")}


def _load() -> dict:
    try:
        with open(STATUS, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def apply(area: str, rows: list[dict], key: str = "member") -> tuple[str, str]:
    """Mark rows reviewed / new since approval (adds a pitfall to new ones). Returns (status, sheet subtitle part)."""
    st = _load().get(area)
    if not st:
        for r in rows:
            r["reviewed"] = False
        return "UTKAST - til gjennomgang", "Utkast til gjennomgang"
    approved = set(st.get("members", []))
    new = 0
    for r in rows:
        r["reviewed"] = r.get(key) in approved
        if not r["reviewed"]:
            new += 1
            note = f"Ny siden godkjenning {st['approved']} – ikke gjennomgått."
            r["pitfalls"] = (note + " " + (r.get("pitfalls") or "")).strip()
    d = dt.date.fromisoformat(st["approved"]).strftime("%d.%m.%Y")
    status = f"GODKJENT {st['approved']} av {st.get('by', 'brukeren')}" + (f"; {new} nye medlemmer ikke gjennomgått" if new else "")
    return status, f"Godkjent {d}" + (f" · {new} nye ikke gjennomgått" if new else "")


def approve(area: str, comment: str = "", by: str = "Tom Hoel") -> int:
    fname, field, key = JSON_OF[area]
    with open(os.path.join(HERE, fname), encoding="utf-8") as f:
        rows = json.load(f)[field]
    allst = _load()
    allst[area] = {"approved": dt.date.today().isoformat(), "by": by, "comment": comment,
                   "members": sorted({r[key] for r in rows if r.get(key)})}
    with open(STATUS, "w", encoding="utf-8") as f:
        json.dump(allst, f, ensure_ascii=False, indent=1)
    return len(allst[area]["members"])


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "approve":
        n = approve(sys.argv[2], " ".join(sys.argv[3:]))
        print(f"{sys.argv[2]}: {n} medlemmer godkjent")
    else:
        print(__doc__)
