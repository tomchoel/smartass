"""A split's sibling drill queries built from ONE of EPM's own (owner decision 2026-11-03, docs/audit/
2026-11-03-parallel-spike.md): EPM answers drill queries one at a time (~0.7 s each), getData runs in parallel.

EPM's drill query (SUI) is a read-only SELECT on the Data Integration table with IN-lists of the cell's members:
`Entityx IN (?,...)` / `Accountx IN (?,...)` hold the member itself plus all its descendants (live 2026-11-03: the
atlas reproduces EPM's list exactly for un.SY.2TEC, 204 members). A sibling differs only in that list, so its query
is the template with the list swapped. Guards, each falling back to asking EPM:
- the template's list must equal the atlas members of the template cell (else the atlas does not match EPM);
- one clause for the column, placeholders = PARAMS = PARAMSTYPE, at most MAX_MEMBERS members;
- the caller accepts a derived piece only when its lines tie to the sibling's cube value."""
from __future__ import annotations

import json
import re

COLUMNS = {"All Entities": "Entityx", "All Accounts": "Accountx"}
MAX_MEMBERS = 1000                 # Oracle's IN-list limit; EPM refuses wider cells (EPMFDM-140373) anyway


def members(connection: str, member: str, db=None) -> list[str] | None:
    """The member and all its descendants from the local atlas (EPM's IN-list), or None when unknown/too many."""
    from smartass.knowledge import atlas
    own = db is None
    db = db or atlas.connect()
    try:
        kids = atlas.descendants(db, connection, member, limit=MAX_MEMBERS + 1)
    except Exception:   # noqa: BLE001 - no atlas: ask EPM instead
        return None
    finally:
        if own:
            db.close()
    names = list(dict.fromkeys([member] + [m["name"] for m in kids]))     # shared members appear once
    return names if len(names) <= MAX_MEMBERS else None


def _clause(query: str, column: str) -> re.Match | None:
    found = list(re.finditer(rf"\b{column}\s+IN\s*\(\s*\?(?:\s*,\s*\?)*\s*\)|\b{column}\s*=\s*\?", query, re.I))
    return found[0] if len(found) == 1 else None


def derive(body: str, dim: str, template_members: list[str], sibling_members: list[str]) -> str | None:
    """The getData body of a sibling cell from the template's body, or None when any guard fails."""
    column = COLUMNS.get(dim)
    if column is None or not sibling_members or len(sibling_members) > MAX_MEMBERS:
        return None
    try:
        j = json.loads(body)
        query, params, types = j["QUERY"], list(j["PARAMS"]), list(j["PARAMSTYPE"])
    except (ValueError, KeyError, TypeError):
        return None
    if not isinstance(query, str) or not query.count("?") == len(params) == len(types):
        return None
    m = _clause(query, column)
    if m is None:
        return None
    at, n = query[:m.start()].count("?"), m.group(0).count("?")
    if (len(set(template_members)) != len(template_members) or n != len(template_members)
            or set(params[at:at + n]) != set(template_members) or any(t != "String" for t in types[at:at + n])):
        return None
    k = len(sibling_members)
    col = re.match(r"\w+", m.group(0)).group(0)                  # EPM's own spelling of the column
    new_query = query[:m.start()] + f"{col} IN ({','.join('?' * k)})" + query[m.end():]
    return json.dumps({"QUERY": new_query, "PARAMS": params[:at] + list(sibling_members) + params[at + n:],
                       "PARAMSTYPE": types[:at] + ["String"] * k + types[at + n:]})
