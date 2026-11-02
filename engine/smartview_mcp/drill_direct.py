"""Drill lines without loading a drill page per piece (spike 2026-11-02, docs/audit/2026-11-02-drill-speed-spike.md).

Per piece, from ONE warm tab (a tab that loaded one real drill page): GET the drill URL's JSON form
(&FORMAT=SUI) - EPM's own drill query - and page getData with it, field for field ({QUERY, PARAMS: PARAMSARRAY,
PARAMSTYPE: PARAMSTYPEARRAY}). Read-only: nothing is built by us. getData from a tab that never loaded a drill page
gets the whole EPM session signed out, so it is only ever sent from the warm tab."""
from __future__ import annotations

import json
from dataclasses import dataclass

from .drill_browser import CellTooWide, DrillError


class NoDrillData(DrillError):
    """EPM's drill query has no lines here (data not loaded through Data Integration, or an empty cell)."""


class SignedOut(DrillError):
    """EPM answered with something that is not JSON (its sign-in / exit page)."""


@dataclass
class SuiQuery:
    body: str               # the getData POST body
    app_id: int | None
    partition: int | None


def sui_url(base: str, url: str, now_ms: int) -> str:
    """The drill URL's JSON form: DRILLURL dropped, FORMAT=SUI and a cache buster added (as the page asks it).

    Built by string operations, not parse/urlencode: every other parameter is passed through byte for byte, as the
    page sends Smart View's URL unchanged."""
    full = base.rstrip("/") + "/" + url.lstrip("./")
    full, _, frag = full.partition("#")
    path, _, query = full.partition("?")
    items = [p for p in query.split("&") if p and p.partition("=")[0] != "DRILLURL"]
    items += ["FORMAT=SUI", f"_={now_ms}"]
    return path + "?" + "&".join(items) + ("#" + frag if frag else "")


def parse_sui(text: str) -> SuiQuery:
    try:
        j = json.loads(text)
    except ValueError:
        raise SignedOut("EPM answered the drill query with a page, not JSON (signed out?)") from None
    if not isinstance(j, dict):
        raise SignedOut("EPM's drill query answer is not a JSON object")
    if "QUERY" not in j:
        why = str(j.get("details") or "no drill query in EPM's answer")
        if "EPMFDM-140373" in why:
            raise CellTooWide(why)
        if "does not return any data" in why:
            raise NoDrillData(why)
        raise DrillError(f"EPM refused the drill: {why}")
    body = json.dumps({"QUERY": j["QUERY"], "PARAMS": j.get("PARAMSARRAY", []),
                       "PARAMSTYPE": j.get("PARAMSTYPEARRAY", [])})
    return SuiQuery(body, j.get("APPLICATION_ID"), j.get("PARTITIONKEY"))
