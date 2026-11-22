"""The drill's pure parts, used by the token route (drill_token), drillstep and svhttp: the errors, the cell's
defaults, the SUI query handling, the answer parsing and the paging bookkeeping. Nothing here touches the network or
Excel.

Error messages and log lines never carry a URL query (the drill link holds an sso_token): redact() every URL that is
logged or put in a message."""
from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass
from typing import Callable

from ..core import config
from ..core.errors import InputError

log = logging.getLogger("smartview.drill")


class DrillError(Exception):
    pass


class TooManyLines(DrillError):
    def __init__(self, total: int):
        super().__init__(f"{total} lines")
        self.total = total


class CellTooWide(DrillError):
    """EPMFDM-140373: the cell spans more than 3 dimensions with over 1000 children each (summary drill refused).
    Narrow it (e.g. child entities) and drill again."""


DRILL_DEFAULTS = {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"}


def drill_cell_defaults(cell: dict[str, str]) -> tuple[dict[str, str], list[str]]:
    """A drill cell needs an account and a period; a missing entity drills the top member, with a warning."""
    missing = {"All Accounts", "All Time Periods"} - set(cell)
    if missing:
        raise InputError(f"fin_drill needs these dimensions in cell: {sorted(missing)}")
    if "All Entities" in cell:
        return dict(cell), []
    return {**cell, "All Entities": "All Entities"}, [
        "All Entities not given - drilled the top member; name an entity to narrow it"]


# ------------------------------------------------------------------ pure helpers
def redact(text: str) -> str:
    """Drop every URL query string (and any stray sso_token) from text meant for messages or logs."""
    text = re.sub(r"sso_token=[^&\s'\"]*", "sso_token=...", str(text))
    return re.sub(r"\?[^\s'\"]*", "?...", text)


def payload(text: str) -> dict | None:
    try:
        pay = json.loads(text)["responsePayload"]
    except (ValueError, KeyError, TypeError):
        return None
    return pay if isinstance(pay, dict) else None


LOST = (0, 301, 302, 303, 307, 401, 403)


def expired(status: int, pay: dict | None) -> bool:
    """A getData answer that means 'signed out': 401/403, a request lost on the way (status 0), a redirect
    (301/302/303/307: the token route never follows one, and the pod answers a lost session with one to its exit page)
    or a 200 without JSON (the sign-in page itself)."""
    return status in LOST or (status == 200 and pay is None)


class Pager:
    """getData paging bookkeeping, one answer at a time (drill_token's loop). take() says what comes next: 'more'
    (post again at .off), 'recover' (take a new session, then post at the same .off) or 'done'."""

    def __init__(self, max_lines: int, progress: Callable[[int, int | None], None] | None = None,
                 can_recover: bool = True):
        self.max_lines, self.progress = max_lines, progress
        self.items: list[dict] = []
        self.off, self.total, self.recovered = 0, None, not can_recover

    def take(self, status: int, text: str) -> str:
        pay = payload(text)
        if status != 200 or pay is None:
            if expired(status, pay) and not self.recovered:
                log.info("drill session expired at offset %d (HTTP %s); signing in again", self.off, status)
                self.recovered = True
                return "recover"
            why = "without data" if pay is None else "with an error"
            raise DrillError(f"getData returned HTTP {status} {why} at offset {self.off}.")
        if self.total is None and pay.get("totalResults") is not None:
            self.total = int(pay["totalResults"])
            if self.total > self.max_lines:
                raise TooManyLines(self.total)
        got = pay.get("items") or []
        self.items += got
        if len(self.items) > self.max_lines:
            raise TooManyLines(max(len(self.items), self.total or 0))
        if self.progress:
            self.progress(len(self.items), self.total)
        if not got or not pay.get("hasMore"):
            return "done"
        if self.total is not None and len(self.items) >= self.total + config.DRILL_PAGE:
            log.warning("drill paging stopped at the safety cap (%d items, totalResults %d, hasMore still set)",
                        len(self.items), self.total)
            return "done"
        self.off += len(got)
        return "more"

    def result(self) -> tuple[list[dict], int | None]:
        if self.total is not None and len(self.items) != self.total:
            log.warning("drill paged %d items but totalResults was %d", len(self.items), self.total)
        return self.items, self.total


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
