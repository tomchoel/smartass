"""Learned defaults (design §3.1): when the user answers a gate question, the answer can be saved so the gate states it
next time instead of asking. One readable per-user file, config.PREFERENCES:

    [[default]]
    item = "measure"        # measure | period | scenario | entity
    phrase = "costs"        # the words in the question it applies to (lower case)
    value = "opex"          # a measure key, ytd|fy, a catalog scenario key or a catalog entity key
    saved = "2026-11-02"

Later the team promotes good defaults into knowledge/capabilities.toml (shared)."""
from __future__ import annotations

import datetime as dt
import tomllib

from .. import catalog
from smartass.core import config
from ..host import InputError
from . import semantic

ITEMS = ("measure", "period", "scenario", "entity")


def _allowed(item: str) -> set[str]:
    if item == "measure":
        return {m.key for m in semantic.all_measures() if not m.vague}
    if item == "period":
        return {"ytd", "fy"}
    cat = catalog.load()
    return set(cat.scenarios) if item == "scenario" else {e.key for e in cat.entities}


def load() -> list[dict]:
    p = config.PREFERENCES
    if not p.exists():
        return []
    return list(tomllib.loads(p.read_text(encoding="utf-8")).get("default", []))


def lookup(item: str, phrase: str) -> str | None:
    phrase = phrase.strip().lower()
    return next((d["value"] for d in load() if d.get("item") == item and d.get("phrase") == phrase), None)


def _q(s: str) -> str:
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def remember(item: str, phrase: str, value: str) -> dict:
    """Save (or replace) the default for item+phrase. Returns the saved entry."""
    if item not in ITEMS:
        raise InputError(f"item must be one of {', '.join(ITEMS)}, not {item!r}")
    phrase = phrase.strip().lower()
    if not phrase:
        raise InputError("phrase is the words in the question the default applies to")
    allowed = _allowed(item)
    if value not in allowed:
        raise InputError(f"{item} default must be one of {', '.join(sorted(allowed))}, not {value!r}")
    entry = {"item": item, "phrase": phrase, "value": value, "saved": dt.date.today().isoformat()}
    rows = [d for d in load() if not (d.get("item") == item and d.get("phrase") == phrase)] + [entry]
    head = "# smartass learned defaults: the scope gate states these instead of asking. Edit or delete freely.\n"
    body = "".join("\n[[default]]\n" + "".join(f"{k} = {_q(str(d[k]))}\n" for k in ("item", "phrase", "value", "saved")
                                                if k in d) for d in rows)
    p = config.PREFERENCES
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(head + body, encoding="utf-8")
    return entry
