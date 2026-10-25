# engine/smartview_mcp/catalog.py
"""Semantic catalogue (knowledge/catalog.toml): P&L lines and identity checks per cube, business areas with aliases,
scenarios, KPIs and materiality - the one place that says what finance words mean in the cubes. templates.py and the
coming fin_* tools read it. parse() checks the structure; validate() checks every member against the atlas."""
from __future__ import annotations

import functools
import re
import tomllib
from dataclasses import dataclass, field
from pathlib import Path

from . import config

KINDS = {"line", "sum", "grand"}
RULES = {"latest_forecast"}


class CatalogError(ValueError):
    """knowledge/catalog.toml is broken (the message names the entry)."""


@dataclass(frozen=True)
class Line:
    key: str
    account: str
    kind: str
    label: str
    better: str = "higher"
    num: str | None = None


@dataclass(frozen=True)
class Check:
    label: str
    expr: str


@dataclass(frozen=True)
class Scenario:
    key: str
    connection: str
    label: str
    scenario: str | None = None
    version: str | None = None
    rule: str | None = None


@dataclass(frozen=True)
class Entity:
    key: str
    label: str
    aliases: tuple[str, ...]
    finmax: str | None = None
    finplan: str | None = None
    abolished: str | None = None
    successors: tuple[str, ...] = ()
    note: str = ""


@dataclass(frozen=True)
class Kpi:
    key: str
    label: str
    connection: str
    numerator: str
    denominator: str
    unit: str


@dataclass(frozen=True)
class Materiality:
    pct: float
    abs_nok: float
    always: tuple[str, ...]


@dataclass(frozen=True)
class Catalog:
    pnl: dict[str, tuple[Line, ...]]
    checks: dict[str, tuple[Check, ...]]
    scenarios: dict[str, Scenario]
    materiality: Materiality
    entities: tuple[Entity, ...]
    kpis: tuple[Kpi, ...]
    _whole: dict[str, Entity] = field(default_factory=dict, repr=False, compare=False)
    _patterns: tuple[tuple[re.Pattern, Entity], ...] = field(default=(), repr=False, compare=False)


COMPANY = "company"   # the entity for the whole company: a generic word like "total" must not beat a named area


def _norm(s: str) -> str:
    return re.sub(r"[^\w&]+", " ", s.casefold()).strip()


def _patterns(entities: tuple[Entity, ...]) -> tuple[tuple[re.Pattern, Entity], ...]:
    """Each alias and label as a regex over the original text: words in any case with any separators between them;
    aliases under 3 characters only in UPPER CASE (so "it" in "is it good?" is not IT)."""
    out = []
    for e in entities:
        for a in (*e.aliases, e.label):
            words = _norm(a).split()
            if not words:
                continue
            if len("".join(words)) < 3:
                rx = re.compile(rf"(?<![\w&]){re.escape(a.upper())}(?![\w&])")
            else:
                body = r"[^\w&]+".join(re.escape(w) for w in words)
                rx = re.compile(rf"(?<![\w&]){body}(?![\w&])", re.IGNORECASE)
            out.append((rx, e))
    return tuple(out)


def _unique(keys, what: str) -> None:
    seen = set()
    for k in keys:
        if k in seen:
            raise CatalogError(f"duplicate {what} '{k}'")
        seen.add(k)


def parse(data: dict) -> Catalog:
    try:
        pnl = {cube: tuple(Line(**x) for x in lines) for cube, lines in data["pnl"].items()}
        checks = {cube: tuple(Check(**x) for x in cs) for cube, cs in data.get("checks", {}).items()}
        scenarios = {k: Scenario(key=k, **v) for k, v in data["scenarios"].items()}
        materiality = Materiality(float(data["materiality"]["pct"]), float(data["materiality"]["abs_nok"]),
                                  tuple(data["materiality"]["always"]))
        entities = tuple(Entity(**{**e, "aliases": tuple(e.get("aliases", ())),
                                   "successors": tuple(e.get("successors", ()))}) for e in data["entity"])
        kpis = tuple(Kpi(**k) for k in data.get("kpi", ()))
    except (KeyError, TypeError) as e:
        raise CatalogError(f"missing or unknown field: {e}") from None
    keys = {cube: {x.key for x in lines} for cube, lines in pnl.items()}
    for cube, lines in pnl.items():
        _unique([x.key for x in lines], f"{cube} line key")
        for x in lines:
            if x.kind not in KINDS:
                raise CatalogError(f"{cube} line '{x.key}': kind '{x.kind}' is not one of {sorted(KINDS)}")
    for cube, cs in checks.items():
        for c in cs:
            for tok in re.findall(r"[a-z_][a-z0-9_]*", c.expr):
                if tok not in keys.get(cube, set()):
                    raise CatalogError(f"{cube} check '{c.label}': '{tok}' is not a {cube} line key")
    for s in scenarios.values():
        if not s.scenario and not s.rule:
            raise CatalogError(f"scenario '{s.key}' needs a scenario or rule")
        if s.rule and s.rule not in RULES:
            raise CatalogError(f"scenario '{s.key}': rule '{s.rule}' is not one of {sorted(RULES)}")
    for k in kpis:
        for part in (k.numerator, k.denominator):
            if part not in keys.get(k.connection, set()):
                raise CatalogError(f"kpi '{k.key}': '{part}' is not a {k.connection} line key")
    every = set().union(*keys.values())
    for a in materiality.always:
        if a not in every:
            raise CatalogError(f"materiality.always: '{a}' is not a line key")
    _unique([e.key for e in entities], "entity key")
    _unique([_norm(a) for e in entities for a in e.aliases], "entity alias")
    ekeys = {e.key for e in entities}
    for e in entities:
        for s in e.successors:
            if s not in ekeys or s == e.key:
                raise CatalogError(f"entity '{e.key}': successor '{s}' is not another entity key")
    whole: dict[str, Entity] = {}
    for e in entities:
        for name in (e.finmax, e.finplan, e.label, *e.aliases):
            if name:
                whole.setdefault(_norm(name), e)
    return Catalog(pnl, checks, scenarios, materiality, entities, kpis, whole, _patterns(entities))


@functools.lru_cache(maxsize=4)
def _load(path: str) -> Catalog:
    try:
        data = tomllib.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, tomllib.TOMLDecodeError) as e:
        raise CatalogError(f"cannot read {path}: {e}") from None
    return parse(data)


def load(path: Path | None = None) -> Catalog:
    return _load(str(path or config.knowledge_file("catalog.toml")))


def resolve_entities(cat: Catalog, text: str) -> tuple[Entity, ...]:
    """Every business area a question or member name refers to, in order of appearance. A member name (finmax/finplan),
    alias or label as the whole text gives that area alone. Otherwise aliases are found as whole words, a match inside
    a longer match is dropped ("IT" in "IT-divisjonen"), and the company only counts when no other area is named."""
    whole = cat._whole.get(_norm(text))
    if whole:
        return (whole,)
    spans = [(m.start(), m.end(), e) for rx, e in cat._patterns for m in rx.finditer(text)]
    kept = [(a, b, e) for a, b, e in spans
            if not any(a2 <= a and b <= b2 and (b2 - a2) > (b - a) for a2, b2, _ in spans)]
    found: list[Entity] = []
    for _, _, e in sorted(kept, key=lambda x: (x[0], -(x[1] - x[0]))):
        if e not in found:
            found.append(e)
    if len(found) > 1:
        found = [e for e in found if e.key != COMPANY]
    return tuple(found)


def resolve_entity(cat: Catalog, text: str) -> Entity | None:
    """The one business area a question or member name refers to (see resolve_entities), or None when it names none
    or several. An abolished area is returned as such: callers check `abolished` and `successors`."""
    found = resolve_entities(cat, text)
    return found[0] if len(found) == 1 else None
