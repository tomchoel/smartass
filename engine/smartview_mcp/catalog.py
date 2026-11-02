# engine/smartview_mcp/catalog.py
"""Semantic catalogue (knowledge/catalog.toml): P&L lines and identity checks per cube, business areas with aliases,
scenarios, KPIs and materiality - the one place that says what finance words mean in the cubes. templates.py and the
coming fin_* tools read it. parse() checks the structure; validate() checks every member against the atlas."""
from __future__ import annotations

import dataclasses
import datetime
import functools
import re
import tomllib
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from types import MappingProxyType

from . import config

KINDS = {"line", "sum", "grand"}
RULES = {"latest_forecast"}
BETTER = {"higher", "lower", "none"}
UNITS = {"pct", "nok", "nok_per_fte"}   # a KPI's unit; the cube sign is FinMax's (costs negative), sign = -1 shows a cost KPI positive
MEASURES = {"flow", "stock"}          # stock: a level (FTE); read Accumulated at period end, never sum Periodic months
BASES = {"period", "average_stock"}   # average_stock: the denominator is the average month-end level over the period
MATERIALITY_RULES = {"either", "both"}
CUBES = ("FinMax", "FinPlan")         # exactly these in [pnl] and [checks]; any other name is a typo
EXPR = re.compile(r"^[a-z_]\w*([+-][a-z_]\w*)*$")   # line keys joined by + and -


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
    concept: str | None = None    # the meaning across cubes: join actual vs budget on this, never on key
    measure: str = "flow"         # flow | stock


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
    generic_aliases: tuple[str, ...] = ()   # generic words ("total"): count only when no named area is matched
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
    basis: str = "period"         # period | average_stock (opening level + each month end, n+1 points)
    annualise: bool = False       # scale a part-year numerator to twelve months
    sign: int = 1                 # 1, or -1 to show a FinMax-negative cost positive (cost per FTE)
    better: str = "higher"        # higher | lower | none: the direction a variance on this KPI is judged by
    headline: bool = False        # the one KPI shown first for its cube (at most one per connection)


@dataclass(frozen=True)
class Materiality:
    pct: float
    abs_nok: float
    always: tuple[str, ...]
    rule: str = "either"          # either: pct OR abs_nok makes a variance material; both: pct AND abs_nok
    pct_floor_nok: float = 0      # the pct test only counts from this amount (keeps small lines quiet)


@dataclass(frozen=True)
class Catalog:
    pnl: Mapping[str, tuple[Line, ...]]           # the three tables are read-only (MappingProxyType)
    checks: Mapping[str, tuple[Check, ...]]
    scenarios: Mapping[str, Scenario]
    materiality: Materiality
    entities: tuple[Entity, ...]
    kpis: tuple[Kpi, ...]
    _whole: Mapping[str, Entity] = field(default_factory=dict, repr=False, compare=False)
    _patterns: tuple[tuple[re.Pattern, Entity, bool], ...] = field(default=(), repr=False, compare=False)

    def line_by_concept(self, cube: str, concept: str) -> Line | None:
        """The line of `cube` that carries `concept` (FinMax c73 and FinPlan oth are both other_items), or None."""
        return next((x for x in self.pnl.get(cube, ()) if x.concept == concept), None)


def _norm(s: str) -> str:
    return re.sub(r"[^\w&]+", " ", s.casefold()).strip()


def _patterns(entities: tuple[Entity, ...]) -> tuple[tuple[re.Pattern, Entity, bool], ...]:
    """Each alias and label as (regex over the original text, entity, generic): words in any case with any separators
    between them; aliases under 3 characters only in UPPER CASE (so "it" in "is it good?" is not IT)."""
    out = []
    for e in entities:
        names = [(a, False) for a in (*e.aliases, e.label)] + [(a, True) for a in e.generic_aliases]
        for a, generic in names:
            words = _norm(a).split()
            if not words:
                continue
            if len("".join(words)) < 3:
                rx = re.compile(rf"(?<![\w&]){re.escape(a.upper())}(?![\w&])")
            else:
                body = r"[^\w&]+".join(re.escape(w) for w in words)
                rx = re.compile(rf"(?<![\w&]){body}(?![\w&])", re.IGNORECASE)
            out.append((rx, e, generic))
    return tuple(out)


def _unique(keys, what: str) -> None:
    seen = set()
    for k in keys:
        if k in seen:
            raise CatalogError(f"duplicate {what} '{k}'")
        seen.add(k)


def _strings(value, where: str, name: str) -> tuple[str, ...]:
    if not isinstance(value, (list, tuple)) or not all(isinstance(x, str) for x in value):
        raise CatalogError(f"{where}: {name} must be a list of strings, not {type(value).__name__} {value!r}")
    return tuple(value)


def _build(cls, raw, where: str, lists: tuple[str, ...] = (), **extra):
    """One catalogue entry as `cls`. A non-table, an unknown or missing field, or a plain string where a list belongs
    is reported as "<where>: ..." so the broken entry can be found."""
    if not isinstance(raw, dict):
        raise CatalogError(f"{where}: expected a table, not {type(raw).__name__} {raw!r}")
    fields = [f for f in dataclasses.fields(cls) if f.name not in extra]
    for k in raw:
        if k not in {f.name for f in fields}:
            raise CatalogError(f"{where}: unknown field '{k}'")
    for f in fields:
        if f.name not in raw and f.default is dataclasses.MISSING and f.default_factory is dataclasses.MISSING:
            raise CatalogError(f"{where}: missing field '{f.name}'")
    args = {**raw, **extra}
    for name in lists:
        if name in args:
            args[name] = _strings(args[name], where, name)
    return cls(**args)


def _section(data: dict, name: str, kind: type, default=None):
    value = data.get(name, default)
    if value is None:
        raise CatalogError(f"missing section [{name}]")
    if not isinstance(value, kind):
        raise CatalogError(f"[{name}]: expected a {'table' if kind is dict else 'list'}, not {type(value).__name__}")
    return value


def _number(raw: dict, name: str) -> float:
    value = raw.get(name)
    if isinstance(value, bool):
        raise CatalogError(f"materiality: {name} {value!r} is not a number")
    try:
        return float(value)
    except (TypeError, ValueError):
        raise CatalogError(f"materiality: {name} {value!r} is not a number") from None


def _materiality(raw: dict) -> Materiality:
    for k in raw:
        if k not in {"pct", "abs_nok", "always", "rule", "pct_floor_nok"}:
            raise CatalogError(f"materiality: unknown field '{k}'")
    pct, abs_nok = _number(raw, "pct"), _number(raw, "abs_nok")
    if not 0 < pct <= 1:
        raise CatalogError(f"materiality: pct {pct} is not in (0, 1] (0.10 = 10 %)")
    if abs_nok < 0:
        raise CatalogError(f"materiality: abs_nok {abs_nok} is negative")
    floor = _number(raw, "pct_floor_nok") if "pct_floor_nok" in raw else 0.0
    if floor < 0:
        raise CatalogError(f"materiality: pct_floor_nok {floor} is negative")
    rule = raw.get("rule", "either")
    if rule not in MATERIALITY_RULES:
        raise CatalogError(f"materiality: rule {rule!r} is not one of {sorted(MATERIALITY_RULES)}")
    return Materiality(pct, abs_nok, _strings(raw.get("always", []), "materiality", "always"), rule, floor)


def _iso_date(value, where: str) -> str | None:
    """An ISO date as text; TOML may also hand a date value (abolished = 2026-06-01 without quotes)."""
    if value is None:
        return None
    if isinstance(value, datetime.date) and not isinstance(value, datetime.datetime):
        return value.isoformat()
    try:
        return datetime.date.fromisoformat(value).isoformat()
    except (TypeError, ValueError):
        raise CatalogError(f"{where}: abolished {value!r} is not an ISO date (YYYY-MM-DD)") from None


def parse(data: dict) -> Catalog:
    pnl = {cube: tuple(_build(Line, x, f"pnl.{cube}[{i}]") for i, x in enumerate(lines))
           for cube, lines in _section(data, "pnl", dict).items()}
    checks = {cube: tuple(_build(Check, x, f"checks.{cube}[{i}]") for i, x in enumerate(cs))
              for cube, cs in _section(data, "checks", dict, {}).items()}
    scenarios = {k: _build(Scenario, v, f"scenarios[{k}]", key=k)
                 for k, v in _section(data, "scenarios", dict).items()}
    materiality = _materiality(_section(data, "materiality", dict))
    entities = tuple(_build(Entity, e, f"entity[{i}]", ("aliases", "generic_aliases", "successors"))
                     for i, e in enumerate(_section(data, "entity", list)))
    entities = tuple(dataclasses.replace(e, abolished=_iso_date(e.abolished, f"entity '{e.key}'")) for e in entities)
    kpis = tuple(_build(Kpi, k, f"kpi[{i}]") for i, k in enumerate(_section(data, "kpi", list, [])))

    for section, table in (("pnl", pnl), ("checks", checks)):
        for cube in table:
            if cube not in CUBES:
                raise CatalogError(f"[{section}]: unknown cube '{cube}' (the cubes are {', '.join(CUBES)})")
        for cube in CUBES:
            if cube not in table:
                raise CatalogError(f"[{section}]: missing cube '{cube}'")
    keys = {cube: {x.key for x in lines} for cube, lines in pnl.items()}
    for cube, lines in pnl.items():
        _unique([x.key for x in lines], f"{cube} line key")
        for x in lines:
            if x.kind not in KINDS:
                raise CatalogError(f"{cube} line '{x.key}': kind '{x.kind}' is not one of {sorted(KINDS)}")
            if x.better not in BETTER:
                raise CatalogError(f"{cube} line '{x.key}': better '{x.better}' is not one of {sorted(BETTER)}")
            if x.measure not in MEASURES:
                raise CatalogError(f"{cube} line '{x.key}': measure '{x.measure}' is not one of {sorted(MEASURES)}")
        _unique([x.concept for x in lines if x.concept is not None], f"{cube} concept")
        grands = [x.key for x in lines if x.kind == "grand"]
        if len(grands) > 1:
            raise CatalogError(f"{cube} has {len(grands)} grand lines ({', '.join(grands)}); at most one")
    for cube, cs in checks.items():
        for c in cs:
            if not EXPR.match(c.expr):
                raise CatalogError(f"{cube} check '{c.label}': expression '{c.expr}' is not line keys joined by "
                                   "+ and -")
            for tok in re.findall(r"[a-z_]\w*", c.expr):
                if tok not in keys.get(cube, set()):
                    raise CatalogError(f"{cube} check '{c.label}': '{tok}' is not a {cube} line key")
    for s in scenarios.values():
        if s.connection not in pnl:
            raise CatalogError(f"scenario '{s.key}': connection '{s.connection}' is not a pnl cube ({', '.join(pnl)})")
        if s.scenario and s.rule:
            raise CatalogError(f"scenario '{s.key}' has both scenario and rule; give one")
        if not s.scenario and not s.rule:
            raise CatalogError(f"scenario '{s.key}' needs a scenario or rule")
        if s.rule and s.rule not in RULES:
            raise CatalogError(f"scenario '{s.key}': rule '{s.rule}' is not one of {sorted(RULES)}")
    _unique([k.key for k in kpis], "kpi key")
    for k in kpis:
        if k.connection not in pnl:
            raise CatalogError(f"kpi '{k.key}': unknown connection '{k.connection}' (not a pnl cube)")
        if k.unit not in UNITS:
            raise CatalogError(f"kpi '{k.key}': unit '{k.unit}' is not one of {sorted(UNITS)}")
        if k.basis not in BASES:
            raise CatalogError(f"kpi '{k.key}': basis '{k.basis}' is not one of {sorted(BASES)}")
        if not isinstance(k.annualise, bool):
            raise CatalogError(f"kpi '{k.key}': annualise must be true or false, not {k.annualise!r}")
        if type(k.sign) is not int or k.sign not in (1, -1):
            raise CatalogError(f"kpi '{k.key}': sign must be 1 or -1, not {k.sign!r}")
        if k.better not in BETTER:
            raise CatalogError(f"kpi '{k.key}': better '{k.better}' is not one of {sorted(BETTER)}")
        if not isinstance(k.headline, bool):
            raise CatalogError(f"kpi '{k.key}': headline must be true or false, not {k.headline!r}")
        for part in (k.numerator, k.denominator):
            if part not in keys[k.connection]:
                raise CatalogError(f"kpi '{k.key}': '{part}' is not a {k.connection} line key")
        num = next(x for x in pnl[k.connection] if x.key == k.numerator)
        if k.annualise and num.measure != "flow":
            raise CatalogError(f"kpi '{k.key}': annualise needs a flow numerator; "
                               f"{k.connection} line '{num.key}' is measure '{num.measure}'")
        den = next(x for x in pnl[k.connection] if x.key == k.denominator)
        if k.basis == "average_stock" and den.measure != "stock":
            raise CatalogError(f"kpi '{k.key}': basis average_stock needs a stock denominator; "
                               f"{k.connection} line '{den.key}' is measure '{den.measure}'")
    heads = [k.connection for k in kpis if k.headline]
    if len(heads) != len(set(heads)):
        raise CatalogError("at most one headline kpi per connection")
    every = set().union(*keys.values())
    for a in materiality.always:
        if a not in every:
            raise CatalogError(f"materiality.always: '{a}' is not a line key")
    _unique([e.key for e in entities], "entity key")
    for e in entities:
        for a in (*e.aliases, *e.generic_aliases):
            if not _norm(a):
                raise CatalogError(f"entity '{e.key}': alias {a!r} is empty after normalising")
    owner: dict[str, str] = {}      # normalised name -> the one area that uses it (one area may repeat a name)
    for e in entities:
        for name in (e.finmax, e.finplan, e.label, *e.aliases, *e.generic_aliases):
            n = _norm(name) if name else ""
            if n and owner.setdefault(n, e.key) != e.key:
                raise CatalogError(f"name '{n}' belongs to two areas: '{owner[n]}' and '{e.key}'")
    ekeys = {e.key for e in entities}
    for e in entities:
        if not e.finmax and not e.finplan:
            raise CatalogError(f"entity '{e.key}' needs finmax or finplan")
        for s in e.successors:
            if s not in ekeys or s == e.key:
                raise CatalogError(f"entity '{e.key}': successor '{s}' is not another entity key")
    whole: dict[str, Entity] = {}
    for e in entities:
        for name in (e.finmax, e.finplan, e.label, *e.aliases, *e.generic_aliases):
            if name:
                whole.setdefault(_norm(name), e)
    return Catalog(MappingProxyType(pnl), MappingProxyType(checks), MappingProxyType(scenarios), materiality,
                   entities, kpis, MappingProxyType(whole), _patterns(entities))


@functools.lru_cache(maxsize=4)
def _load(path: str) -> Catalog:
    try:
        data = tomllib.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, tomllib.TOMLDecodeError) as e:
        raise CatalogError(f"cannot read {path}: {e}") from None
    return parse(data)


def load(path: Path | None = None) -> Catalog:
    """The parsed catalogue, cached per path for the life of the process: an edit to catalog.toml takes effect only
    after the server (or Python) is restarted."""
    return _load(str(path or config.knowledge_file("catalog.toml")))


def resolve_entities(cat: Catalog, text: str) -> tuple[Entity, ...]:
    """Every business area a question or member name refers to, in order of appearance. A member name (finmax/finplan),
    alias or label as the whole text gives that area alone. Otherwise aliases are found as whole words, a match inside
    a longer match is dropped ("IT" in "IT-divisjonen"), and a generic alias ("total", "selskapet") only counts when
    nothing is named: "total opex for SSM" is SSM, but "Telenor Norge vs Wholesale" is both."""
    whole = cat._whole.get(_norm(text))
    if whole:
        return (whole,)
    spans = [(m.start(), m.end(), e, g) for rx, e, g in cat._patterns for m in rx.finditer(text)]
    if any(not g for *_, g in spans):
        spans = [x for x in spans if not x[3]]
    kept = [(a, b, e) for a, b, e, _ in spans
            if not any(a2 <= a and b <= b2 and (b2 - a2) > (b - a) for a2, b2, *_ in spans)]
    found: list[Entity] = []
    for _, _, e in sorted(kept, key=lambda x: (x[0], -(x[1] - x[0]))):
        if e not in found:
            found.append(e)
    return tuple(found)


ACCOUNT_DIM = {"FinMax": "All Accounts", "FinPlan": "Account"}
ENTITY_DIM = {"FinMax": "All Entities", "FinPlan": "Entity"}
SCENARIO_DIM = {"FinMax": "All Scenarios", "FinPlan": "Scenario"}
VERSION_DIM = {"FinPlan": "Version"}


def validate(cat: Catalog, db) -> tuple[list[str], list[str], list[str]]:
    """(problems, unchecked, notes): catalogue members the atlas does not have; dimensions not checked because the
    atlas has not built them; and members of abolished areas that are gone (expected after a reorg, so only a note).
    The atlas's age is not judged here (the doctor's atlas check does that)."""
    from . import atlas
    problems: list[str] = []
    unchecked: list[str] = []
    notes: list[str] = []

    def check(conn: str, dim: str | None, name: str, what: str, abolished: str | None = None) -> None:
        if dim is None:
            return
        if not atlas.built(db, conn, dim):
            if f"{conn} / {dim}" not in unchecked:
                unchecked.append(f"{conn} / {dim}")
            return
        if not atlas.has_member(db, conn, dim, name):
            if abolished:
                notes.append(f"{what}: '{name}' no longer in {conn} / {dim} (abolished {abolished})")
            else:
                problems.append(f"{what}: '{name}' is not in {conn} / {dim}")

    for cube, lines in cat.pnl.items():
        for x in lines:
            check(cube, ACCOUNT_DIM.get(cube), x.account, f"{cube} line '{x.key}'")
    for e in cat.entities:
        what = e.key if e.abolished else f"entity '{e.key}'"
        if e.finmax:
            check("FinMax", ENTITY_DIM["FinMax"], e.finmax, what, e.abolished)
        if e.finplan:
            check("FinPlan", ENTITY_DIM["FinPlan"], e.finplan, what, e.abolished)
    for s in cat.scenarios.values():
        if s.scenario:
            check(s.connection, SCENARIO_DIM.get(s.connection), s.scenario, f"scenario '{s.key}'")
        if s.version:
            check(s.connection, VERSION_DIM.get(s.connection), s.version, f"scenario '{s.key}' version")
    return problems, unchecked, notes
