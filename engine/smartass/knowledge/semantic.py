"""The semantic layer's source files (sub-spec 2 §2): one TOML file per metric (`metrics/<id>.toml`), per fact
(`facts/<id>.toml`) and per product concept (`concepts/<id>.toml`, Plan 2C), reviewable in git. `load_semantic` and
`load_concepts` validate them; `build` compiles them with the atlas into knowledge.db."""
from __future__ import annotations

import json
import re
import tomllib
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator, model_validator

ID = re.compile(r"^[a-z0-9_]+$")
MONTH = re.compile(r"^\d{4}-\d{2}$")
CUBES = ("FinMax", "FinLite", "FinPlan")


class SemanticError(ValueError):
    pass


def _month(v: str, info) -> str:
    if v and not MONTH.match(v):
        raise ValueError(f"{info.field_name} must be 'YYYY-MM' or empty, got {v!r}")
    return v


def _id(v: str) -> str:
    if not ID.match(v):
        raise ValueError(f"id must match [a-z0-9_]+, got {v!r}")
    return v


class Members(BaseModel):
    model_config = ConfigDict(extra="forbid")
    members: list[str]


class Metric(BaseModel):
    model_config = ConfigDict(extra="forbid")
    id: str
    name: str
    aliases: list[str] = []
    description: str
    unit: Literal["NOK", "FTE", "count", "ratio"] = "NOK"
    sign: Literal["income", "cost", "neutral"] = "neutral"
    balance: Literal["Periodic", "Accumulated"] = "Periodic"
    aggregation: Literal["sum", "last"] = "sum"
    materiality: tuple[float, float] = (5_000_000.0, 0.10)
    status: Literal["approved", "draft"] = "draft"
    approved: dict | None = None
    finmax: Members | None = None
    finlite: Members | None = None
    finplan: Members | None = None
    differs: str = ""
    valid_from: str = ""
    valid_to: str = ""

    @field_validator("id")
    @classmethod
    def _check_id(cls, v: str) -> str:
        return _id(v)

    @field_validator("valid_from", "valid_to")
    @classmethod
    def _check_month(cls, v: str, info) -> str:
        return _month(v, info)

    @model_validator(mode="after")
    def _check(self):
        if self.balance == "Accumulated" and self.aggregation == "sum":
            raise ValueError(f"metric {self.id!r}: an Accumulated balance cannot be summed over months "
                             "(Accumulated with aggregation 'sum' double counts); use aggregation 'last'")
        if not self.members():
            raise ValueError(f"metric {self.id!r} needs members in at least one cube (finmax, finlite, finplan)")
        return self

    def members(self) -> dict[str, tuple[str, ...]]:
        """cube -> members, only the cubes the metric defines."""
        out = {}
        for cube in CUBES:
            m = getattr(self, cube.lower())
            if m is not None and m.members:
                out[cube] = tuple(m.members)
        return out


class Fact(BaseModel):
    model_config = ConfigDict(extra="forbid")
    id: str
    text: str
    scope_members: dict[str, list[str]] = {}
    scope_metrics: list[str] = []
    valid_from: str = ""
    valid_to: str = ""
    severity: Literal["info", "warn", "block"] = "info"
    source: str
    status: Literal["approved", "draft"] = "draft"
    approved: dict | None = None                 # {"by", "on"}, set by `smartass knowledge approve`

    @field_validator("id")
    @classmethod
    def _check_id(cls, v: str) -> str:
        return _id(v)

    @field_validator("valid_from", "valid_to")
    @classmethod
    def _check_month(cls, v: str, info) -> str:
        return _month(v, info)

    @field_validator("scope_members")
    @classmethod
    def _cubes(cls, v: dict[str, list[str]]):
        bad = [c for c in v if c not in CUBES]
        if bad:
            raise ValueError(f"scope_members cubes must be among {CUBES}, got {bad}")
        return v

    @model_validator(mode="after")
    def _check(self):
        if not self.scope_metrics and not any(self.scope_members.values()):
            raise ValueError(f"fact {self.id!r} needs a scope (scope_members or scope_metrics)")
        return self


# ---- product concepts (Plan 2C)

TERM = re.compile(r"\s*([+-]?)\s*(\$src|[a-z][a-z0-9_]*)\s*")


def expr_terms(expr: str) -> list[tuple[int, str]]:
    """'a-b+c' -> [(1, 'a'), (-1, 'b'), (1, 'c')]: a concept expression is item keys (and $src in a check) added or
    subtracted, nothing else."""
    out, pos = [], 0
    while pos < len(expr):
        m = TERM.match(expr, pos)
        if not m or m.end() == pos or (out and not m.group(1)):
            raise ValueError(f"expr {expr!r}: only item keys joined by + and - are allowed")
        out.append((-1 if m.group(1) == "-" else 1, m.group(2)))
        pos = m.end()
    if not out:
        raise ValueError("expr is empty")
    return out


class ConceptPick(BaseModel):
    """A cube amount (account x products at the concept entity, or account x entities with no product), or an
    expression over item keys (a key's numerator only)."""
    model_config = ConfigDict(extra="forbid")
    label: str = ""
    account: str = ""
    products: list[str] = []
    entities: list[str] = []
    expr: str = ""

    @model_validator(mode="after")
    def _check(self):
        if bool(self.account) == bool(self.expr):
            raise ValueError("give either account (with products or entities) or expr")
        if self.expr:
            expr_terms(self.expr)
            if self.products or self.entities:
                raise ValueError("an expr takes no products or entities")
        elif bool(self.products) == bool(self.entities):
            raise ValueError(f"account {self.account}: give products or entities (one of them)")
        return self


class KeyShare(BaseModel):
    """The share of a shared base that belongs to the concept: num / den, or the same key as another item."""
    model_config = ConfigDict(extra="forbid")
    text: str = ""
    num: ConceptPick | None = None
    den: ConceptPick | None = None
    same_as: str = ""

    @model_validator(mode="after")
    def _check(self):
        if self.same_as:
            if self.num or self.den:
                raise ValueError("key_share: same_as takes no num/den")
        elif not (self.num and self.den):
            raise ValueError("key_share needs num and den (or same_as)")
        elif not self.den.account:
            raise ValueError("key_share.den must be a cube amount (account), not an expr")
        return self


class ConceptItem(BaseModel):
    """booked: tagged to the concept in the cube; attributed: not tagged, belongs to it in full (or by key_share);
    estimated: base x key_share; sum: a subtotal of other items; missing: not in the cube (a note, no value)."""
    model_config = ConfigDict(extra="forbid")
    key: str
    label: str
    kind: Literal["booked", "attributed", "estimated", "sum", "missing"]
    account: str = ""
    products: list[str] = []
    entities: list[str] = []
    base: ConceptPick | None = None
    key_share: KeyShare | None = None
    of: list[str] = []
    note: str = ""

    @field_validator("key")
    @classmethod
    def _check_key(cls, v: str) -> str:
        if not re.fullmatch(r"[a-z][a-z0-9_]*", v) or v == "total":
            raise ValueError(f"item key must match [a-z][a-z0-9_]* and not be 'total', got {v!r}")
        return v

    @model_validator(mode="after")
    def _check(self):
        k = self.kind
        if k in ("booked", "attributed"):
            ConceptPick(account=self.account, products=self.products, entities=self.entities)
        if k == "booked" and self.key_share:
            raise ValueError(f"item {self.key}: a booked item has no key_share")
        if k == "estimated" and (self.base is None or not self.base.account or self.key_share is None):
            raise ValueError(f"item {self.key}: estimated needs base (a cube amount) and key_share")
        if k == "sum" and not self.of:
            raise ValueError(f"item {self.key}: sum needs of")
        if k == "missing" and (self.account or self.base or self.key_share):
            raise ValueError(f"item {self.key}: a missing item has no account, base or key_share (only a note)")
        return self

    def pick(self) -> ConceptPick | None:
        """The cube amount of a booked or attributed item (an estimated item's is its base)."""
        if self.kind in ("booked", "attributed"):
            return ConceptPick(label=self.label, account=self.account, products=self.products,
                               entities=self.entities)
        return self.base if self.kind == "estimated" else None


class ConceptBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")
    title: str
    item: list[ConceptItem] = Field(min_length=1)


class ConceptCheck(BaseModel):
    model_config = ConfigDict(extra="forbid")
    label: str
    expr: str
    source: ConceptPick | None = None


class ConceptTotal(BaseModel):
    model_config = ConfigDict(extra="forbid")
    label: str
    of: list[str] = Field(min_length=1)


class Concept(BaseModel):
    model_config = ConfigDict(extra="forbid")
    id: str
    title: str
    names: list[str] = []
    kind: Literal["product_offer", "product_family"]
    description: str
    entity: str = "un.SY"
    source: str = ""
    status: Literal["approved", "draft"] = "draft"
    approved: dict | None = None
    notes: list[str] = []
    decisions: dict = {}
    total: ConceptTotal
    check: list[ConceptCheck] = []
    block: list[ConceptBlock] = Field(min_length=1)

    @field_validator("id")
    @classmethod
    def _check_id(cls, v: str) -> str:
        return _id(v)

    def items(self) -> list[ConceptItem]:
        return [it for b in self.block for it in b.item]

    @model_validator(mode="after")
    def _refs(self):
        keys: list[str] = [it.key for it in self.items()]
        dup = sorted({k for k in keys if keys.count(k) > 1})
        if dup:
            raise ValueError(f"concept {self.id}: duplicate item key(s) {dup}")
        by = {it.key: it for it in self.items()}

        def known(names, where, extra=()):
            bad = [n for n in names if n not in by and n not in extra]
            if bad:
                raise ValueError(f"concept {self.id}: {where} names unknown item(s) {bad}")
        for it in self.items():
            known(it.of, f"sum {it.key}")
            ks = it.key_share
            if ks and ks.same_as:
                other = by.get(ks.same_as)
                if other is None or other.key_share is None or other.key_share.same_as:
                    raise ValueError(f"concept {self.id}: item {it.key} key_share.same_as must name an item with "
                                     "its own num/den key")
            if ks and ks.num and ks.num.expr:
                known([t for _, t in expr_terms(ks.num.expr)], f"key of {it.key}")
        known(self.total.of, "total")
        for ch in self.check:
            terms = [t for _, t in expr_terms(ch.expr)]
            if "$src" in terms and ch.source is None:
                raise ValueError(f"concept {self.id}: check {ch.label!r} uses $src but has no source")
            known(terms, f"check {ch.label!r}", ("total", "$src"))
        return self

    def picks(self) -> list[ConceptPick]:
        """Every cube amount the concept reads: items, bases, key denominators and check sources."""
        out = [p for it in self.items() if (p := it.pick()) is not None]
        out += [it.key_share.den for it in self.items() if it.key_share and it.key_share.den]
        out += [it.key_share.num for it in self.items() if it.key_share and it.key_share.num
                and it.key_share.num.account]
        out += [ch.source for ch in self.check if ch.source]
        return out

    def members(self) -> dict[str, list[str]]:
        """The cube members the concept names (accounts, products, entities), for the atlas checks."""
        out: dict[str, list[str]] = {"account": [], "product": [], "entity": [self.entity]}
        for p in self.picks():
            out["account"].append(p.account)
            out["product"] += p.products
            out["entity"] += p.entities
        return {k: list(dict.fromkeys(v)) for k, v in out.items()}


def concept_terms(c: Concept) -> list[str]:
    """What a concept is found by: its id (also with spaces), its title and its names, without case duplicates."""
    out, seen = [], set()
    for t in (c.id, c.id.replace("_", " "), c.title, *c.names):
        t = t.strip()
        if t and t.casefold() not in seen:
            seen.add(t.casefold())
            out.append(t)
    return out


def _scalar(v) -> str:
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return repr(v)
    if isinstance(v, str):
        return json.dumps(v, ensure_ascii=False)          # a JSON string is a valid TOML basic string
    if isinstance(v, (list, tuple)):
        return "[" + ", ".join(_scalar(x) for x in v) + "]"
    if isinstance(v, dict):
        return "{ " + ", ".join(f"{_key(k)} = {_scalar(x)}" for k, x in v.items()) + " }"
    raise TypeError(f"cannot write {type(v).__name__} to TOML")


def _key(k: str) -> str:
    return k if re.fullmatch(r"[A-Za-z0-9_-]+", k) else json.dumps(k, ensure_ascii=False)


def _array(v: list) -> str:
    """A string array; one item per line when long (git diffs stay readable)."""
    one = _scalar(v)
    if len(one) <= 100:
        return one
    return "[\n" + "".join(f"  {_scalar(x)},\n" for x in v) + "]"


def _empty(v) -> bool:
    return v is None or v == "" or v == [] or v == {}


def _doc(data: dict, path: str = "") -> list[str]:
    """A nested dict as TOML: its own keys first, then sub-dicts as [tables] and lists of dicts as [[arrays]]."""
    data = {k: v for k, v in data.items() if not _empty(v)}
    subs = {k: v for k, v in data.items() if isinstance(v, dict)}
    arrays = {k: v for k, v in data.items() if isinstance(v, list) and v and all(isinstance(x, dict) for x in v)}
    lines = [f"{_key(k)} = {_array(v) if isinstance(v, list) else _scalar(v)}" for k, v in data.items()
             if k not in subs and k not in arrays]
    for k, v in subs.items():
        name = f"{path}{_key(k)}"
        lines += ["", f"[{name}]", *_doc(v, name + ".")]
    for k, v in arrays.items():
        name = f"{path}{_key(k)}"
        for x in v:
            lines += ["", f"[[{name}]]", *_doc(x, name + ".")]
    return lines


def to_toml(item: "Metric | Fact | Concept") -> str:
    """The item as its source file (without `id`, which is the file stem). Every set field is written, defaults
    included, so a file reads on its own; None and empty values are left out. Per-cube members (a metric's
    finmax/finlite/finplan, a fact's scope_members) are tables; a concept's total, checks, blocks and items are
    tables and arrays of tables."""
    if isinstance(item, Concept):
        return "\n".join(_doc(item.model_dump(exclude={"id"}))) + "\n"
    data = {k: v for k, v in item.model_dump(exclude={"id"}).items() if v not in (None, "", [], {})}
    if isinstance(item, Metric):
        tables = {c: data.pop(c) for c in ("finmax", "finlite", "finplan") if c in data}
    else:
        tables = {"scope_members": data.pop("scope_members")} if "scope_members" in data else {}
    lines = [f"{_key(k)} = {_array(v) if isinstance(v, list) else _scalar(v)}" for k, v in data.items()]
    for name, t in tables.items():
        lines += ["", f"[{name}]"] + [f"{_key(k)} = {_array(v)}" for k, v in t.items()]
    return "\n".join(lines) + "\n"


def _load(path: Path, cls):
    try:
        data = tomllib.loads(path.read_text(encoding="utf-8"))
        if "id" in data and data["id"] != path.stem:
            raise ValueError(f"id {data['id']!r} differs from the file name")
        return cls(**{**data, "id": path.stem})
    except (tomllib.TOMLDecodeError, ValidationError, ValueError, TypeError) as e:
        raise SemanticError(f"{path.name}: {e}") from e


def load_semantic(directory: Path) -> tuple[list[Metric], list[Fact]]:
    """Every metrics/*.toml and facts/*.toml under `directory`, validated; the id is the file stem and is unique
    across both kinds. A missing directory loads nothing."""
    directory = Path(directory)
    metrics = [_load(p, Metric) for p in sorted((directory / "metrics").glob("*.toml"))]
    facts = [_load(p, Fact) for p in sorted((directory / "facts").glob("*.toml"))]
    seen: dict[str, str] = {}
    for kind, items in (("metrics", metrics), ("facts", facts)):
        for it in items:
            if it.id in seen:
                raise SemanticError(f"duplicate id {it.id!r}: {seen[it.id]}/{it.id}.toml and {kind}/{it.id}.toml")
            seen[it.id] = kind
    return metrics, facts


def load_concepts(directory: Path) -> list[Concept]:
    """Every concepts/*.toml under `directory`, validated (the id is the file stem). A missing directory loads
    nothing."""
    return [_load(p, Concept) for p in sorted((Path(directory) / "concepts").glob("*.toml"))]
