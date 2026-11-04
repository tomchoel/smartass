"""The semantic layer's source files (sub-spec 2 §2): one TOML file per metric (`metrics/<id>.toml`) and per fact
(`facts/<id>.toml`), reviewable in git. `load_semantic` validates them; `build` compiles them with the atlas into
knowledge.db."""
from __future__ import annotations

import json
import re
import tomllib
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, ValidationError, field_validator, model_validator

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


def to_toml(item: Metric | Fact) -> str:
    """The item as its source file (without `id`, which is the file stem). Every set field is written, defaults
    included, so a file reads on its own; None and empty values are left out. Per-cube members (a metric's
    finmax/finlite/finplan, a fact's scope_members) are tables."""
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
