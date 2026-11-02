"""Answer spec v1: the one contract the agent writes, recipes produce, follow-ups patch, recurring reports save and
the add-in will render (design §4).

The envelope (question, gate, assumptions, pulls, commentary) is strict. `layout` reuses the proven grammar of
specs.py / answer.py (title, columns, sections, lines, cells, checks, chart) and keeps unknown keys, so existing
layouts pass unchanged; validate() adds the cross-reference checks that grammar never had."""
from __future__ import annotations

import copy
import fnmatch
import json
import re
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

CELL_KINDS = ("pull", "var", "varpct", "sum", "avg", "expr", "ref", "rowsum", "value", "text", "drill")
CHECKLIST = ("period", "scenario", "entity", "measure", "currency_sign", "granularity")


class _Strict(BaseModel):
    model_config = ConfigDict(extra="forbid")


class _Open(BaseModel):
    model_config = ConfigDict(extra="allow")


# ------------------------------------------------------------------ gate
class Gate(_Strict):
    """The advisor gate's verdict. checklist: one entry per CHECKLIST item, the resolved value or the open gap."""
    outcome: Literal["clear", "ambiguous", "infeasible"]
    checklist: dict[str, str] = Field(default_factory=dict)
    questions: list[str] = Field(default_factory=list, max_length=3)
    reason: str | None = None                        # infeasible: why it cannot be answered as asked
    proxy: str | None = None                         # infeasible: what is answered instead

    @model_validator(mode="after")
    def _consistent(self) -> "Gate":
        unknown = set(self.checklist) - set(CHECKLIST)
        if unknown:
            raise ValueError(f"unknown checklist items {sorted(unknown)}; allowed: {', '.join(CHECKLIST)}")
        if self.outcome == "ambiguous" and not self.questions:
            raise ValueError("an ambiguous gate needs 1-3 clarifying questions")
        if self.outcome != "ambiguous" and self.questions:
            raise ValueError(f"a {self.outcome} gate asks no questions")
        if self.outcome == "infeasible" and not (self.reason and self.proxy):
            raise ValueError("an infeasible gate needs a reason and the proxy answered instead")
        return self


# ------------------------------------------------------------------ pulls
class Axis(_Strict):
    dimension: str
    members: list[str] = Field(min_length=1)


class Pull(_Strict):
    id: str = Field(pattern=r"^[A-Za-z_][A-Za-z0-9_]*$")
    connection: Literal["FinMax", "FinLite", "FinPlan"]
    rows: list[Axis] = Field(min_length=1)
    columns: list[Axis] = Field(min_length=1)
    pov: dict[str, str] = Field(default_factory=dict)
    title: str | None = None

    @model_validator(mode="after")
    def _dims_once(self) -> "Pull":
        dims = [a.dimension for a in self.rows + self.columns] + list(self.pov)
        twice = sorted({d for d in dims if dims.count(d) > 1})
        if twice:
            raise ValueError(f"pull {self.id}: dimension(s) used twice: {', '.join(twice)}")
        return self


# ------------------------------------------------------------------ layout (specs.py / answer.py grammar)
class Column(_Open):
    header: str = ""
    kind: str | None = None


class Line(_Open):
    key: str | None = None
    label: str | None = None
    kind: Literal["line", "sum", "grand"] | None = None
    cells: list[dict[str, Any] | None] = Field(default_factory=list)


class Section(_Open):
    title: str | None = None
    lines: list[Line] = Field(default_factory=list)
    repeat: dict[str, Any] | None = None


class Check(_Open):
    label: str
    expr: str | None = None
    static: Literal["OK", "DIFF"] | None = None      # a check computed by smartass (e.g. warehouse = sheet)
    detail: str | None = None
    columns: list[int] | None = None
    # Warning: a DIFF is a known definitional difference the user must see, not an untrustworthy number
    severity: Literal["Critical", "Warning"] = "Critical"
    @model_validator(mode="after")
    def _one(self) -> "Check":
        if (self.expr is None) == (self.static is None):
            raise ValueError(f"check '{self.label}': exactly one of expr or static")
        return self


class Layout(_Open):
    title: str
    subtitle: str | None = None
    columns: list[Column] = Field(min_length=1)
    sections: list[Section] = Field(min_length=1)
    checks: list[Check] = Field(default_factory=list)
    key_points: list[str] = Field(default_factory=list)
    notes: list[str] = Field(default_factory=list)
    source: str | None = None
    chart: dict[str, Any] | None = None


# ------------------------------------------------------------------ commentary
class Driver(_Strict):
    text: str
    basis: Literal["data", "hypothesis"]
    evidence: str | None = None                      # data: the pull/drill cell it rests on
    line: str | None = None                          # the layout line key it explains (material-variance check)

    @model_validator(mode="after")
    def _evidence(self) -> "Driver":
        if self.basis == "data" and not self.evidence:
            raise ValueError("a data-backed driver cites its evidence (pull or drill)")
        return self


class Commentary(_Strict):
    materiality: dict[str, float] = Field(default_factory=dict)
    always_comment: list[str] = Field(default_factory=list)
    drivers: list[Driver] = Field(default_factory=list)


# ------------------------------------------------------------------ recipe provenance
class Recipe(_Strict):
    name: str = Field(pattern=r"^[a-z_]+$")
    params: dict[str, Any] = Field(default_factory=dict)
    overrides: dict[str, Any] = Field(default_factory=dict)   # a follow-up-shaped patch on the draft


# ------------------------------------------------------------------ the spec
class AnswerSpec(_Strict):
    version: Literal[1] = 1
    question: str = Field(min_length=1)
    gate: Gate
    assumptions: list[str] = Field(default_factory=list)
    pulls: list[Pull] = Field(default_factory=list)
    layout: Layout | None = None
    commentary: Commentary = Field(default_factory=Commentary)
    recipe: Recipe | None = None

    @model_validator(mode="after")
    def _shape(self) -> "AnswerSpec":
        if self.gate.outcome == "ambiguous":
            if self.pulls or self.layout:
                raise ValueError("an ambiguous question pulls nothing until it is answered")
        elif not self.layout:
            raise ValueError(f"a {self.gate.outcome} answer needs a layout")
        ids = [p.id for p in self.pulls]
        dup = sorted({i for i in ids if ids.count(i) > 1})
        if dup:
            raise ValueError(f"pull ids used twice: {', '.join(dup)}")
        if self.layout:
            errs = _layout_refs(self.layout, set(ids), [d.line for d in self.commentary.drivers if d.line])
            if errs:
                raise ValueError("layout references: " + "; ".join(errs))
        elif any(d.line for d in self.commentary.drivers):
            raise ValueError("a driver names a line, but there is no layout")
        return self


def _layout_refs(layout: Layout, pull_ids: set[str], driver_lines: list[str] = ()) -> list[str]:
    """Cross-references the layout grammar cannot check by itself. Repeat-template keys ('m{i}') count as the
    lines they will expand to."""
    ncol = len(layout.columns)
    lines = [ln for sec in layout.sections for ln in sec.lines]
    keys = [ln.key for ln in lines if ln.key]
    templates = [re.sub(r"\{\w+\}", "*", str(sec.repeat.get("line", {}).get("key", "")))
                 for sec in layout.sections if sec.repeat]
    templates = [t for t in templates if t]
    known = set(keys)

    def is_key(k: str) -> bool:
        return k in known or any(fnmatch.fnmatchcase(k, t) for t in templates)

    def pattern_hits(p: str) -> bool:
        return any(fnmatch.fnmatchcase(k, p) for k in keys) or any(
            fnmatch.fnmatchcase(t.replace("*", "0"), p) for t in templates)

    errs: list[str] = []
    dup = sorted({k for k in keys if keys.count(k) > 1})
    if dup:
        errs.append(f"line keys used twice: {', '.join(dup)}")
    rep_lines = [Line.model_validate(sec.repeat.get("line", {})) for sec in layout.sections if sec.repeat]
    for ln in lines + rep_lines:
        where = ln.key or ln.label or "?"
        if len(ln.cells) > ncol:
            errs.append(f"{where}: {len(ln.cells)} cells for {ncol} columns")
        for j, c in enumerate(ln.cells):
            if c is None:
                continue
            kinds = [k for k in CELL_KINDS if k in c]
            if len(kinds) != 1:
                errs.append(f"{where}[{j}]: a cell has exactly one of {', '.join(CELL_KINDS)} (got {kinds or 'none'})")
                continue
            k, v = kinds[0], c[kinds[0]]
            if k == "pull":
                if v not in pull_ids:
                    errs.append(f"{where}[{j}]: unknown pull '{v}'")
                if not c.get("row") or not c.get("col"):
                    errs.append(f"{where}[{j}]: a pull cell needs row and col")
            elif k in ("var", "varpct", "rowsum"):
                idx = list(v) if isinstance(v, list) else []
                if (k != "rowsum" and len(idx) != 2) or not idx or \
                        any(not isinstance(i, int) or not 0 <= i < ncol for i in idx):
                    errs.append(f"{where}[{j}]: {k} needs column indexes within 0..{ncol - 1}")
            elif k in ("sum", "avg"):
                if isinstance(v, str):
                    if not pattern_hits(v):
                        errs.append(f"{where}[{j}]: {k} pattern '{v}' matches no line")
                elif isinstance(v, list):
                    errs += [f"{where}[{j}]: {k} of unknown line '{x}'" for x in v if not is_key(x)]
            elif k == "ref" and not is_key(str((v or {}).get("key", ""))):
                errs.append(f"{where}[{j}]: ref to unknown line")
            elif k == "expr":
                errs += [f"{where}[{j}]: expr uses unknown line '{x}'"
                         for x in re.findall(r"[A-Za-z_]\w*", str(v)) if not is_key(x)]
    for ch in layout.checks:
        if ch.expr is not None:
            errs += [f"check '{ch.label}': unknown line '{x}'"
                     for x in re.findall(r"[A-Za-z_]\w*", ch.expr) if not is_key(x)]
        if ch.columns and any(not 0 <= i < ncol for i in ch.columns):
            errs.append(f"check '{ch.label}': column outside 0..{ncol - 1}")
    errs += [f"commentary driver of unknown line '{x}'" for x in driver_lines if not is_key(x)]
    return errs


# ------------------------------------------------------------------ helpers
def validate(spec: dict) -> AnswerSpec:
    """Parse and check a spec dict; pydantic.ValidationError lists the problems."""
    return AnswerSpec.model_validate(spec)


_IDS = {"pulls": "id", "sections": "title", "lines": "key", "checks": "label"}


def _merge_list(old: list, new, ident: str | None, where: str) -> list:
    if isinstance(new, dict) and set(new) == {"append"}:
        return list(old) + copy.deepcopy(list(new["append"]))
    if ident is None or not isinstance(new, list):
        return copy.deepcopy(new)
    out = copy.deepcopy(list(old))
    for item in new:
        k = item.get(ident) if isinstance(item, dict) else None
        if k is None:
            out.append(copy.deepcopy(item))
            continue
        hits = [i for i, x in enumerate(out) if isinstance(x, dict) and x.get(ident) == k]
        if len(hits) > 1:
            raise ValueError(f"{where}: {ident} '{k}' is used twice; the patch cannot tell which one")
        if item.get("drop"):
            if hits:
                out.pop(hits[0])
            continue
        if hits:
            out[hits[0]] = _merge(out[hits[0]], item, where)
        else:
            out.append(copy.deepcopy(item))
    return out


def _merge(base: dict, fu: dict, where: str = "spec") -> dict:
    """specs.merge (dicts merge, None deletes, lists replace) plus identity merges for _IDS lists and
    {"append": [...]}."""
    out = copy.deepcopy(base)
    for k, v in fu.items():
        if v is None:
            out.pop(k, None)
        elif isinstance(out.get(k), list) or (isinstance(v, dict) and set(v) == {"append"}):
            out[k] = _merge_list(out.get(k) or [], v, _IDS.get(k), f"{where}.{k}")
        elif isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge(out[k], v, f"{where}.{k}")
        else:
            out[k] = copy.deepcopy(v)
    return out


def patched(spec: dict, follow_up: dict) -> dict:
    """The follow-up applied to a stored spec (raw dict, validated). Dicts merge, None deletes; pulls merge by id,
    layout.sections by title, lines by key, layout.checks by label ({..., "drop": true} removes; unmatched items
    are appended); {"append": [...]} appends to any list; any other list replaces."""
    if (follow_up.get("recipe") or {}).get("params") is not None:
        raise ValueError("recipe parameters cannot be patched: ask it as a new question, or patch pulls/layout")
    out = _merge(spec, follow_up)
    validate(out)
    return out


def patch(spec: dict, follow_up: dict) -> AnswerSpec:
    """patched() as a validated model."""
    return validate(patched(spec, follow_up))


def outline(raw: dict) -> dict:
    """What a follow-up needs to address parts of an answer, without the full spec."""
    lay = raw.get("layout") or {}
    return {"recipe": (raw.get("recipe") or {}).get("name"),
            "pulls": [{"id": p["id"], "connection": p["connection"]} for p in raw.get("pulls") or []],
            "columns": [c.get("header", "") for c in lay.get("columns") or []],
            "sections": [{"title": s.get("title"), "lines": [ln.get("key") for ln in s.get("lines") or []],
                          **({"repeat": (s["repeat"].get("line") or {}).get("key")} if s.get("repeat") else {})}
                         for s in lay.get("sections") or []],
            "checks": [c["label"] for c in lay.get("checks") or []]}


def json_schema() -> str:
    return json.dumps(AnswerSpec.model_json_schema(), indent=2)
