"""Answer spec v1: the one contract the agent writes, recipes produce, follow-ups patch, recurring reports save and
the add-in will render (design §4).

The envelope (question, gate, assumptions, pulls, commentary) is strict. `layout` reuses the proven grammar of
specs.py / answer.py (title, columns, sections, lines, cells, checks, chart) and keeps unknown keys, so existing
layouts pass unchanged; validate() adds the cross-reference checks that grammar never had."""
from __future__ import annotations

import fnmatch
import json
import re
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from ..specs import merge

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
    expr: str
    columns: list[int] | None = None


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

    @model_validator(mode="after")
    def _evidence(self) -> "Driver":
        if self.basis == "data" and not self.evidence:
            raise ValueError("a data-backed driver cites its evidence (pull or drill)")
        return self


class Commentary(_Strict):
    materiality: dict[str, float] = Field(default_factory=dict)
    always_comment: list[str] = Field(default_factory=list)
    drivers: list[Driver] = Field(default_factory=list)


# ------------------------------------------------------------------ the spec
class AnswerSpec(_Strict):
    version: Literal[1] = 1
    question: str = Field(min_length=1)
    gate: Gate
    assumptions: list[str] = Field(default_factory=list)
    pulls: list[Pull] = Field(default_factory=list)
    layout: Layout | None = None
    commentary: Commentary = Field(default_factory=Commentary)

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
            errs = _layout_refs(self.layout, set(ids))
            if errs:
                raise ValueError("layout references: " + "; ".join(errs))
        return self


def _layout_refs(layout: Layout, pull_ids: set[str]) -> list[str]:
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
        errs += [f"check '{ch.label}': unknown line '{x}'"
                 for x in re.findall(r"[A-Za-z_]\w*", ch.expr) if not is_key(x)]
        if ch.columns and any(not 0 <= i < ncol for i in ch.columns):
            errs.append(f"check '{ch.label}': column outside 0..{ncol - 1}")
    return errs


# ------------------------------------------------------------------ helpers
def validate(spec: dict) -> AnswerSpec:
    """Parse and check a spec dict; pydantic.ValidationError lists the problems."""
    return AnswerSpec.model_validate(spec)


def patch(spec: dict, follow_up: dict) -> AnswerSpec:
    """A follow-up: deep-merge (specs.merge: dicts merge, lists replace, None deletes) and re-validate.
    `pulls` merge by id, so a follow-up adds or replaces one pull ({"id": .., "drop": true} removes it)
    without restating the rest."""
    fu = dict(follow_up)
    new_pulls = fu.pop("pulls", None)
    out = merge(spec, fu)
    if new_pulls is not None:
        by_id = {p["id"]: p for p in out.get("pulls", [])}
        for p in new_pulls:
            if p.get("drop"):
                by_id.pop(p["id"], None)
            else:
                by_id[p["id"]] = p
        out["pulls"] = list(by_id.values())
    return validate(out)


def json_schema() -> str:
    return json.dumps(AnswerSpec.model_json_schema(), indent=2)
