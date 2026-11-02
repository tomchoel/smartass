"""Follow-ups (design §3.3): which answer a follow-up patches and which source sheets it can keep. A pull whose
definition did not change keeps its sheet (no re-fetch); a changed one is fetched again under the same name; a removed
one's sheet is deleted after the redraw succeeded."""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .. import userexcel
from . import contract

_ANSWER = re.compile(r"^Q(\d+) Answer$")


@dataclass
class SourcePlan:
    reuse: dict[str, str] = field(default_factory=dict)      # pull id -> existing sheet
    fetch: list[str] = field(default_factory=list)           # pull ids to pull, in spec order
    replace: dict[str, str] = field(default_factory=dict)    # changed pull id -> its old sheet (parked, then dropped)
    drop: list[str] = field(default_factory=list)            # sheets of removed pulls


def _norm(p: dict) -> dict:
    """The pull's normalised definition; a stored pull that no longer validates never equals a new one."""
    try:
        return contract.Pull.model_validate(p).model_dump()
    except ValueError:                                       # pydantic ValidationError is a ValueError
        return {"_invalid": p}


def plan_sources(old: dict, old_ids: dict[str, str], new: dict, exists) -> SourcePlan:
    before = {p["id"]: _norm(p) for p in old.get("pulls") or []}
    plan = SourcePlan()
    for p in new.get("pulls") or []:
        pid, sheet = p["id"], old_ids.get(p["id"])
        if sheet and before.get(pid) == _norm(p) and exists(sheet):
            plan.reuse[pid] = sheet
            continue
        plan.fetch.append(pid)
        if sheet and pid in before and exists(sheet):
            plan.replace[pid] = sheet
    kept = {p["id"] for p in new.get("pulls") or []}
    plan.drop = [s for pid, s in old_ids.items() if pid not in kept and exists(s)]
    return plan


def drop_stale_expect(layout: dict, changed: list[str]) -> dict:
    """layout.expect without the expectations of pulls that are fetched again (their values may have moved)."""
    exp = layout.get("expect")
    if not exp:
        return layout
    cells = {k: v for k, v in (exp.get("cells") or {}).items() if k not in changed}
    out = {k: v for k, v in layout.items() if k != "expect"}
    return {**out, "expect": {**exp, "cells": cells}} if cells else out


def _active_sheet(t) -> str:
    return str(userexcel.com(lambda: t.wb.Application.ActiveSheet.Name))


def current_question(t) -> int | None:
    """The answer a follow-up edits: the active Q<n> Answer sheet if it has a stored v1 spec, else the latest one."""
    stored = userexcel.stored_questions(t)
    m = _ANSWER.match(_active_sheet(t))
    if m and int(m.group(1)) in stored:
        return int(m.group(1))
    return stored[-1] if stored else None
