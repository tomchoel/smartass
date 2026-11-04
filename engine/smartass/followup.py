# engine/smartass/followup.py
"""Follow-ups on a stored v2 answer. A patch is a JSON merge patch (RFC 7386) on the stored spec: objects merge,
lists (queries, blocks) are replaced whole, null deletes a key. The question is the explicit q, else the active
'Q<n> Answer' sheet when it holds a v2 answer, else the latest v2 answer in the workbook."""
from __future__ import annotations

import copy
import re

from .model.validate import GuardError

_ANSWER = re.compile(r"^Q(\d+) Answer$")


def merge_patch(target, patch):
    if not isinstance(patch, dict):
        return copy.deepcopy(patch)
    out = copy.deepcopy(target) if isinstance(target, dict) else {}
    for k, v in patch.items():
        if v is None:
            out.pop(k, None)
        else:
            out[k] = merge_patch(out.get(k), v)
    return out


def v2_questions(stored: dict[int, dict]) -> list[int]:
    return sorted(q for q, a in stored.items() if isinstance(a.get("spec"), dict) and "v2" in a["spec"])


def find_question(stored: dict[int, dict], q: int | None, active_sheet: str | None) -> tuple[int, dict, dict | None]:
    """(q, raw v2 spec, owned range) of the answer a follow-up changes."""
    ours = v2_questions(stored)
    if q is not None:
        if q not in ours:
            raise GuardError(f"Q{q} is not a v2 answer in this workbook",
                             f"use one of {', '.join(f'Q{x}' for x in ours) or 'none'}, or ask a new question")
    else:
        m = _ANSWER.match(active_sheet or "")
        if m and int(m.group(1)) in ours:
            q = int(m.group(1))
        elif ours:
            q = ours[-1]
        else:
            raise GuardError("there is no v2 answer in this workbook to follow up",
                             "ask a new question with fin_analyze(spec)")
    a = stored[q]
    return q, a["spec"]["v2"], a.get("owned")
