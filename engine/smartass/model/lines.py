"""Ledger lines behind one cube cell (a FinLite drill), as one set per lines block. Written to Excel as table
Lines_Q<n>_<alias> on sheet 'Q<n> Lines <alias>' (v1 drill_xl.write_lines), with the cube value in D2."""
from __future__ import annotations

from dataclasses import dataclass, field

AMOUNT = "AMOUNT"
CUBE_CELL = "$D$2"
INITIATIVE, DERIVED = "INITIATIVE", "INITIATIVE_DERIVED"
NO_DESC = "(no description)"


def add_initiative(rows: list[dict], columns: list[str] | None = None) -> list[str]:
    """Mark each line's initiative (in place): INITIATIVE = the booked project (PROJECT_DESC, else PROJECT) when
    PROJECT is filled and not 000000; otherwise INITIATIVE_DERIVED = the cost centre (CC_DESC), else the first 40
    characters of the line text without number-only words. Exactly one of the two is filled. Returns the columns
    with both added."""
    for r in rows:
        proj = str(r.get("PROJECT") or "").strip()
        if proj and proj.strip("0"):
            r[INITIATIVE], r[DERIVED] = str(r.get("PROJECT_DESC") or "").strip() or proj, ""
        else:
            text = " ".join(w for w in str(r.get("GL_DESCRIPTION_LINE") or "").split() if not w.isdigit())[:40]
            r[INITIATIVE], r[DERIVED] = "", str(r.get("CC_DESC") or "").strip() or text.strip() or NO_DESC
    cols = list(columns or [])
    return cols + [c for c in (INITIATIVE, DERIVED) if c not in cols]


@dataclass
class LineSet:
    alias: str
    rows: list[dict]
    columns: list[str]
    cube_value: float | None = None
    reconciliation: dict = field(default_factory=dict)
    cell: dict = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


def norm(v) -> str:
    return str(v or "").strip().casefold()
