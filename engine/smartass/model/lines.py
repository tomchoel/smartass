"""Ledger lines behind one cube cell (a FinLite drill), as one set per lines block. Written to Excel as table
Lines_Q<n>_<alias> on sheet 'Q<n> Lines <alias>' (v1 drill_xl.write_lines), with the cube value in D2."""
from __future__ import annotations

from dataclasses import dataclass, field

AMOUNT = "AMOUNT"
CUBE_CELL = "$D$2"


@dataclass
class LineSet:
    alias: str
    rows: list[dict]
    columns: list[str]
    cube_value: float | None = None
    reconciliation: dict = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


def norm(v) -> str:
    return str(v or "").strip().casefold()
