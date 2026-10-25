"""A Frame: one cube's warehouse values keyed by (entity, line key, period, scenario key), with the catalogue lines
attached - the calc engine's only input. Scenario keys are the catalogue's ([scenarios]: actual, actual_finplan,
budget, forecast). Flow lines come from Periodic cells, stock lines (FTE) from Accumulated cells (the month-end
level)."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Mapping

from ..catalog import Line

Key = tuple[str, str, str, str]                                # entity, line key, period, scenario key
WhKey = tuple[str, str, str, str, str, str, str]               # warehouse key (see warehouse.py)


@dataclass(frozen=True)
class Frame:
    cube: str
    lines: tuple[Line, ...]
    values: Mapping[Key, float | None]
    sources: Mapping[Key, WhKey] = field(default_factory=dict)

    def line(self, key: str) -> Line:
        for x in self.lines:
            if x.key == key:
                return x
        raise KeyError(f"{self.cube} has no line '{key}'")

    def get(self, entity: str, line: str, period: str, scenario: str) -> tuple[float, WhKey | tuple] | None:
        """(value, warehouse key) - a cell read as blank (#Missing) counts as 0 - or None when it was never read."""
        k = (entity, line, period, scenario)
        if k not in self.values:
            return None
        v = self.values[k]
        fallback = (self.cube, scenario, "", entity, self.line(line).account, period, balance_of(self.line(line)))
        return (0.0 if v is None else float(v)), self.sources.get(k, fallback)


def balance_of(line: Line) -> str:
    return "Accumulated" if line.measure == "stock" else "Periodic"


def from_rows(cube: str, lines, rows: Mapping[WhKey, float | None], scenarios: Mapping[tuple[str, str], str]) -> Frame:
    """Frame from warehouse rows. scenarios maps (scenario member, version) to a scenario key, e.g.
    {('Actual', ''): 'actual'} for FinMax or {('Target', 'Final'): 'budget'} for FinPlan; other rows are ignored."""
    by_account = {x.account: x for x in lines}
    if len(by_account) != len(lines):
        raise ValueError(f"{cube}: two catalogue lines share an account")
    mapped = list(scenarios.values())
    if len(set(mapped)) != len(mapped):
        raise ValueError(f"{cube}: two scenario members both map to one scenario key {sorted(mapped)}")
    values, sources = {}, {}
    for key, v in rows.items():
        c, scen, ver, ent, account, period, bal = key
        line = by_account.get(account)
        s = scenarios.get((scen, ver))
        if c != cube or line is None or s is None or bal != balance_of(line):
            continue
        if (ent, line.key, period, s) in values:
            raise ValueError(f"{cube}: two warehouse rows both map to {(ent, line.key, period, s)}")
        values[(ent, line.key, period, s)] = v
        sources[(ent, line.key, period, s)] = key
    return Frame(cube, tuple(lines), values, sources)
