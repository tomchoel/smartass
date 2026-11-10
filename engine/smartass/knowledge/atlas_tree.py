"""The atlas half of a Resolver: member existence, spelling, tree, labels and the latest forecast, read from v1's
atlas.db (every member of the three cubes). The consolidation operators are behaviour cards in knowledge.db: this
class weighs every child 1.0, KnowledgeResolver overrides weight() with the cards."""
from __future__ import annotations

import re
import sqlite3

from ..model.resolver import UnknownMetric


class AtlasTree:
    def __init__(self, db: sqlite3.Connection, forecast_year: int):
        self.db = db
        self.forecast_year = forecast_year

    def _row(self, member: str, cube: str):
        return self.db.execute("SELECT name, dimension, label FROM member WHERE connection=? AND name=? "
                               "COLLATE NOCASE ORDER BY depth LIMIT 1", (cube, member)).fetchone()

    def canonical(self, member: str, cube: str) -> str:
        return self._spelling(member, cube)

    def _spelling(self, member: str, cube: str) -> str:
        row = self._row(member, cube)
        return row[0] if row else member

    def exists(self, member: str, cube: str) -> bool:
        return self._row(member, cube) is not None

    def children(self, member: str, cube: str) -> tuple[str, ...]:
        row = self._row(member, cube)
        if row is None:
            return ()
        names = [r[0] for r in self.db.execute(
            "SELECT name FROM member WHERE connection=? AND dimension=? AND parent=? ORDER BY ord",
            (cube, row[1], row[0]))]
        return tuple(dict.fromkeys(names))

    def parent(self, member: str, cube: str) -> str | None:
        row = self.db.execute("SELECT parent FROM member WHERE connection=? AND name=? COLLATE NOCASE "
                              "ORDER BY depth LIMIT 1", (cube, member)).fetchone()
        return row[0] if row and row[0] else None

    def weight(self, member: str, cube: str, parent: str | None = None) -> float:
        return 1.0

    def label(self, member: str, cube: str) -> str:
        row = self._row(member, cube)
        if not row:
            return member
        return re.sub(r"\s*\([^()]*\)\s*$", "", row[2] or row[0]) or row[0]

    def latest_forecast(self) -> str:
        names = [r[0] for r in self.db.execute("SELECT DISTINCT name FROM member WHERE connection='FinPlan' AND "
                                               "dimension='Scenario' AND name LIKE 'Forecast Q%'")]
        cands = sorted(((int(m.group(1)), n) for n in names
                        if (m := re.fullmatch(r"Forecast Q(\d)(\d\d)", n)) and int(m.group(2)) == self.forecast_year),
                       reverse=True)
        if not cands:
            raise UnknownMetric(f"no Forecast Q<q>{self.forecast_year:02d} in the atlas")
        return cands[0][1]
