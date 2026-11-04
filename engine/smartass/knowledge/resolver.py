"""KnowledgeResolver: the model's Resolver over knowledge.db (metrics, facts) and the atlas (tree, labels, existence),
plus the caveats for what an answer touched and search for fin_lookup (sub-spec 2 §4)."""
from __future__ import annotations

import sqlite3

from ..model.periods import Month
from ..model.resolver import DEFAULT_MATERIALITY, MetricDef, UnknownMetric
from . import atlas
from .atlas_tree import AtlasTree
from .build import search_metrics
from .semantic import Fact, Metric

SEVERITY_ORDER = {"block": 0, "warn": 1, "info": 2}
DRAFT_CAVEAT = "unapproved knowledge change touches this answer"


def _month(s: str, default: Month) -> Month:
    return (int(s[:4]), int(s[5:7])) if s else default


class KnowledgeResolver(AtlasTree):
    def __init__(self, kdb: sqlite3.Connection, atlas_db: sqlite3.Connection, forecast_year: int):
        super().__init__(atlas_db, forecast_year)
        self.kdb = kdb
        self._metrics = {mid: Metric.model_validate_json(js)
                         for mid, js in kdb.execute("SELECT id, json FROM metric ORDER BY id")}
        self._facts = {fid: Fact.model_validate_json(js) for fid, js in kdb.execute("SELECT id, json FROM fact")}
        self._lookup: dict[str, Metric] = {}
        for mid, terms in self._terms().items():
            for t in terms:
                self._lookup.setdefault(t.casefold(), self._metrics[mid])

    def _terms(self) -> dict[str, list[str]]:
        out: dict[str, list[str]] = {mid: [mid, m.name, *m.aliases] for mid, m in self._metrics.items()}
        for mid, text in self.kdb.execute("SELECT metric_id, text FROM alias"):
            out.setdefault(mid, []).append(text)
        return out

    # ---- metrics

    def _metric(self, name: str) -> Metric | None:
        return self._lookup.get(name.strip().casefold())

    def metric(self, name: str) -> Metric | None:
        """The semantic metric for an id, name or alias (any case); None for a raw member."""
        return self._metric(name)

    def metrics(self) -> list[MetricDef]:
        return [MetricDef(m.name, m.members(), sign=m.sign, status=m.status, materiality=m.materiality,
                          balance=m.balance, aggregation=m.aggregation) for m in self._metrics.values()]

    def members(self, metric: str, cube: str) -> tuple[str, ...]:
        m = self._metric(metric)
        if m is None:
            if self.exists(metric, cube):
                return (self._spelling(metric, cube),)
            raise UnknownMetric(metric)
        ms = m.members()
        if cube not in ms:
            raise UnknownMetric(f"{metric} has no {cube} definition")
        return ms[cube]

    def sign(self, metric):
        m = self._metric(metric)
        return m.sign if m else "neutral"

    def status(self, metric):
        m = self._metric(metric)
        return m.status if m else "approved"

    def materiality(self, metric):
        m = self._metric(metric)
        return tuple(m.materiality) if m else DEFAULT_MATERIALITY

    def balance(self, metric):
        m = self._metric(metric)
        return m.balance if m else "Periodic"

    def aggregation(self, metric):
        m = self._metric(metric)
        return m.aggregation if m else "sum"

    # ---- caveats

    def _in_scope(self, touched: dict[str, set[str]], metric_ids: set[str]) -> set[str]:
        ids: set[str] = set()
        for cube, members in touched.items():
            for m in members:
                ids.update(r[0] for r in self.kdb.execute(
                    "SELECT fact_id FROM fact_scope WHERE cube=? AND member=? COLLATE NOCASE", (cube, m)))
        if metric_ids:
            marks = ",".join("?" * len(metric_ids))
            ids.update(r[0] for r in self.kdb.execute(
                f"SELECT fact_id FROM fact_scope WHERE metric_id IN ({marks})", sorted(metric_ids)))
        return ids

    @staticmethod
    def _valid(f: Fact, months: set[Month] | None) -> bool:
        if not months:
            return True
        lo, hi = _month(f.valid_from, (0, 1)), _month(f.valid_to, (9999, 12))
        return any(lo <= m <= hi for m in months)

    def caveats(self, touched: dict[str, set[str]], metric_ids: set[str] = frozenset(),
                months: set[Month] | None = None) -> list[dict]:
        """{id, text, severity} for every approved fact whose scope holds a touched member (per cube, any case) or
        metric and whose validity overlaps a touched month (no months: no month filter). Draft facts in scope add one
        warn instead of their text. Sorted block, warn, info."""
        hits = [self._facts[i] for i in self._in_scope(touched, set(metric_ids)) if i in self._facts]
        hits = [f for f in hits if self._valid(f, months)]
        out = [{"id": f.id, "text": f.text, "severity": f.severity} for f in hits if f.status == "approved"]
        drafts = sorted(f.id for f in hits if f.status == "draft")
        if drafts:
            out.append({"id": "unapproved_knowledge", "severity": "warn",
                        "text": f"{DRAFT_CAVEAT} ({', '.join(drafts)}; smartass knowledge approve)."})
        return sorted(out, key=lambda c: (SEVERITY_ORDER[c["severity"]], c["id"]))

    # ---- search (fin_lookup)

    def _metric_hit(self, m: Metric) -> dict:
        return {"id": m.id, "name": m.name, "aliases": m.aliases, "description": m.description,
                "differs": m.differs, "unit": m.unit, "sign": m.sign, "status": m.status,
                "members": {c: list(ms) for c, ms in m.members().items()},
                "caveats": self.caveats({}, {m.id})}

    def search(self, text: str, kind: str = "any", cube: str | None = None, limit: int = 10) -> dict:
        """Metrics by alias or name (exact, then FTS prefix), members by name or label from the atlas; every hit
        carries its caveats (no month filter). kind: any | metric | member."""
        metrics, members = [], []
        if kind in ("any", "metric"):
            for mid in search_metrics(self.kdb, text, limit):
                m = self._metrics.get(mid)
                if m and (cube is None or cube in m.members()):
                    metrics.append(self._metric_hit(m))
        if kind in ("any", "member"):
            for h in atlas.search(self.db, text, connection=cube, limit=limit):
                members.append({"cube": h["connection"], "dimension": h["dimension"], "name": h["name"],
                                "label": h["label"], "parents": h["parents"], "depth": h["depth"],
                                "caveats": self.caveats({h["connection"]: {h["name"]}})})
        return {"metrics": metrics, "members": members}
