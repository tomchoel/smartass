"""KnowledgeResolver: the model's Resolver over knowledge.db (metrics, facts) and the atlas (tree, labels, existence),
plus the caveats for what an answer touched and search for fin_lookup (sub-spec 2 §4)."""
from __future__ import annotations

import contextlib
import json
import re
import sqlite3
import threading
from pathlib import Path

from ..model.periods import Month
from ..model.resolver import MetricDef, UnknownMetric, metric_of_member, pct_floor_of
from . import atlas
from .atlas_tree import AtlasTree
from .build import atlas_stamp, build, card_targets, search_concepts, search_metrics, semantic_hash, summary
from .semantic import Concept, Entity, Fact, Kpi, Metric, Rule, Vocab

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
        self._concepts = {cid: Concept.model_validate_json(js)
                          for cid, js in kdb.execute("SELECT id, json FROM concept ORDER BY id")}
        self._concept_lookup = {t.casefold(): self._concepts[cid]
                                for cid, t in kdb.execute("SELECT concept_id, text FROM concept_alias")
                                if cid in self._concepts}
        self._rules = {i: Rule.model_validate_json(js) for i, js in kdb.execute("SELECT id, json FROM rule ORDER BY id")}
        self._entities = {i: Entity.model_validate_json(js) for i, js in kdb.execute("SELECT id, json FROM entity")}
        self._kpis = {i: Kpi.model_validate_json(js) for i, js in kdb.execute("SELECT id, json FROM kpi")}
        self._vocab = {i: Vocab.model_validate_json(js) for i, js in kdb.execute("SELECT id, json FROM vocab")}
        self._levels = {(cube, m.casefold()) for cube, m in kdb.execute(
            "SELECT DISTINCT cube, member FROM card WHERE time_balance IN ('balance_last', 'average')")}
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
                          balance=m.balance, aggregation=m.aggregation, unit=m.unit, pct_floor=m.pct_floor)
                for m in self._metrics.values()]

    def metric_of(self, member: str, cube: str) -> str | None:
        """The metric of an account: the one whose members in the cube are exactly it, else its parent's, else None."""
        return metric_of_member(self.metrics(), member, cube, self.parent(member, cube))

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
        if m:
            return tuple(m.materiality)
        d = self._required("materiality").data           # a raw member takes the rule
        return (float(d["abs_nok"]), float(d["pct"]))

    def pct_floor(self, metric):
        """|base| below it: no variance %. A raw member (no semantic metric) has no floor: its unit is unknown."""
        m = self._metric(metric)
        return pct_floor_of(m.unit, tuple(m.materiality), m.pct_floor) if m else None

    def _member_metric(self, metric: str) -> Metric | None:
        key = metric.strip().casefold()
        for m in self._metrics.values():
            for ms in m.members().values():
                if len(ms) == 1 and ms[0].casefold() == key:
                    return m
        return None

    def balance(self, metric):
        m = self._metric(metric) or self._member_metric(metric)
        if m:
            return m.balance
        for cube in ("FinMax", "FinPlan"):
            card = self.card(metric, cube)
            if card and card.get("time_balance") == "balance_last":
                return "Accumulated"
        return "Periodic"

    def aggregation(self, metric):
        m = self._metric(metric) or self._member_metric(metric)
        if m:
            return m.aggregation
        for cube in ("FinMax", "FinPlan"):
            card = self.card(metric, cube)
            if card and card.get("time_balance") in ("balance_last", "average"):
                return "last"
        return "sum"

    # ---- concepts

    def concept(self, name: str) -> Concept | None:
        """The product concept for an id, title or name (any case); None when there is none."""
        return self._concepts.get(name.strip()) or self._concept_lookup.get(name.strip().casefold())

    # ---- rules, entities, KPIs, vocabulary, member notes

    def scenarios(self) -> dict:
        return self._required("scenarios").data

    def _matches(self, text: str) -> dict[str, list[str]]:
        """{kind: item ids} the phrase names, by the alias rules (catalog header, spec §4): terms match as whole
        words; one shorter than 3 characters only as the whole phrase or written in UPPER CASE; a generic alias
        ('total') only when no other entity matched; a match inside a longer match (any kind) is dropped."""
        orig = re.findall(r"[\w&%]+", text)
        words = [w.casefold() for w in orig]
        found: list[tuple[int, int, str, str, bool]] = []      # start, end, kind, id, generic
        if words:
            for term, kind, item in self.kdb.execute("SELECT text, kind, item_id FROM term "
                                                     "WHERE kind IN ('entity','entity_generic','kpi','vocab')"):
                tw = re.findall(r"[\w&%]+", term.casefold())
                n = len(tw)
                short = len(term.strip()) < 3
                for i in range(len(words) - n + 1):
                    if words[i:i + n] != tw:
                        continue
                    if short and not (n == len(words) or [w.upper() for w in tw] == orig[i:i + n]):
                        continue
                    found.append((i, i + n, kind.removesuffix("_generic"), item, kind.endswith("_generic")))
        found = [f for f in found
                 if not any(g[1] - g[0] > f[1] - f[0] and g[0] <= f[0] and f[1] <= g[1] for g in found)]
        if any(f[2] == "entity" and not f[4] for f in found):
            found = [f for f in found if not f[4]]
        out: dict[str, list[str]] = {"entity": [], "kpi": [], "vocab": []}
        for _, _, kind, item, _ in sorted(found, key=lambda f: (f[0], f[0] - f[1])):
            if item not in out[kind]:
                out[kind].append(item)
        return out

    def rule(self, rid: str):
        return self._rules.get(rid)

    def rule_of_kind(self, kind: str):
        """The approved rule of a kind (a draft is not in force), else None."""
        return next((r for r in self._rules.values() if r.kind == kind and r.status == "approved"), None)

    def _required(self, kind: str):
        r = self.rule_of_kind(kind)
        if r is None:
            raise LookupError(f"no approved rule of kind {kind!r} in knowledge/semantic/rules")
        return r

    def _entity_hit(self, e) -> dict:
        out = {"id": e.id, "label": e.label, "aliases": e.aliases, "members": e.members(), "note": e.note,
               "abolished": e.abolished, "successors": e.successors, "status": e.status,
               "caveats": self.caveats({c: {m} for c, m in e.members().items()})}
        if e.finplan_total:
            out["finplan_total"] = e.finplan_total
            out["finplan_total_children"] = list(self.children(e.finplan_total, "FinPlan"))
        return out

    def _kpi_hit(self, k) -> dict:
        return {**k.model_dump(exclude={"source", "approved", "verified"}),
                "caveats": self.caveats({}, {k.numerator, k.denominator})}

    def note(self, member: str, cube: str) -> str:
        cube = "FinMax" if cube == "FinLite" else cube        # FinLite shares the FinMax members
        row = self.kdb.execute("SELECT text FROM member_note WHERE cube=? AND member=? COLLATE NOCASE",
                               (cube, member)).fetchone()
        return row[0] if row else ""

    # ---- behaviour cards

    def card(self, member: str, cube: str, parent: str | None = None) -> dict | None:
        """The merged behaviour card of a member under a parent (derived, with manual fields over it); without a
        parent the shallowest placement's. None when there is no card. {operator, time_balance, account_type,
        storage, shared, note, provenance}; unknown fields are None."""
        cols = "SELECT operator, time_balance, account_type, storage, shared, note, provenance FROM card "
        if parent is None:        # no atlas call: the placement that has an operator, else the first by parent
            row = self.kdb.execute(cols + "WHERE cube=? AND member=? COLLATE NOCASE "
                                   "ORDER BY operator IS NULL, parent LIMIT 1", (cube, member)).fetchone()
        else:
            row = self.kdb.execute(cols + "WHERE cube=? AND member=? COLLATE NOCASE AND parent=? COLLATE NOCASE "
                                   "ORDER BY dimension LIMIT 1", (cube, member, parent)).fetchone()
        if row is None:
            return None
        op, tb, at, st, shared, note, prov = row
        return {"operator": op, "time_balance": tb, "account_type": at, "storage": st, "shared": bool(shared),
                "note": note or "", "provenance": json.loads(prov or "{}")}

    def compact_card(self, member: str, cube: str) -> dict:
        """The card a member hit carries: its time balance and account type, and which of its own children the
        consolidation ignores ('~') or subtracts ('-'). Fields that are not known are left out; {} for no card."""
        card = self.card(member, cube) or {}
        out = {k: card[k] for k in ("time_balance", "account_type") if card.get(k)}
        for key, op in (("children_ignored", "~"), ("children_subtracted", "-")):
            kids = [r[0] for r in self.kdb.execute(
                "SELECT member FROM card WHERE cube=? AND parent=? COLLATE NOCASE AND operator=? "
                "ORDER BY member", (cube, member, op))]
            if kids:
                out[key] = list(dict.fromkeys(kids))
        return out

    GROUP_OVER = 3

    def card_caveats(self, touched: dict[str, set[str]]) -> list[dict]:
        """Info caveats for touched accounts that are a level (time balance last) or an average, unless a metric with
        aggregation 'last' already covers them (FinMax also covers FinLite). Up to 3: one caveat each, id
        card_<member>; more: one grouped caveat card_levels naming the first 3. Only the members that have such a
        card are looked at (no atlas call)."""
        covered = {(c, x.casefold()) for m in self._metrics.values() if m.aggregation == "last"
                   for cube, ms in m.members().items() for c in card_targets(cube) for x in ms}
        found: dict[str, tuple[str, str]] = {}
        for cube, members in touched.items():
            for member in sorted(m for m in members if (cube, m.casefold()) in self._levels
                                 and (cube, m.casefold()) not in covered):
                card = self.card(member, cube)
                if card and card["time_balance"] in ("balance_last", "average"):
                    found.setdefault(f"card_{member}", (self.label(member, cube), card["time_balance"]))
        if len(found) > self.GROUP_OVER:
            first = ", ".join(lab for lab, _ in list(found.values())[:3])
            return [{"id": "card_levels", "severity": "info", "text": (
                f"{len(found)} accounts are levels or averages: {first} and {len(found) - 3} more. Read the "
                "period-end value of a level (or the period's average), never sum months.")}]
        return [{"id": cid, "severity": "info", "text": (
            f"{lab} is a level (time balance last): read the period-end value, never sum months."
            if tb == "balance_last" else
            f"{lab} is an average (time balance average): read the period's average, never sum months.")}
            for cid, (lab, tb) in sorted(found.items())]

    OPERATOR_WEIGHT = {"+": 1.0, "-": -1.0, "~": 0.0}

    def weight(self, member: str, cube: str, parent: str | None = None) -> float:
        """Consolidation weight of a member in its parent from the behaviour card: '+' 1.0, '-' -1.0, '~' 0.0; 1.0
        when there is no card or its operator is unknown. Without a parent the shallowest placement's."""
        card = self.card(member, cube, parent)
        return self.OPERATOR_WEIGHT.get(card["operator"], 1.0) if card else 1.0

    # ---- caveats

    def _in_scope(self, touched: dict[str, set[str]], metric_ids: set[str]) -> dict[str, set[str]]:
        """{fact id: the members it came from when it reached a touched member by inheritance}."""
        ids: dict[str, set[str]] = {}
        direct: set[str] = set()
        for cube, members in touched.items():
            hit = {m.casefold() for m in members}
            for m in members:
                for fid, member, via in self.kdb.execute(
                        "SELECT fact_id, member, via FROM fact_scope WHERE cube=? AND member=? COLLATE NOCASE",
                        (cube, m)):
                    got = ids.setdefault(fid, set())
                    if not via or via.casefold() == member.casefold():
                        direct.add(fid)
                    elif via.casefold() not in hit:
                        got.add(via)
        for fid in direct:
            ids[fid] = set()
        if metric_ids:
            marks = ",".join("?" * len(metric_ids))
            for (fid,) in self.kdb.execute(f"SELECT fact_id FROM fact_scope WHERE metric_id IN ({marks})",
                                           sorted(metric_ids)):
                ids.setdefault(fid, set())
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
        scope = self._in_scope(touched, set(metric_ids))
        hits = [self._facts[i] for i in scope if i in self._facts]
        hits = [f for f in hits if self._valid(f, months)]
        out = [{"id": f.id, "text": f.text, "severity": f.severity, **({"via": sorted(scope[f.id])} if scope[f.id]
                                                                        else {})}
               for f in hits if f.status == "approved"]
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

    @staticmethod
    def _concept_hit(c: Concept) -> dict:
        return {"id": c.id, "title": c.title, "kind": c.kind, "description": c.description, "names": c.names,
                "entity": c.entity, "status": c.status, "notes": c.notes, "total": c.total.label,
                "checks": [x.label for x in c.check],
                "blocks": [{"title": b.title, "items": [{"key": it.key, "label": it.label, "kind": it.kind}
                                                       for it in b.item]} for b in c.block]}

    def search(self, text: str, kind: str = "any", cube: str | None = None, limit: int = 10) -> dict:
        """Metrics by alias or name (exact, then FTS prefix), product concepts the same way, members by name or label
        from the atlas; every metric and member hit carries its caveats (no month filter).
        kind: any | metric | member | concept | entity | kpi | vocab | rule (rule: only by kind; '' lists all)."""
        metrics, members, concepts = [], [], []
        entities, kpis, vocab, rules = [], [], [], []
        if kind in ("any", "concept") and cube in (None, "FinMax", "FinLite"):
            concepts = [self._concept_hit(self._concepts[cid]) for cid in search_concepts(self.kdb, text, limit)
                        if cid in self._concepts]
        if kind in ("any", "metric"):
            for mid in search_metrics(self.kdb, text, limit):
                m = self._metrics.get(mid)
                if m and (cube is None or cube in m.members()):
                    metrics.append(self._metric_hit(m))
        if kind in ("any", "member"):
            for h in atlas.search(self.db, text, connection=cube, limit=limit):
                members.append({"cube": h["connection"], "dimension": h["dimension"], "name": h["name"],
                                "label": h["label"], "parents": h["parents"], "depth": h["depth"],
                                "note": self.note(h["name"], h["connection"]),
                                "card": self.compact_card(h["name"], h["connection"]),
                                "caveats": self.caveats({h["connection"]: {h["name"]}})})
        hit = self._matches(text) if kind in ("any", "entity", "kpi", "vocab") else {}
        if kind in ("any", "entity"):
            entities = [self._entity_hit(self._entities[i]) for i in hit["entity"]][:limit]
        if kind in ("any", "kpi"):
            kpis = [self._kpi_hit(self._kpis[i]) for i in hit["kpi"]][:limit]
        if kind in ("any", "vocab"):
            vocab = [self._vocab[i].model_dump(exclude={"source", "approved", "verified"})
                     for i in hit["vocab"]][:limit]
        if kind == "rule":
            rules = [{"id": r.id, "title": r.title, "text": r.text} for r in self._rules.values()
                     if not text.strip() or text.strip().casefold() in f"{r.id} {r.title} {r.text}".casefold()]
        return {"metrics": metrics, "members": members, "concepts": concepts, "entities": entities, "kpis": kpis,
                "vocab": vocab, "rules": rules}

    def close(self) -> None:
        for db in (self.kdb, self.db):
            with contextlib.suppress(Exception):
                db.close()


_BUILD_LOCK = threading.Lock()


REBUILT_NOTE = "knowledge.db rebuilt: semantic files or the atlas changed"


def _atlas_stamp(path: Path) -> str:
    """The raw atlas build time as build() stores it in meta atlas_built_ts ('' without an atlas)."""
    if not path.exists():
        return ""
    db = sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True)
    try:
        return atlas_stamp(db)
    finally:
        db.close()


def ensure_built(notes: list[str] | None = None) -> Path:
    """config.KNOWLEDGE_DB, compiled from config.SEMANTIC_DIR and the atlas first when it is missing, or rebuilt when
    the semantic files or the atlas changed since it was built (the hash or the atlas build time in its meta differs); either is said in notes. A build
    with errors is a GuardError: the answer would rest on knowledge that does not match the files or the cubes (the
    previous knowledge.db is left untouched). Exception: when only the atlas changed and the
    rebuild fails, the existing knowledge.db is kept and a note says so (a refresh never blocks answers)."""
    from ..core import config
    from ..model.validate import GuardError
    path = Path(config.KNOWLEDGE_DB)
    semantic = Path(config.SEMANTIC_DIR)
    with _BUILD_LOCK:                           # two answers at once never write the same .tmp
        missing = not path.exists()
        meta = summary(path) if not missing else {}
        same_semantic = str(meta.get("semantic_hash", "")) == semantic_hash(semantic)
        fresh = same_semantic and str(meta.get("atlas_built_ts", "")) == _atlas_stamp(Path(config.ATLAS_DB))
        if not missing and fresh:
            return path
        report = build(semantic, Path(config.ATLAS_DB), path)
    if not report.ok:
        atlas.need_atlas()          # no atlas at all: one command to run, not the build's member errors
    if not report.ok and not missing and same_semantic:
        if notes is not None:       # only the atlas changed: never block answers on a refresh, keep what works
            notes.append(f"knowledge.db kept from {meta.get('built_at', '?')}: the new atlas breaks it: "
                         f"{'; '.join(report.errors[:3])} - fix knowledge/semantic and run `smartass knowledge build`")
        return path
    if not report.ok:
        what = "is missing and could not be built" if missing else "is out of date and could not be rebuilt"
        raise GuardError(f"knowledge.db {what}: {'; '.join(report.errors[:3])}",
                         "run `smartass knowledge build` in a terminal and fix the errors it lists")
    if notes is not None:
        notes.append(f"knowledge.db was missing and has been built ({report.metrics} metrics, {report.facts} facts, "
                     f"{report.rules} rules, {report.entities} entities, {report.kpis} KPIs, {report.vocab} vocab, "
                     f"{report.notes} notes) in {path.parent}" if missing else REBUILT_NOTE)
    return path


def freshness_notes(db) -> list[str]:
    """Warnings for the answer when the knowledge may be out of date (spec 2026-11-10 §6: warn, never stop)."""
    st = atlas.status(db)
    if st["stale"]:
        age = st["oldest_age_days"]
        return [f"The atlas is {age:.0f} days old (over {atlas.MAX_AGE_DAYS}): new or renamed members may be "
                "missing - run `smartass atlas build`."] if age is not None else [
                "The atlas has never been built - run `smartass atlas build`."]
    return []


def open_without_atlas(forecast_year: int) -> KnowledgeResolver | None:
    """fin_lookup on a PC with no atlas: a KnowledgeResolver on the knowledge.db that is there when it matches the
    semantic files (the one ensure_built would keep), over an empty atlas in memory - no file is created, and no member
    is found (members are the atlas's). None when there is no usable knowledge.db (missing, unreadable, or older than
    the semantic files: it cannot be rebuilt without an atlas)."""
    from ..core import config
    path = Path(config.KNOWLEDGE_DB)
    try:
        if not path.exists() or str(summary(path).get("semantic_hash", "")) != semantic_hash(Path(config.SEMANTIC_DIR)):
            return None
    except (OSError, sqlite3.Error):
        return None
    empty = atlas.empty()
    known = None
    try:
        known = sqlite3.connect(path)
        return KnowledgeResolver(known, empty, forecast_year)
    except Exception:
        for db in (known, empty):
            if db is not None:
                db.close()
        raise


def open_resolver(forecast_year: int, notes: list[str] | None = None) -> KnowledgeResolver:
    """A KnowledgeResolver on knowledge.db (built when missing) and the local atlas; the caller closes it (close())."""
    path = ensure_built(notes)
    db = atlas.connect()
    try:
        if notes is not None:
            notes.extend(freshness_notes(db))
        return KnowledgeResolver(sqlite3.connect(path), db, forecast_year)
    except Exception:
        db.close()
        raise
