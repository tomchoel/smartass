"""Compile the atlas and the semantic files into knowledge.db (sub-spec 2 §3). The build checks first; an error
stops it and leaves the previous knowledge.db untouched. A good build writes `<out>.tmp` and swaps it in with
os.replace, so readers never see half a database."""
from __future__ import annotations

import hashlib
import json
import os
import re
import sqlite3
import time
from dataclasses import dataclass, field
from pathlib import Path

from .expand import LazyTrees, expand
from .semantic import SemanticError, concept_terms, load_concepts, load_semantic

SCHEMA = """
CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE metric (id TEXT PRIMARY KEY, name TEXT, json TEXT);
CREATE TABLE metric_member (metric_id TEXT, cube TEXT, member TEXT);
CREATE TABLE alias (metric_id TEXT, text TEXT);
CREATE INDEX ix_alias_text ON alias(text COLLATE NOCASE);
CREATE VIRTUAL TABLE alias_fts USING fts5(text, metric_id UNINDEXED);
CREATE TABLE fact (id TEXT PRIMARY KEY, json TEXT);
CREATE TABLE fact_scope (fact_id TEXT, cube TEXT, member TEXT, metric_id TEXT, via TEXT);
CREATE INDEX ix_fact_scope_member ON fact_scope(member COLLATE NOCASE);
CREATE INDEX ix_fact_scope_metric ON fact_scope(metric_id);
CREATE VIRTUAL TABLE member_fts USING fts5(name, label, cube UNINDEXED, dimension UNINDEXED);
CREATE TABLE concept (id TEXT PRIMARY KEY, title TEXT, json TEXT);
CREATE TABLE concept_alias (concept_id TEXT, text TEXT);
CREATE INDEX ix_concept_alias_text ON concept_alias(text COLLATE NOCASE);
CREATE VIRTUAL TABLE concept_fts USING fts5(text, concept_id UNINDEXED);
"""
SCHEMA_VERSION = "3a"      # part of the semantic hash: a new schema rebuilds knowledge.db like changed files
# Member labels indexed for search: the account and entity trees (the dimensions questions name).
SEARCH_DIMENSIONS = ("All Accounts", "All Entities", "Account", "Entity")


@dataclass
class BuildReport:
    out: Path
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    metrics: int = 0
    facts: int = 0
    concepts: int = 0
    drafts: int = 0
    built: bool = False

    @property
    def ok(self) -> bool:
        return not self.errors

    def to_text(self) -> str:
        head = (f"knowledge.db {'built' if self.built else 'NOT built'}: {self.metrics} metrics, {self.facts} facts, "
                f"{self.concepts} concepts, {self.drafts} drafts, {len(self.errors)} errors, {len(self.warnings)} warnings ({self.out})")
        return "\n".join([head, *(f"ERROR {e}" for e in self.errors), *(f"WARN  {w}" for w in self.warnings)])


def _atlas(path: Path) -> sqlite3.Connection:
    if not Path(path).exists():
        raise FileNotFoundError(f"no atlas at {path} (run `smartass atlas build`)")
    return sqlite3.connect(Path(path).resolve().as_uri() + "?mode=ro", uri=True)


def _missing(atlas: sqlite3.Connection, cube: str, members) -> list[str]:
    """The members the atlas does not have for the cube (exact name first, then any case, like Essbase)."""
    out = []
    for m in members:
        if atlas.execute("SELECT 1 FROM member WHERE connection=? AND name=? LIMIT 1", (cube, m)).fetchone():
            continue
        if atlas.execute("SELECT 1 FROM member WHERE connection=? AND lower(name)=lower(?) LIMIT 1",
                         (cube, m)).fetchone():
            continue
        out.append(m)
    return out


def semantic_hash(semantic_dir: Path) -> str:
    h = hashlib.sha256(f"schema {SCHEMA_VERSION}\n".encode())
    for p in sorted(Path(semantic_dir).glob("*/*.toml")):
        h.update(p.relative_to(semantic_dir).as_posix().encode())
        h.update(p.read_bytes())
    return h.hexdigest()


def check(metrics, facts, atlas: sqlite3.Connection, report: BuildReport, concepts=()) -> None:
    ids = {m.id for m in metrics}
    for m in metrics:
        for cube, members in m.members().items():
            for name in _missing(atlas, cube, members):
                report.errors.append(f"metric {m.id}: {cube} member {name!r} is not in the atlas")
    owners: dict[str, set[str]] = {}
    for m in metrics:
        for term in {m.name.casefold(), *(a.casefold() for a in m.aliases)}:
            owners.setdefault(term, set()).add(m.id)
    for term, who in sorted(owners.items()):
        if len(who) > 1:
            report.errors.append(f"alias {term!r} is used by {len(who)} metrics: {', '.join(sorted(who))}")
    for f in facts:
        for mid in f.scope_metrics:
            if mid not in ids:
                report.errors.append(f"fact {f.id}: scope metric {mid!r} is not a metric")
        for cube, members in f.scope_members.items():
            gone = _missing(atlas, cube, members)
            if gone:
                report.warnings.append(f"fact {f.id}: {len(gone)} {cube} scope member(s) not in the atlas "
                                       f"(historical?): {', '.join(gone[:10])}{' ...' if len(gone) > 10 else ''}")
    check_concepts(metrics, facts, concepts, atlas, report)
    report.drafts = sum(x.status == "draft" for x in [*metrics, *facts, *concepts])
    if report.drafts:
        report.warnings.append(f"{report.drafts} draft metric(s)/fact(s) await approval (smartass knowledge approve)")


CONCEPT_CUBE = "FinMax"      # concept members are FinMax names (FinLite shares them)


def check_concepts(metrics, facts, concepts, atlas: sqlite3.Connection, report: BuildReport) -> None:
    """Concept ids are unique across all semantic files; a concept's aliases are its own (never another concept's,
    never a metric's): errors. Concept members missing from the atlas are warnings (the answer guard refuses them)."""
    taken = {x.id for x in [*metrics, *facts]}
    for c in concepts:
        if c.id in taken:
            report.errors.append(f"concept {c.id}: the id is also a metric or fact id")
    metric_terms: dict[str, str] = {}
    for m in metrics:
        for t in (m.id.replace("_", " "), m.name, *m.aliases):
            metric_terms.setdefault(t.casefold(), m.id)
    owner: dict[str, str] = {}
    for c in concepts:
        for t in concept_terms(c):
            k = t.casefold()
            if k in metric_terms:
                report.errors.append(f"concept {c.id}: alias {t!r} is also an alias of metric {metric_terms[k]}")
            elif owner.setdefault(k, c.id) != c.id:
                report.errors.append(f"alias {t!r} is used by concepts {owner[k]} and {c.id}")
        gone = [x for xs in c.members().values() for x in _missing(atlas, CONCEPT_CUBE, xs)]
        if gone:
            report.warnings.append(f"concept {c.id}: {len(gone)} {CONCEPT_CUBE} member(s) not in the atlas: "
                                   f"{', '.join(gone[:10])}{' ...' if len(gone) > 10 else ''}")


MAX_SCOPE = 2000        # a fact reaching more members of one cube than this gets a build warning


def _write(db: sqlite3.Connection, metrics, facts, atlas: sqlite3.Connection, meta: dict, concepts=(),
           entities=(), warnings: list[str] | None = None) -> None:
    db.executescript(SCHEMA)
    for m in metrics:
        db.execute("INSERT INTO metric VALUES (?,?,?)", (m.id, m.name, m.model_dump_json()))
        db.executemany("INSERT INTO metric_member VALUES (?,?,?)",
                       [(m.id, cube, x) for cube, ms in m.members().items() for x in ms])
        terms = list(dict.fromkeys([m.name, m.id.replace("_", " "), *m.aliases]))
        db.executemany("INSERT INTO alias VALUES (?,?)", [(m.id, t) for t in terms])
        db.executemany("INSERT INTO alias_fts (text, metric_id) VALUES (?,?)", [(t, m.id) for t in terms])
    stops: dict[str, set[str]] = {}
    for m in metrics:
        for cube, ms in m.members().items():
            stops.setdefault(cube, set()).update(ms)
    for e in entities:
        for cube, member in e.members().items():
            stops.setdefault(cube, set()).add(member)
    trees = LazyTrees(atlas)                    # each cube's member table is read once per build
    for f in facts:
        db.execute("INSERT INTO fact VALUES (?,?)", (f.id, f.model_dump_json()))
        scope = expand(f, trees, stops)
        db.executemany("INSERT INTO fact_scope VALUES (?,?,?,NULL,?)",
                       [(f.id, cube, member, via) for (cube, member), via in scope.items()])
        if warnings is not None:
            per_cube: dict[str, int] = {}
            for cube, _member in scope:
                per_cube[cube] = per_cube.get(cube, 0) + 1
            for cube, n in sorted(per_cube.items()):
                if n > MAX_SCOPE:
                    warnings.append(f"fact {f.id}: inherit {f.inherit!r} expands to {n} {cube} members "
                                    f"(over {MAX_SCOPE}); check the scope")
        db.executemany("INSERT INTO fact_scope VALUES (?,NULL,NULL,?,NULL)", [(f.id, mid) for mid in f.scope_metrics])
    for c in concepts:
        db.execute("INSERT INTO concept VALUES (?,?,?)", (c.id, c.title, c.model_dump_json()))
        terms = concept_terms(c)
        db.executemany("INSERT INTO concept_alias VALUES (?,?)", [(c.id, t) for t in terms])
        db.executemany("INSERT INTO concept_fts (text, concept_id) VALUES (?,?)", [(t, c.id) for t in terms])
    marks = ",".join("?" * len(SEARCH_DIMENSIONS))
    rows = atlas.execute(f"SELECT DISTINCT name, label, connection, dimension FROM member WHERE dimension IN ({marks})",
                         SEARCH_DIMENSIONS)
    db.executemany("INSERT INTO member_fts (name, label, cube, dimension) VALUES (?,?,?,?)", rows)
    db.executemany("INSERT INTO meta VALUES (?,?)", [(k, str(v)) for k, v in meta.items()])


CARRIED_META = ("refresh_month",)      # kept from the previous knowledge.db by every rebuild (set by a refresh)


def build(semantic_dir: Path, atlas_db: Path, out_db: Path, extra_meta: dict | None = None) -> BuildReport:
    """extra_meta joins the meta table (a refresh passes its month); CARRIED_META keys of the old db are kept."""
    out_db = Path(out_db)
    report = BuildReport(out=out_db)
    try:
        metrics, facts = load_semantic(Path(semantic_dir))
        concepts = load_concepts(Path(semantic_dir))
    except SemanticError as e:
        report.errors.append(str(e))
        return report
    report.metrics, report.facts, report.concepts = len(metrics), len(facts), len(concepts)
    try:
        atlas = _atlas(atlas_db)
    except FileNotFoundError as e:
        report.errors.append(str(e))
        return report
    try:
        check(metrics, facts, atlas, report, concepts)
        if report.errors:
            return report
        atlas_built = atlas.execute("SELECT max(built_at) FROM build").fetchone()[0]
        meta = {"built_at": time.strftime("%Y-%m-%d %H:%M:%S"), "semantic_hash": semantic_hash(Path(semantic_dir)),
                "atlas_built_ts": repr(float(atlas_built)) if atlas_built else "",
                "atlas_built_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(atlas_built)) if atlas_built
                else "", "metrics": len(metrics), "facts": len(facts), "concepts": len(concepts),
                "drafts": report.drafts}
        old = summary(out_db)
        meta.update({k: old[k] for k in CARRIED_META if k in old})
        meta.update(extra_meta or {})
        out_db.parent.mkdir(parents=True, exist_ok=True)
        tmp = out_db.with_name(out_db.name + ".tmp")
        tmp.unlink(missing_ok=True)
        db = sqlite3.connect(tmp)
        try:
            with db:
                _write(db, metrics, facts, atlas, meta, concepts, warnings=report.warnings)
        finally:
            db.close()
        os.replace(tmp, out_db)
        report.built = True
        return report
    finally:
        atlas.close()


def _fts_query(text: str) -> str | None:
    words = re.findall(r"\w+", text)
    return " AND ".join(f'"{w}"*' for w in words) or None


def search_metrics(db: sqlite3.Connection, text: str, limit: int = 10) -> list[str]:
    """Metric ids for a phrase: an exact alias or name (any case) first, then FTS prefix matches by rank."""
    out = [r[0] for r in db.execute("SELECT DISTINCT metric_id FROM alias WHERE text = ? COLLATE NOCASE",
                                    (text.strip(),))]
    q = _fts_query(text)
    if q:
        for (mid,) in db.execute("SELECT metric_id FROM alias_fts WHERE alias_fts MATCH ? ORDER BY rank", (q,)):
            if mid not in out:
                out.append(mid)
    return out[:limit]


def search_members(db: sqlite3.Connection, text: str, cube: str | None = None, limit: int = 25) -> list[dict]:
    """Members by name or label (FTS prefix match over the account and entity trees)."""
    q = _fts_query(text)
    if not q:
        return []
    sql = ("SELECT name, label, cube, dimension FROM member_fts WHERE member_fts MATCH ? "
           "AND coalesce(?, cube) = cube ORDER BY rank LIMIT ?")
    return [{"name": n, "label": lab, "cube": c, "dimension": d}
            for n, lab, c, d in db.execute(sql, (q, cube, limit))]


def summary(path: Path) -> dict:
    """meta of an existing knowledge.db ({} if there is none)."""
    if not Path(path).exists():
        return {}
    db = sqlite3.connect(Path(path).resolve().as_uri() + "?mode=ro", uri=True)
    try:
        return {k: (json.loads(v) if v.isdigit() else v) for k, v in db.execute("SELECT key, value FROM meta")}
    finally:
        db.close()


def search_concepts(db: sqlite3.Connection, text: str, limit: int = 10) -> list[str]:
    """Concept ids for a phrase: an exact alias (any case) first, then FTS prefix matches by rank."""
    out = [r[0] for r in db.execute("SELECT DISTINCT concept_id FROM concept_alias WHERE text = ? COLLATE NOCASE",
                                    (text.strip(),))]
    q = _fts_query(text)
    if q:
        for (cid,) in db.execute("SELECT concept_id FROM concept_fts WHERE concept_fts MATCH ? ORDER BY rank", (q,)):
            if cid not in out:
                out.append(cid)
    return out[:limit]
