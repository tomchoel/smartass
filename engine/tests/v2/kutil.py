"""A tiny knowledge layer for wiring tests: the v2util metrics (Revenue, Opex, EBITDA) as semantic files, a few facts,
and an atlas with the v2util tree. Every figure is invented."""
from __future__ import annotations

import shutil
import sqlite3
from pathlib import Path

from smartass.knowledge import atlas
from smartass.knowledge.build import build

REPO_RULES = Path(__file__).resolve().parents[3] / "knowledge" / "semantic" / "rules"

ATLAS_ROWS = [("FinMax", "All Entities", "un.SY.2BUS", "2BUS Business Norway", "un.SY", 2, 1),
              ("FinMax", "All Entities", "un.SY.4BUS_B", "4BUS_B Business", "un.SY.2BUS", 3, 1),
              ("FinMax", "All Entities", "un.SY.G910", "G910 Talkmore Business", "un.SY.2BUS", 3, 2),
              ("FinMax", "All Entities", "un.SY.B1", "B1 One", "un.SY.4BUS_B", 4, 1),
              ("FinMax", "All Entities", "un.SY.B2", "B2 Two", "un.SY.4BUS_B", 4, 2),
              ("FinMax", "All Entities", "un.SY.2CON", "2CON Consumer Norway", "un.SY", 2, 3),
              ("FinMax", "All Entities", "un.SY.2TEC", "2TEC Technology", "un.SY", 2, 4),
              ("FinMax", "All Accounts", "ac.C30", "C30 REVENUES", "ac.C48", 3, 1),
              ("FinMax", "All Accounts", "ac.C50S", "C50S", "ac.C70", 3, 1),
              ("FinMax", "All Accounts", "ac.C60", "C60", "ac.C70", 3, 2),
              ("FinMax", "All Accounts", "ac.EBITDA", "EBITDA", "ac.X", 3, 3)]

METRICS = {
    "revenue": '''name = "Revenue"
aliases = ["inntekter", "C30"]
description = "Revenue."
sign = "income"
status = "approved"
[finmax]
members = ["ac.C30"]
''',
    "opex": '''name = "Opex"
aliases = ["driftskostnader", "C70"]
description = "Operating expenses."
sign = "cost"
materiality = [5.0, 0.05]
status = "approved"
differs = "FinPlan groups accounts differently."
[finmax]
members = ["ac.C50S", "ac.C60"]
''',
    "ebitda": '''name = "EBITDA"
description = "EBITDA."
sign = "income"
status = "approved"
[finmax]
members = ["ac.EBITDA"]
''',
}

FACTS = {
    "g910_moved": '''text = "G910 moved in the reorg."
valid_from = "2026-06"
severity = "info"
source = "test"
status = "approved"
[scope_members]
FinMax = ["un.SY.G910"]
''',
    "opex_note": '''text = "Opex note."
scope_metrics = ["opex"]
severity = "warn"
source = "test"
status = "approved"
''',
    "bus_old": '''text = "Business note until mid 2025."
valid_to = "2025-06"
severity = "block"
source = "test"
status = "approved"
[scope_members]
FinMax = ["un.SY.2BUS"]
''',
    "abolished_consumer": '''text = "Consumer Norway was abolished on 1 Jun 2026."
valid_from = "2026-06"
severity = "warn"
source = "test"
status = "approved"
[scope_members]
FinMax = ["un.SY.2CON"]
''',
    "tec_lines": '''text = "Technology C60 lines note."
severity = "info"
source = "test"
status = "approved"
[scope_members]
FinMax = ["un.SY.2TEC"]
''',
}


def make_atlas(path) -> None:
    con = atlas.connect(path)
    con.executemany("INSERT INTO member (connection, dimension, name, label, parent, depth, ord) "
                    "VALUES (?,?,?,?,?,?,?)", ATLAS_ROWS)
    con.commit()
    con.close()


def make_semantic(root, metrics=None, facts=None):
    for kind, files in (("metrics", metrics or METRICS), ("facts", FACTS if facts is None else facts)):
        d = root / kind
        d.mkdir(parents=True, exist_ok=True)
        for name, text in files.items():
            (d / f"{name}.toml").write_text(text, encoding="utf-8")
    rules = root / "rules"                     # the engine reads these two rules: copy the repo's
    rules.mkdir(parents=True, exist_ok=True)
    for name in ("scenarios", "materiality"):
        shutil.copy(REPO_RULES / f"{name}.toml", rules / f"{name}.toml")
    return root


def knowledge(tmp_path, monkeypatch=None, build_db: bool = True):
    """atlas.db + semantic/ (+ knowledge.db when build_db) in tmp_path. With monkeypatch: config points at them and
    atlas.connect() opens the tiny atlas."""
    from smartass.core import config
    a, sem, k = tmp_path / "atlas.db", tmp_path / "semantic", tmp_path / "home" / "knowledge.db"
    make_atlas(a)
    make_semantic(sem)
    if build_db:
        assert build(sem, a, k).ok
    if monkeypatch is not None:
        monkeypatch.setattr(config, "ATLAS_DB", a)
        monkeypatch.setattr(config, "SEMANTIC_DIR", sem)
        monkeypatch.setattr(config, "KNOWLEDGE_DB", k)
        monkeypatch.setattr(atlas, "connect", lambda path=None: sqlite3.connect(path or a))
    return a, sem, k


def kresolver(tmp_path, yy: int = 26):
    from smartass.knowledge.resolver import KnowledgeResolver
    a, _, k = knowledge(tmp_path)
    return KnowledgeResolver(sqlite3.connect(k), sqlite3.connect(a), forecast_year=yy)
