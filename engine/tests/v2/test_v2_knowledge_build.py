import sqlite3

import pytest

from smartass.knowledge import atlas
from smartass.knowledge.build import build, search_metrics

OPEX = """
name = "Opex"
aliases = ["opex", "driftskostnader", "C70"]
description = "Total operating expenses."
sign = "cost"
status = "approved"
[finmax]
members = ["ac.C70"]
[finplan]
members = ["OPEX"]
"""
REVENUE = """
name = "Revenue"
aliases = ["revenue", "inntekter", "C30"]
description = "Revenue."
sign = "income"
status = "approved"
[finmax]
members = ["ac.C30"]
"""
FACT = """
text = "Consumer Norway was abolished on 1 Jun 2026."
valid_from = "2026-06"
severity = "warn"
source = "test"
status = "approved"
[scope_members]
FinMax = ["un.SY.2CON", "un.SY.GONE"]
"""


@pytest.fixture
def atlas_db(tmp_path):
    p = tmp_path / "atlas.db"
    con = atlas.connect(p)
    rows = [("FinMax", "All Entities", "un.SY.2CON", "2CON Consumer Norway", "un.SY", 2, 1),
            ("FinMax", "All Accounts", "ac.C70", "C70 Opex", "ac.EBITDA_BOI", 3, 1),
            ("FinMax", "All Accounts", "ac.C30", "C30 REVENUES", "ac.C48", 3, 1),
            ("FinPlan", "Account", "OPEX", "OPEX", "EBITDA before Other Items", 3, 1)]
    con.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)", rows)
    con.execute("INSERT INTO build VALUES ('FinMax','All Accounts',1790000000.0,3,1.0,NULL)")
    con.commit()
    con.close()
    return p


def _semantic(tmp_path, files: dict[str, str]):
    d = tmp_path / "semantic"
    for rel, text in files.items():
        (d / rel).parent.mkdir(parents=True, exist_ok=True)
        (d / rel).write_text(text, encoding="utf-8")
    from kutil import copy_required_rules
    copy_required_rules(d)
    return d


def _good(tmp_path):
    return _semantic(tmp_path, {"metrics/opex.toml": OPEX, "metrics/revenue.toml": REVENUE,
                                "facts/abolished_consumer.toml": FACT})


def test_good_build(tmp_path, atlas_db):
    out = tmp_path / "knowledge.db"
    r = build(_good(tmp_path), atlas_db, out)
    assert r.ok and r.errors == [] and (r.metrics, r.facts, r.drafts) == (2, 1, 0)
    assert any("un.SY.GONE" in w for w in r.warnings)               # historical scope member: warning only
    assert not (tmp_path / "knowledge.db.tmp").exists()
    db = sqlite3.connect(out)
    assert db.execute("SELECT member FROM metric_member WHERE metric_id='opex' AND cube='FinPlan'").fetchall() == \
        [("OPEX",)]
    assert ("FinMax", "un.SY.2CON") in db.execute("SELECT cube, member FROM fact_scope").fetchall()
    meta = dict(db.execute("SELECT key, value FROM meta"))
    assert {"built_at", "semantic_hash", "atlas_built_at"} <= set(meta)
    assert "\n" not in r.to_text().splitlines()[0]


def test_missing_member_is_an_error_and_the_old_db_is_kept(tmp_path, atlas_db):
    out = tmp_path / "knowledge.db"
    assert build(_good(tmp_path), atlas_db, out).ok
    before = out.read_bytes()
    bad = _semantic(tmp_path / "b", {"metrics/opex.toml": OPEX.replace('"OPEX"]', '"OPEXX"]')})
    r = build(bad, atlas_db, out)
    assert not r.ok and any("OPEXX" in e and "FinPlan" in e for e in r.errors)
    assert out.read_bytes() == before


def test_alias_clash_is_an_error(tmp_path, atlas_db):
    d = _semantic(tmp_path, {"metrics/opex.toml": OPEX,
                             "metrics/revenue.toml": REVENUE.replace('"inntekter"', '"Driftskostnader"')})
    r = build(d, atlas_db, tmp_path / "k.db")
    assert not r.ok and any("driftskostnader" in e.lower() and "opex" in e and "revenue" in e for e in r.errors)
    assert not (tmp_path / "k.db").exists()


def test_invalid_semantic_file_is_an_error(tmp_path, atlas_db):
    d = _semantic(tmp_path, {"metrics/opex.toml": "name = "})
    r = build(d, atlas_db, tmp_path / "k.db")
    assert not r.ok and "opex.toml" in r.errors[0]


def test_drafts_are_counted_as_a_warning(tmp_path, atlas_db):
    d = _semantic(tmp_path, {"metrics/opex.toml": OPEX.replace('status = "approved"', 'status = "draft"')})
    r = build(d, atlas_db, tmp_path / "k.db")
    assert r.ok and r.drafts == 1 and any("1 draft" in w for w in r.warnings)


def test_fts_finds_a_norwegian_alias(tmp_path, atlas_db):
    out = tmp_path / "knowledge.db"
    build(_good(tmp_path), atlas_db, out)
    db = sqlite3.connect(out)
    assert search_metrics(db, "driftskostnader")[0] == "opex"
    assert search_metrics(db, "Driftskost")[0] == "opex"            # prefix, any case
    assert search_metrics(db, "revenue")[0] == "revenue"
    assert search_metrics(db, "nothing like it") == []
    assert db.execute("SELECT count(*) FROM member_fts WHERE member_fts MATCH 'consumer'").fetchone()[0] == 1


def test_cli_build_prints_the_report(tmp_path, atlas_db, capsys):
    from smartass.ops import cli
    rc = cli.main(["knowledge", "build", "--semantic", str(_good(tmp_path)), "--atlas", str(atlas_db),
                   "--out", str(tmp_path / "k.db")])
    out = capsys.readouterr().out
    assert rc == 0 and "2 metrics" in out and "1 facts" in out and (tmp_path / "k.db").exists()
