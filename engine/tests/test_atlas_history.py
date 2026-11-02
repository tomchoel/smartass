import sqlite3
from datetime import datetime

import pytest

from smartview_mcp import atlas, atlas_history as h

D1 = datetime(2026, 5, 20, 12).timestamp()
D2 = datetime(2026, 6, 5, 12).timestamp()
C, DIM = "FinMax", "All Entities"

BEFORE = [
    ("un.SY.1", "Root", ""),
    ("un.SY.2CON", "Consulting", "un.SY.1"),
    ("un.SY.2CON1", "Con child 1", "un.SY.2CON"),
    ("un.SY.2CON2", "Con child 2", "un.SY.2CON"),
    ("un.SY.3WHS_C", "Whs C", "un.SY.1"),
    ("un.SY.3WHS_W", "Whs W", "un.SY.1"),
    ("un.SY.4REV_W", "4REV_W Revenues Wholesale", "un.SY.3WHS_C"),
    ("un.SY.OLD", "Old name", "un.SY.1"),
    ("un.SY.SHARED", "Shared", "un.SY.3WHS_C"),
    ("un.SY.SHARED", "Shared", "un.SY.ALT"),
]
AFTER = [
    ("un.SY.1", "Root", ""),
    ("un.SY.2CON", "Consulting", "un.SY.1"),
    ("un.SY.2SSM", "Shared services", "un.SY.1"),
    ("un.SY.3WHS_C", "Whs C", "un.SY.1"),
    ("un.SY.3WHS_W", "Whs W", "un.SY.1"),
    ("un.SY.4REV_W", "4REV_W Revenues Wholesale", "un.SY.3WHS_W"),
    ("un.SY.OLD", "New name", "un.SY.1"),
    ("un.SY.SHARED", "Shared", "un.SY.3WHS_C"),
    ("un.SY.SHARED", "Shared", "un.SY.ALT2"),
]


def load(db, rows, ordered=True):
    db.execute("DELETE FROM member")
    rows = rows if ordered else list(reversed(rows))
    db.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)",
                   [(C, DIM, n, lab, p, 0, i) for i, (n, lab, p) in enumerate(rows)])
    db.execute("INSERT OR REPLACE INTO build VALUES (?,?,?,?,?,?)", (C, DIM, 1.0, len(rows), 1.0, None))
    db.commit()


@pytest.fixture
def db(tmp_path, monkeypatch):
    d = sqlite3.connect(tmp_path / "a.db")
    d.executescript(atlas.SCHEMA)
    return d


def at(monkeypatch, ts):
    monkeypatch.setattr(h.time, "time", lambda: ts)


def two_snaps(db, monkeypatch):
    load(db, BEFORE)
    at(monkeypatch, D1)
    s1 = h.take_snapshot(db, C, DIM)
    load(db, AFTER)
    at(monkeypatch, D2)
    s2 = h.take_snapshot(db, C, DIM)
    return s1, s2


def test_reorg_diff(db, monkeypatch):
    s1, s2 = two_snaps(db, monkeypatch)
    assert s1 != s2
    r = h.changes(db, C)
    e = r["dimensions"][DIM]
    assert [x["name"] for x in e["added"]] == ["un.SY.2SSM"]
    assert {x["name"] for x in e["removed"]} == {"un.SY.2CON1", "un.SY.2CON2"}
    moved = {x["name"]: x for x in e["moved"]}
    assert moved["un.SY.4REV_W"]["from"] == ["un.SY.3WHS_C"] and moved["un.SY.4REV_W"]["to"] == ["un.SY.3WHS_W"]
    assert moved["un.SY.SHARED"]["from"] == ["un.SY.3WHS_C", "un.SY.ALT"]
    assert moved["un.SY.SHARED"]["to"] == ["un.SY.3WHS_C", "un.SY.ALT2"]
    assert e["relabelled"] == [{"name": "un.SY.OLD", "old": "Old name", "new": "New name"}]
    assert e["counts"] == {"added": 1, "removed": 2, "moved": 2, "relabelled": 1}
    md = h.changes_markdown(r)
    assert "moved `un.SY.4REV_W` 4REV_W Revenues Wholesale: un.SY.3WHS_C → un.SY.3WHS_W" in md
    assert "added: 1" in md and "1 added" in r["summary"]


def test_duplicate_suppressed_and_order_ignored(db, monkeypatch):
    s1, _ = two_snaps(db, monkeypatch)
    at(monkeypatch, D2 + 100)
    again = h.take_snapshot(db, C, DIM)
    load(db, AFTER, ordered=False)  # pure reordering
    at(monkeypatch, D2 + 200)
    again2 = h.take_snapshot(db, C, DIM)
    assert again == again2
    assert len(h.list_snapshots(db, C, DIM)) == 2


def test_search_and_snapshot_lists_filter_only_when_asked(db, monkeypatch):
    two_snaps(db, monkeypatch)
    assert len(h.list_snapshots(db)) == len(h.list_snapshots(db, C, "")) == 2
    assert h.list_snapshots(db, "FinLite") == []
    assert {r["name"] for r in atlas.search(db, "whs", C, "")} == {"un.SY.3WHS_C", "un.SY.3WHS_W"}
    assert atlas.search(db, "whs", dimension="Account") == []


def test_single_snapshot_message(db, monkeypatch):
    load(db, BEFORE)
    at(monkeypatch, D1)
    h.take_snapshot(db, C, DIM)
    r = h.changes(db, C)
    assert "baseline" in r["summary"] and "next" in r["summary"]
    assert "baseline" in h.changes_markdown(r)


def test_no_members_returns_none(db):
    assert h.take_snapshot(db, C, DIM) is None


def test_since_until_by_date(db, monkeypatch):
    two_snaps(db, monkeypatch)
    load(db, AFTER + [("un.SY.NEW3", "Third", "un.SY.1")])
    at(monkeypatch, datetime(2026, 7, 1, 12).timestamp())
    h.take_snapshot(db, C, DIM)
    # default: latest vs previous -> only NEW3
    assert [x["name"] for x in h.changes(db, C)["dimensions"][DIM]["added"]] == ["un.SY.NEW3"]
    # since 2026-05-25 (base = May 20 snap), until 2026-06-10 (June 5 snap)
    e = h.changes(db, C, DIM, since="2026-05-25", until="2026-06-10")["dimensions"][DIM]
    assert [x["name"] for x in e["added"]] == ["un.SY.2SSM"]
    # since May 25 to latest
    e = h.changes(db, C, DIM, since="2026-05-25")["dimensions"][DIM]
    assert {x["name"] for x in e["added"]} == {"un.SY.2SSM", "un.SY.NEW3"}
    # until on the date of the June snapshot includes it (whole day)
    e = h.changes(db, C, DIM, since="2026-05-25", until="2026-06-05")["dimensions"][DIM]
    assert e["counts"]["added"] == 1


def test_snapshot_all_skips_errors(db, monkeypatch):
    load(db, BEFORE)
    db.execute("INSERT INTO build VALUES ('FinLite','X',1,0,1,'boom')")
    db.commit()
    at(monkeypatch, D1)
    res = h.snapshot_all(db)
    assert len(res) == 1 and res[0]["new"] and res[0]["connection"] == C
    res2 = h.snapshot_all(db)
    assert res2[0]["new"] is False
