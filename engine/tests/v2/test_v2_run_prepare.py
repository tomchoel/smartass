# engine/tests/v2/test_v2_run_prepare.py
import sqlite3
import threading

from test_v2_pipeline import FakeProvider
from v2util import facts, spec_dict

from smartass import pipeline
from smartview_mcp import atlas
from smartass.core import config

ROWS = [("FinMax", "All Entities", "un.SY.2BUS", "2BUS Business Norway", "un.SY", 2, 1),
        ("FinMax", "All Entities", "un.SY.4BUS_B", "4BUS_B Business", "un.SY.2BUS", 3, 1),
        ("FinMax", "All Entities", "un.SY.G910", "G910 Talkmore Business", "un.SY.2BUS", 3, 2),
        ("FinMax", "All Entities", "un.SY.B1", "B1 One", "un.SY.4BUS_B", 4, 1),
        ("FinMax", "All Entities", "un.SY.B2", "B2 Two", "un.SY.4BUS_B", 4, 2),
        ("FinMax", "All Accounts", "ac.C30", "C30 REVENUES", "ac.C48", 3, 1),
        ("FinMax", "All Accounts", "ac.C50S", "C50S", "ac.C70", 3, 1),
        ("FinMax", "All Accounts", "ac.C60", "C60", "ac.C70", 3, 2),
        ("FinMax", "All Accounts", "ac.EBITDA", "EBITDA", "ac.X", 3, 3)]


def make_atlas(path):
    con = atlas.connect(path)
    con.executemany("INSERT INTO member (connection, dimension, name, label, parent, depth, ord) "
                    "VALUES (?,?,?,?,?,?,?)", ROWS)
    con.commit()
    con.close()


def test_run_prepare_opens_and_closes_atlas_in_the_calling_thread(tmp_path, monkeypatch):
    db_path = tmp_path / "atlas.db"
    make_atlas(db_path)
    opened = []

    def connect():
        con = sqlite3.connect(db_path)          # check_same_thread default True
        opened.append((con, threading.get_ident()))
        return con
    monkeypatch.setattr(atlas, "connect", connect)
    monkeypatch.setattr(config, "ATLAS_DB", db_path)

    class Counting(FakeProvider):
        n = 0

        def closed_months(self):
            Counting.n += 1
            return super().closed_months()

    out = {}

    def work():
        out["pre"] = pipeline.run_prepare(spec_dict(), Counting(facts()))
        out["thread"] = threading.get_ident()
    th = threading.Thread(target=work)
    th.start()
    th.join()
    assert "pre" in out, "run_prepare raised in the worker thread"
    assert opened and opened[0][1] == out["thread"] != threading.get_ident()
    assert Counting.n == 1
    try:
        opened[0][0].execute("select 1")
        closed = False
    except sqlite3.ProgrammingError:
        closed = True
    assert closed


def test_prepare_uses_the_atlas_spelling_of_the_entity(tmp_path, monkeypatch):
    db_path = tmp_path / "atlas.db"
    make_atlas(db_path)
    monkeypatch.setattr(atlas, "connect", lambda: sqlite3.connect(db_path))
    d = spec_dict()
    d["queries"][0]["entity"] = "un.sy.2bus"
    prov = FakeProvider(facts())
    pre = pipeline.run_prepare(d, prov)
    assert pre.spec.queries[0].entity == "un.SY.2BUS"
    assert all(e == e.replace("un.sy.", "un.SY.") for r in prov.asked for e in r.entities)
    assert any("un.SY.2BUS" in r.entities for r in prov.asked)
