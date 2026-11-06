"""Plan 2A task 5: the knowledge layer wired into prepare/reply, the Answer sheet, run_prepare, the test set runner,
fin_lookup, fin_status and the doctor."""
import importlib
import sqlite3

import pytest
from kutil import knowledge, kresolver
from test_v2_pipeline import FakeProvider
from v2util import facts, spec_dict

from smartass import pipeline
from smartass.core import config
from smartass.model.analyze import compute
from smartass.model.spec import AnalysisSpec
from smartass.pipeline import prepare, reply
from smartass.render.layout import layout


def _ids(cavs):
    return [c["id"] for c in cavs]


# ---- caveats in prepare and reply

def test_prepare_collects_caveats_for_touched_members_metrics_and_months(tmp_path):
    r = kresolver(tmp_path)
    pre = prepare(spec_dict(), r, FakeProvider(facts()))
    ids = _ids(pre.caveats)
    assert "g910_moved" in ids            # G910 sits in the 2BUS subtree the drivers block fetches; Sep 26 >= 2026-06
    assert "opex_note" in ids             # the query asks for Opex (metric scope)
    assert "bus_old" not in ids           # valid to Jun 25; the answer touches Sep 26 and Sep 25 only
    assert "abolished_consumer" not in ids and "tec_lines" not in ids
    assert ids == ["opex_note", "g910_moved"]          # warn before info
    out = reply(pre, q=1, workbook="W", sheets=[], mismatches=[], lint=[])
    assert out["caveats"] == pre.caveats
    assert out["caveats"][0] == {"id": "opex_note", "text": "Opex note.", "severity": "warn"}


def test_a_resolver_without_caveats_gives_none():
    from v2util import resolver
    pre = prepare(spec_dict(), resolver(), FakeProvider(facts()))
    assert pre.caveats == []
    assert reply(pre, 1, "W", [], [], [])["caveats"] == []


def test_lines_blocks_contribute_their_cell_members(tmp_path):
    from smartass.model.lines import LineSet
    r = kresolver(tmp_path)
    cell = {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Yr 26"}
    d = spec_dict(blocks=[{"type": "lines", "cell": cell, "group_by": "supplier", "top": 2}])
    d["queries"] = []
    ls = LineSet("b0", [{"AMOUNT": -10.0, "AP_CUSTOMER_NAME": "A"}], ["AP_CUSTOMER_NAME", "AMOUNT"], -10.0,
                 reconciliation={"gap": 0.0}, warnings=[], cell=cell)
    pre = prepare(d, r, FakeProvider(facts()), lines={"b0": ls})
    assert _ids(pre.caveats) == ["tec_lines"]        # FinLite cell, fact scoped to the FinMax member it shares


# ---- the Answer sheet

def test_answer_sheet_gets_a_caveats_section():
    from v2util import CLOSED, resolver
    s = AnalysisSpec.model_validate(spec_dict())
    res = compute(s, resolver(), CLOSED, facts())
    long = "A long caveat text " * 8
    p = layout(3, s, res, stock=set(), subtitle="x",
               caveats=[{"id": "a", "text": long, "severity": "warn"}, {"id": "b", "text": "Short.", "severity": "info"}])
    labels = {r["row"]: p.cells.get((r["row"], 2)) for r in p.rows}
    sec = [row for row, v in labels.items() if v == "Caveats"]
    assert len(sec) == 1
    notes = [labels[r["row"]] for r in p.rows if r["kind"] == "note" and r["row"] > sec[0]]
    assert notes[0].startswith("Warn: A long caveat text") and 78 < len(notes[0]) <= 140   # the table width
    assert any(n == "Info: Short." for n in notes)
    src = next(i for i, n in enumerate(notes) if n.startswith("Source:"))      # the source note (wrapped) comes last
    assert src > notes.index("Info: Short.")


def test_no_caveats_no_section():
    from v2util import CLOSED, resolver
    s = AnalysisSpec.model_validate(spec_dict())
    res = compute(s, resolver(), CLOSED, facts())
    p = layout(3, s, res, stock=set(), subtitle="x")
    assert "Caveats" not in [p.cells.get((r["row"], 2)) for r in p.rows]


# ---- run_prepare and the test set runner

def test_run_prepare_uses_the_knowledge_resolver(tmp_path, monkeypatch):
    knowledge(tmp_path, monkeypatch)
    pre = pipeline.run_prepare(spec_dict(), FakeProvider(facts()))
    assert "opex_note" in _ids(pre.caveats)
    assert not any("knowledge.db" in w for w in pre.warnings)


def test_missing_knowledge_db_is_built_on_the_fly_with_a_warning(tmp_path, monkeypatch):
    _, _, k = knowledge(tmp_path, monkeypatch, build_db=False)
    assert not k.exists()
    pre = pipeline.run_prepare(spec_dict(), FakeProvider(facts()))
    assert k.exists()
    assert any("knowledge.db" in w and "built" in w for w in pre.warnings)
    assert "opex_note" in _ids(pre.caveats)


def test_a_failing_on_the_fly_build_is_a_clear_refusal(tmp_path, monkeypatch):
    from smartass.model.validate import GuardError
    _, sem, _ = knowledge(tmp_path, monkeypatch, build_db=False)
    (sem / "metrics" / "bad.toml").write_text('name = "Bad"\ndescription = "x"\n[finmax]\nmembers = ["ac.NOPE"]\n',
                                              encoding="utf-8")
    with pytest.raises(GuardError, match="knowledge.db"):
        pipeline.run_prepare(spec_dict(), FakeProvider(facts()))


def test_testset_runner_uses_the_knowledge_resolver(tmp_path, monkeypatch):
    from smartass.knowledge.resolver import KnowledgeResolver
    from smartass.testset import runner
    knowledge(tmp_path, monkeypatch)
    r = runner._resolver_for({"FinMax": (2026, 9), "FinPlan": (2026, 9)})
    try:
        assert isinstance(r, KnowledgeResolver) and r.forecast_year == 26
    finally:
        runner._close(r)
    with pytest.raises(sqlite3.ProgrammingError):
        r.kdb.execute("select 1")


def test_interim_knowledge_is_gone():
    for mod in ("smartass.knowledge.interim", "smartass.knowledge.atlas_resolver"):
        with pytest.raises(ImportError):
            importlib.import_module(mod)


# ---- fin_lookup, fin_status, doctor

def test_lookup_searches_aliases_and_members_with_caveats(tmp_path, monkeypatch):
    from smartass import server
    knowledge(tmp_path, monkeypatch)
    out = server.lookup("driftskostnader", kind="metric")
    m = out["metrics"][0]
    assert m["id"] == "opex" and "driftskostnader" in m["aliases"]
    assert m["description"] == "Operating expenses." and m["differs"].startswith("FinPlan groups")
    assert out["members"] == []
    mem = server.lookup("2CON", kind="member")["members"]
    hit = next(x for x in mem if x["name"] == "un.SY.2CON")
    assert [c["id"] for c in hit["caveats"]] == ["abolished_consumer"]


def test_fin_status_reports_the_knowledge_build(tmp_path, monkeypatch):
    import asyncio

    from smartass import server
    from smartass.ops import doctor
    knowledge(tmp_path, monkeypatch)
    monkeypatch.setattr(server.atlas, "status", lambda db: {"stale": False, "oldest_age_days": 3.0, "built": [1]})
    monkeypatch.setattr(doctor, "run", lambda: [doctor.Check("excel", "OK", "x")])
    out = asyncio.run(server.fin_status())
    kn = out["knowledge"]
    assert (kn["metrics"], kn["facts"], kn["drafts"]) == (3, 5, 0) and kn["built_at"]
    assert out["metrics"] == 3


def test_doctor_knowledge_check(tmp_path, monkeypatch):
    import os

    from smartass.ops import doctor
    assert doctor._knowledge in doctor.CHECKS and not hasattr(doctor, "_catalog")
    a, sem, k = knowledge(tmp_path, monkeypatch)
    old = os.path.getmtime(a) - 10
    os.utime(a, (old, old))
    c = doctor._knowledge()
    assert (c.name, c.status) == ("knowledge", "OK"), c
    assert "3 metrics" in c.detail and "5 facts" in c.detail
    os.utime(a, None)                                   # the atlas is newer than knowledge.db
    os.utime(k, (old - 10, old - 10))
    c = doctor._knowledge()
    assert c.status == "WARN" and "older than the atlas" in c.detail
    (sem / "facts" / "d.toml").write_text('text = "d"\nsource = "s"\nstatus = "draft"\n[scope_members]\n'
                                          'FinMax = ["un.SY.B1"]\n', encoding="utf-8")
    c = doctor._knowledge()
    assert c.status == "WARN" and "draft" in c.detail
    (sem / "metrics" / "bad.toml").write_text('name = "Bad"\ndescription = "x"\n[finmax]\nmembers = ["ac.NOPE"]\n',
                                              encoding="utf-8")
    c = doctor._knowledge()
    assert c.status == "FAIL" and "ac.NOPE" in c.detail
    k.unlink()
    (sem / "metrics" / "bad.toml").unlink()
    c = doctor._knowledge()
    assert c.status == "WARN" and "not built" in c.detail
    monkeypatch.setattr(config, "ATLAS_DB", tmp_path / "none.db")
    assert doctor._knowledge().status == "WARN"
