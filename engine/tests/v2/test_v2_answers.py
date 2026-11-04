# engine/tests/v2/test_v2_answers.py
"""Plan 2B task 3: answer memory (answers.db under a temporary home), fin_critic and fin_lookup(kind='answer')."""
import asyncio
import datetime as dt

import pytest

from smartass import answers, server
from smartass.core import config


@pytest.fixture
def home(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "HOME", tmp_path)
    return tmp_path


def _run(q=1, workbook="C:/a/Analysis.xlsx", question="How is Business opex developing vs last year?", **over):
    run = {"workbook": workbook, "q": q, "question": question, "purpose": "Explain Business opex vs LY",
           "spec": {"version": 2, "question": question}, "metrics": ["Opex", "EBITDA"],
           "entities": ["un.SY.2BUS"], "periods": ["closed-ytd"],
           "key_values": {f"P&L · line {i}": {"Actual": float(i)} for i in range(8)},
           "critic_brief": {"purpose": "Explain Business opex vs LY", "gaps": []}}
    run.update(over)
    return run


def test_store_and_search_round_trip(home):
    answers.store(_run(1))
    answers.store(_run(2, question="Wholesale revenue YTD", purpose="Wholesale revenue", metrics=["Revenue"],
                       entities=["un.SY.2WHS"]))
    assert (home / "answers.db").exists()
    hits = answers.search("business opex", limit=5)
    assert [h["q"] for h in hits] == [1]
    assert hits[0]["question"].startswith("How is Business opex") and hits[0]["workbook"] == "C:/a/Analysis.xlsx"
    assert [h["q"] for h in answers.search("2WHS", 5)] == [2]          # entities are searched
    assert [h["q"] for h in answers.search("revenue", 5)] == [2]
    assert answers.get("C:/a/Analysis.xlsx", 1)["spec"] == {"version": 2,
                                                            "question": "How is Business opex developing vs last year?"}
    assert answers.get("c:\\a\\analysis.xlsx", 1) is not None           # the same file, other spelling
    assert answers.get("C:/a/Analysis.xlsx", 9) is None


def test_a_rerun_replaces_nothing_but_get_and_search_show_the_latest(home):
    answers.store(_run(1, key_values={"a": {"Actual": 1.0}}))
    answers.store(_run(1, key_values={"a": {"Actual": 2.0}}))
    assert answers.get("C:/a/Analysis.xlsx", 1)["key_values"] == {"a": {"Actual": 2.0}}
    assert len(answers.search("opex", 10)) == 1


def test_storing_a_score(home):
    answers.store(_run(1))
    assert answers.set_score("C:/a/Analysis.xlsx", 1, 4.3, "drivers thin on 4BUS_B") is True
    row = answers.get("C:/a/Analysis.xlsx", 1)
    assert row["score"] == 4.3 and row["notes"] == "drivers thin on 4BUS_B"
    assert answers.search("opex", 5)[0]["score"] == 4.3
    assert answers.set_score("C:/a/Other.xlsx", 1, 4.0, "") is False


def test_rows_older_than_24_months_are_purged(home):
    old = (dt.datetime.now() - dt.timedelta(days=800)).isoformat(timespec="seconds")
    answers.store(_run(1), at=old)
    answers.store(_run(2))
    assert [h["q"] for h in answers.search("opex", 10)] == [2]


def test_lookup_kind_answer_gives_the_top_five_key_values(home):
    answers.store(_run(3))
    answers.set_score("C:/a/Analysis.xlsx", 3, 4.0, "")
    out = server.lookup("Business opex", kind="answer")
    a = out["answers"][0]
    assert a["workbook"] == "C:/a/Analysis.xlsx" and a["q"] == 3 and a["score"] == 4.0
    assert a["question"].startswith("How is Business opex") and a["when"]
    assert len(a["key_values"]) == 5 and "P&L · line 0" in a["key_values"]


def _wire(monkeypatch, stored=None, active="Q2 Answer", path="C:/a/Analysis.xlsx"):
    async def user(fn, what):
        return fn(None)
    monkeypatch.setattr(server, "_user", user)
    stored = stored if stored is not None else {q: {"spec": {"v2": {"question": f"q{q}"}}} for q in (1, 2)}
    monkeypatch.setattr(server, "read_stored", lambda h, target="session": (stored, active, path))


def test_fin_critic_without_a_score_returns_the_brief_and_the_rubric(home, monkeypatch):
    _wire(monkeypatch)
    answers.store(_run(2, critic_brief={"purpose": "p2", "gaps": ["material line 'Opex' has no driver"]}))
    out = asyncio.run(server.fin_critic())                     # the active Q2 Answer
    assert out["q"] == 2 and out["brief"]["gaps"] == ["material line 'Opex' has no driver"]
    rubric = (config.KNOWLEDGE_DIR / "agent" / "critic.md").read_text(encoding="utf-8")
    assert out["rubric"] == rubric and "Drivers backed by data" in out["rubric"]


def test_fin_critic_with_a_score_stores_it(home, monkeypatch):
    _wire(monkeypatch)
    answers.store(_run(1))
    out = asyncio.run(server.fin_critic(q=1, score=4.5, notes="clear"))
    assert out == {"stored": True, "q": 1, "score": 4.5}
    assert answers.get("C:/a/Analysis.xlsx", 1)["score"] == 4.5


def test_fin_critic_refuses_a_score_outside_1_to_5_and_an_unstored_answer(home, monkeypatch):
    from smartass.model.validate import GuardError
    _wire(monkeypatch)
    answers.store(_run(1))
    with pytest.raises(GuardError, match="1"):
        asyncio.run(server.fin_critic(q=1, score=6))
    with pytest.raises(GuardError, match="Q2"):
        asyncio.run(server.fin_critic(q=2, score=4))


def test_excel_job_stores_the_run(home, monkeypatch):
    from test_v2_pipeline import FakeProvider
    from test_v2_pipeline_job import wire
    from v2util import facts, resolver, spec_dict

    from smartass import pipeline
    wire(monkeypatch, [])
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    out = pipeline.excel_job(None, pre, target="session")
    row = answers.get(out["workbook"], out["q"])
    assert row["question"] == spec_dict()["question"] and row["purpose"] == "Explain Business opex vs LY"
    assert row["metrics"] == ["EBITDA", "Opex", "Revenue"] and row["entities"] == ["un.SY.2BUS"]
    assert row["periods"] == ["closed"] and row["critic_brief"]["purpose"] == "Explain Business opex vs LY"
    assert row["key_values"] == out["key_values"]


def test_tool_names_include_fin_critic():
    assert "fin_critic" in server.TOOL_NAMES
