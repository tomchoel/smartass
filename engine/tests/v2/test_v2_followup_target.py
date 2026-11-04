# engine/tests/v2/test_v2_followup_target.py
"""Follow-ups, fin_refresh and fin_critic default to the workbook of the latest answer (answers.db), not the
session workbook; an explicit target wins; without a stored answer the session workbook is used."""
import asyncio

import pytest

from smartass import answers, pipeline, server
from smartass.core import config, session


@pytest.fixture
def home(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "HOME", tmp_path)
    return tmp_path


def _store(workbook, q=1):
    answers.store({"workbook": workbook, "q": q, "question": f"q{q}", "spec": {"v2": {"question": f"q{q}"}}})


def _wire(monkeypatch):
    seen = []

    async def user(fn, what):
        return fn(None)

    def read_stored(h, target="session"):
        seen.append(target)
        return {1: {"spec": {"v2": {"question": "q1"}}, "ids": {}, "owned": None}}, "Q1 Answer", "Book1"
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "_provider", lambda: "provider")
    monkeypatch.setattr(server, "read_stored", read_stored)
    monkeypatch.setattr(server, "run_prepare", lambda raw, provider, lines=None: raw)
    monkeypatch.setattr(server, "redraw_job", lambda h, pre, q, owned, target="session": {"q": q, "workbook": target})
    return seen


def test_followup_defaults_to_the_workbook_of_the_latest_answer(home, monkeypatch):
    monkeypatch.setattr(session, "workbook", lambda: "C:/an/Session.xlsx")
    _store("C:/an/Session.xlsx")
    _store("Book1")
    seen = _wire(monkeypatch)
    out = asyncio.run(server.fin_analyze(patch={"plan": {"purpose": "x"}}))
    assert seen == ["Book1"] and out["workbook"] == "Book1"


def test_refresh_defaults_to_the_workbook_of_the_latest_answer(home, monkeypatch):
    _store("Book1")
    seen = _wire(monkeypatch)
    asyncio.run(server.fin_refresh())
    assert seen == ["Book1"]


def test_an_explicit_target_wins(home, monkeypatch):
    _store("Book1")
    seen = _wire(monkeypatch)
    asyncio.run(server.fin_analyze(patch={"plan": {"purpose": "x"}}, target="active"))
    assert seen == ["active"]


def test_without_a_stored_answer_the_session_workbook_is_used(home, monkeypatch):
    seen = _wire(monkeypatch)
    asyncio.run(server.fin_analyze(patch={"plan": {"purpose": "x"}}))
    assert seen == ["session"]


def test_latest_answer_in_the_session_workbook_keeps_target_session(home, monkeypatch):
    monkeypatch.setattr(session, "workbook", lambda: "C:/an/Session.xlsx")
    _store("c:/AN/session.xlsx")
    seen = _wire(monkeypatch)
    asyncio.run(server.fin_analyze(patch={"plan": {"purpose": "x"}}))
    assert seen == ["session"]                          # the session file stays 'ours' (saved after each answer)


def test_followup_on_book1_resolves_the_open_unsaved_workbook(home, monkeypatch):
    """read_stored with the defaulted 'Book1' finds the open unsaved book by name (resolve_target, fix A)."""
    from test_open_excel import Books, FakeApp, NamedWb, _by_name

    from smartass.xl import jobs, login, sheets
    monkeypatch.setattr(jobs, "hold_turn", lambda *a, **k: True)
    monkeypatch.setattr(login, "wait_login_cleared", lambda *a, **k: None)
    monkeypatch.setattr(session, "workbook", lambda: "C:/an/Session.xlsx")
    _store("Book1")
    book1 = NamedWb("Book1")
    book1.ActiveSheet = type("S", (), {"Name": "Q1 Answer"})()
    _by_name(monkeypatch, {22: ["Book1"]}, {22: FakeApp(Books([book1]))})
    monkeypatch.setattr(sheets, "_stored_specs", lambda t: {1: {"spec": {"v2": {}}, "wb": t.wb}})
    monkeypatch.setattr(jobs, "com", lambda fn, **k: fn())
    stored, active, path = pipeline.read_stored(None, server._default_target("session"))
    assert stored[1]["wb"] is book1 and active == "Q1 Answer" and path == "Book1"


# ------------------------------------------------------------------ fin_critic: target, default = latest answer's book
def _critic_wire(monkeypatch, fail=None):
    seen = []

    async def user(fn, what):
        return fn(None)

    def read_stored(h, target="session"):
        seen.append(target)
        if fail is not None:
            raise fail
        return {1: {"spec": {"v2": {"question": "q1"}}}}, "Q1 Answer", target if target != "session" else "C:/s.xlsx"
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "read_stored", read_stored)
    return seen


def _store_brief(workbook, q=1):
    answers.store({"workbook": workbook, "q": q, "question": f"q{q} in {workbook}",
                   "spec": {"v2": {"question": f"q{q}"}}, "critic_brief": {"purpose": workbook}})


def test_critic_defaults_to_the_workbook_of_the_latest_answer(home, monkeypatch):
    monkeypatch.setattr(session, "workbook", lambda: "C:/s.xlsx")
    _store_brief("C:/s.xlsx")
    _store_brief("Book1")
    seen = _critic_wire(monkeypatch)
    out = asyncio.run(server.fin_critic())
    assert seen == ["Book1"] and out["workbook"] == "Book1" and out["brief"] == {"purpose": "Book1"}
    assert asyncio.run(server.fin_critic(q=1, score=4.0))["stored"] is True
    assert answers.get("Book1", 1)["score"] == 4.0 and answers.get("C:/s.xlsx", 1)["score"] is None


def test_critic_takes_an_explicit_target(home, monkeypatch):
    _store_brief("Book1")
    _store_brief("C:/other.xlsx")
    seen = _critic_wire(monkeypatch)
    out = asyncio.run(server.fin_critic(target="Book1"))
    assert seen == ["Book1"] and out["brief"] == {"purpose": "Book1"}


def test_critic_on_a_closed_unsaved_book_falls_back_to_its_stored_answer(home, monkeypatch):
    from smartass.xl.host import InputError
    _store_brief("Book1")
    _store_brief("C:/x.xlsx", q=1)
    _store_brief("Book1", q=2)
    _critic_wire(monkeypatch, fail=InputError("Give an absolute path for target (got 'Book1')"))
    answers.store({"workbook": "Book1", "q": 2, "question": "latest", "spec": {}, "critic_brief": {"purpose": "B2"}})
    out = asyncio.run(server.fin_critic(q=1))
    assert out["workbook"] == "Book1" and out["q"] == 1


def test_critic_explicit_target_that_is_not_open_is_an_error(home, monkeypatch):
    from smartass.xl.host import InputError
    _store_brief("Book1")
    _critic_wire(monkeypatch, fail=InputError("Give an absolute path for target (got 'Book7')"))
    with pytest.raises(InputError):
        asyncio.run(server.fin_critic(target="Book7"))
