import asyncio

import pytest

from smartview_mcp import tools_core


def _gate(outcome, recipe=None, checklist=None, warnings=()):
    return {"question": "q", "suggested_outcome": outcome, "questions": ["Which measure?"] if outcome == "ambiguous" else [],
            "reason": None, "proxy": None, "warnings": list(warnings), "recipe": recipe, "closed_month": "Oct 26",
            "checklist": checklist or {"period": {"value": "YTD Oct 26", "source": "convention"},
                                       "entity": {"value": "Wholesale", "source": "question"},
                                       "currency_sign": {"value": "NOK", "source": "learned"}},
            "next": "..."}


def test_a_clear_question_is_answered_in_one_call(monkeypatch):
    seen = {}

    async def gate(question, answers=None):
        seen["answers"] = answers
        return _gate("clear", {"name": "pnl", "params": {"area": "un.SY.2WHS"}}, warnings=["note w"])

    async def spec_answer(spec, followup, q, target, allow_archive, detail="lean"):
        seen["spec"] = spec
        return {"ok": True}
    monkeypatch.setattr(tools_core, "gate", gate)
    monkeypatch.setattr(tools_core, "_answer_spec", spec_answer)
    out = asyncio.run(tools_core.answer(question="q", answers={"measure": "EBITDA"}))
    assert out == {"ok": True} and seen["answers"] == {"measure": "EBITDA"}
    spec = seen["spec"]
    assert spec["recipe"] == {"name": "pnl", "params": {"area": "un.SY.2WHS"}} and spec["question"] == "q"
    assert spec["assumptions"] == ["Period: YTD Oct 26 - convention.", "Currency sign: NOK - learned.", "note w"]
    assert spec["gate"]["answers"] == {"measure": "EBITDA"}


@pytest.mark.parametrize("outcome", ["ambiguous", "infeasible"])
def test_an_unclear_question_returns_the_gate_and_pulls_nothing(monkeypatch, outcome):
    async def gate(question, answers=None):
        return _gate(outcome)

    async def never(*a, **k):
        raise AssertionError("must not answer")
    monkeypatch.setattr(tools_core, "gate", gate)
    monkeypatch.setattr(tools_core, "_answer_spec", never)
    out = asyncio.run(tools_core.answer(question="q"))
    assert out["suggested_outcome"] == outcome


def test_clear_without_a_recipe_asks_for_a_spec(monkeypatch):
    async def gate(question, answers=None):
        return _gate("clear", None)
    monkeypatch.setattr(tools_core, "gate", gate)
    out = asyncio.run(tools_core.answer(question="q"))
    assert "spec" in out["next"]


def test_exactly_one_of_question_spec_followup():
    from smartass.xl.host import InputError
    with pytest.raises(InputError):
        asyncio.run(tools_core.answer(question="q", spec={"version": 1}))
    with pytest.raises(InputError):
        asyncio.run(tools_core.answer())


def test_http_answer_accepts_question_and_answers():
    from smartview_mcp import httpapi
    assert {"question", "answers"} <= httpapi._FIELDS["/answer"]
