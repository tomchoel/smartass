"""Answer spec v1 contract (core/contract.py): the invariants the old layout grammar never checked."""
import copy
import json
from pathlib import Path

import pytest
from pydantic import ValidationError

from smartview_mcp.core import contract

B02 = json.loads((Path(__file__).parent / "fixtures" / "b02_answer_spec.json").read_text(encoding="utf-8"))


def spec(**over):
    s = copy.deepcopy(B02)
    s.update(over)
    return s


def errors(s) -> str:
    with pytest.raises(ValidationError) as e:
        contract.validate(s)
    return str(e.value)


def test_b02_live_matrix_layout_is_a_valid_v1_spec():
    out = contract.validate(spec())
    assert out.gate.outcome == "clear" and [p.id for p in out.pulls] == ["m", "ytd"]


def test_layout_cross_references_are_checked():
    s = spec()
    lines = s["layout"]["sections"][1]["lines"]
    lines[0]["cells"][0]["pull"] = "nope"                       # unknown pull id
    lines[0]["cells"][2] = {"var": [0, 9]}                      # column outside the table
    s["layout"]["checks"][0]["expr"] = "ytd-ghost"              # unknown line in a check
    msg = errors(s)
    assert "unknown pull 'nope'" in msg and "var needs column indexes" in msg and "unknown line 'ghost'" in msg


def test_repeat_template_keys_count_as_lines():
    s = spec()
    s["layout"]["sections"][0]["lines"][0]["cells"][0] = {"sum": ["m0", "m8"]}   # expanded from 'm{i}'
    contract.validate(s)


def test_a_cell_has_exactly_one_kind():
    s = spec()
    s["layout"]["sections"][1]["lines"][0]["cells"][2] = {"var": [0, 1], "sum": "m*"}
    assert "exactly one of" in errors(s)


@pytest.mark.parametrize("gate, want", [
    ({"outcome": "ambiguous"}, "1-3 clarifying questions"),
    ({"outcome": "ambiguous", "questions": ["a", "b", "c", "d"]}, "at most 3"),
    ({"outcome": "clear", "questions": ["why?"]}, "asks no questions"),
    ({"outcome": "infeasible", "reason": "opex not by product"}, "reason and the proxy"),
    ({"outcome": "clear", "checklist": {"mood": "?"}}, "unknown checklist items"),
])
def test_gate_outcome_is_consistent(gate, want):
    assert want in errors(spec(gate=gate))


def test_ambiguous_question_pulls_nothing_and_clear_needs_a_layout():
    amb = {"outcome": "ambiguous", "questions": ["Budget Final or latest forecast?"]}
    assert "pulls nothing" in errors(spec(gate=amb))
    contract.validate({"question": "How are we doing on costs?", "gate": amb})
    assert "needs a layout" in errors(spec(layout=None))


def test_infeasible_answers_the_proxy():
    gate = {"outcome": "infeasible", "reason": "FinMax has no opex by product",
            "proxy": "contribution margin (revenue - COGS) per product"}
    assert contract.validate(spec(gate=gate)).gate.proxy.startswith("contribution margin")


def test_data_driver_cites_evidence_hypothesis_need_not():
    hyp = {"text": "June one-off low opex", "basis": "hypothesis"}
    contract.validate(spec(commentary={"drivers": [hyp]}))
    assert "cites its evidence" in errors(spec(commentary={"drivers": [{**hyp, "basis": "data"}]}))


def test_pull_dimension_used_twice_is_rejected():
    s = spec()
    s["pulls"][0]["pov"]["All Accounts"] = "ac.C30"
    assert "used twice: All Accounts" in errors(s)


def test_follow_up_merges_pulls_by_id_and_revalidates():
    fp = {"id": "fp", "connection": "FinPlan", "rows": [{"dimension": "Account", "members": ["EBITDA"]}],
          "columns": [{"dimension": "Scenario", "members": ["Actual"]}], "pov": {"Entity": "2BUS Business"}}
    out = contract.patch(spec(), {"pulls": [fp], "layout": {"title": "Patched"}})
    assert [p.id for p in out.pulls] == ["m", "ytd", "fp"] and out.layout.title == "Patched"
    assert out.layout.sections == contract.validate(spec()).layout.sections          # untouched parts stay
    with pytest.raises(ValidationError, match="unknown pull 'm'"):
        contract.patch(spec(), {"pulls": [{"id": "m", "drop": True}]})                # still referenced
