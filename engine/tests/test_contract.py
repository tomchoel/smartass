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


def test_a_static_check_needs_no_expr_and_exactly_one_of_the_two():
    s = spec()
    s["layout"]["checks"] = [{"label": "Period is closed", "static": "OK", "detail": "Sep 26 vs Oct 26"}]
    assert contract.validate(s).layout.checks[0].static == "OK"
    s["layout"]["checks"] = [{"label": "both", "static": "OK", "expr": "ytd-eacc"}]
    assert "exactly one of expr or static" in errors(s)
    s["layout"]["checks"] = [{"label": "neither"}]
    assert "exactly one of expr or static" in errors(s)


def _sec_titles(raw):
    return [x["title"] for x in raw["layout"]["sections"]]


def test_patch_adds_a_line_to_a_section_by_title_and_a_pull_by_id():
    base = spec()
    title = base["layout"]["sections"][1]["title"]
    fp = {"id": "fp", "connection": "FinPlan", "rows": [{"dimension": "Account", "members": ["EBITDA"]}],
          "columns": [{"dimension": "Scenario", "members": ["Actual"]}]}
    out = contract.patched(base, {"pulls": [fp], "layout": {"sections": [
        {"title": title, "lines": [{"key": "fp", "label": "EBITDA YTD (FinPlan)",
                                    "cells": [{"pull": "fp", "row": ["EBITDA"], "col": ["Actual"]}]}]}]}})
    sec = out["layout"]["sections"][1]
    assert sec["lines"][-1]["key"] == "fp" and len(sec["lines"]) == len(base["layout"]["sections"][1]["lines"]) + 1
    assert _sec_titles(out) == _sec_titles(base) and [p["id"] for p in out["pulls"]] == ["m", "ytd", "fp"]
    contract.validate(out)


def test_patch_appends_notes_replaces_title_and_drops_by_identity():
    base = spec()
    base["layout"]["notes"] = ["first"]
    k = base["layout"]["sections"][1]["lines"][0]["key"]
    out = contract.patched(base, {"layout": {"title": "New", "notes": {"append": ["second"]},
                                             "checks": [{"label": base["layout"]["checks"][0]["label"], "drop": True}],
                                             "sections": [{"title": base["layout"]["sections"][1]["title"],
                                                           "lines": [{"key": k, "drop": True}]}]}})
    assert out["layout"]["title"] == "New" and out["layout"]["notes"] == ["first", "second"]
    assert k not in [ln["key"] for ln in out["layout"]["sections"][1]["lines"]]
    out2 = contract.patched(base, {"layout": {"notes": ["only"]}})
    assert out2["layout"]["notes"] == ["only"]                                  # a plain list replaces


def test_patch_appends_a_new_section_and_merges_checks_by_label():
    base = spec()
    base["layout"]["checks"] = [{"label": "c1", "expr": "eacc"}]
    out = contract.patched(base, {"layout": {"sections": [{"title": "By business area", "lines": []}],
                                             "checks": [{"label": "c1", "columns": [0]},
                                                        {"label": "c2", "static": "OK"}]}})
    assert _sec_titles(out)[-1] == "By business area"
    assert out["layout"]["checks"] == [{"label": "c1", "expr": "eacc", "columns": [0]},
                                       {"label": "c2", "static": "OK"}]


def test_patch_on_a_title_used_twice_is_refused():
    base = spec()
    base["layout"]["sections"][1]["title"] = base["layout"]["sections"][0]["title"]
    with pytest.raises(ValueError, match="used twice"):
        contract.patched(base, {"layout": {"sections": [{"title": base["layout"]["sections"][0]["title"],
                                                         "lines": []}]}})


def test_patch_cannot_change_recipe_params():
    base = {**spec(), "recipe": {"name": "monthly_trend", "params": {"entity": "un.SY"}}}
    with pytest.raises(ValueError, match="recipe parameters cannot be patched"):
        contract.patched(base, {"recipe": {"params": {"entity": "un.SY.2BUS"}}})


def test_recipe_field_and_outline():
    s = {**spec(), "recipe": {"name": "monthly_trend", "params": {"entity": "un.SY"}}}
    assert contract.validate(s).recipe.name == "monthly_trend"
    o = contract.outline(s)
    assert o["pulls"] == [{"id": "m", "connection": "FinMax"}, {"id": "ytd", "connection": "FinMax"}]
    assert o["sections"][0]["title"] == s["layout"]["sections"][0]["title"]
    assert o["columns"] == [c.get("header", "") for c in s["layout"]["columns"]]
    assert o["recipe"] == "monthly_trend"


@pytest.mark.parametrize("text,bad", [
    ("Konsulenter = kontoene C60080 og C60090 (valgt av meg)", True),
    ("Kostnader hittil i år", True),
    ("Periode: september", False),
    ("Beløp i NOK", True),
    ("Supplier 'Bøler Bygg AS' invoiced most", False),
    ("Entity un.SY; consultants = ac.C60080 + ac.C60090", False),
    ("Actual vs budget, NOK; the user's request", False)])
def test_english_text_and_quoted_names_pass(text, bad):
    assert contract.non_english(text) is bad


def test_sheet_texts_name_their_paths():
    s = {**spec(), "assumptions": ["a"], "deviation": "d"}
    paths = dict(contract.sheet_texts(s))
    assert paths["assumptions[0]"] == "a" and paths["deviation"] == "d"
    assert paths["layout.title"] == s["layout"]["title"] and "question" not in paths
    assert any(p.startswith("layout.columns[") for p in paths)


# ------------------------------------------------------------------ B3: extras (topic sheets)
def _extra(topic="By project", title="By project"):
    return {"topic": topic, "layout": {"title": title, "columns": [{"header": "a"}],
                                       "sections": [{"title": "S", "lines": [{"key": "x", "label": "X",
                                                                              "cells": [{"value": 1}]}]}],
                                       "checks": [{"label": "c", "expr": "x-x"}]}}


def test_extras_merge_by_topic_in_followups():
    s = spec(extras=[_extra(), _extra("By month")])
    contract.validate(s)
    out = contract.patched(s, {"extras": [{"topic": "By project", "layout": {"title": "Projects"}},
                                          {"topic": "By month", "drop": True}]})
    assert [(x["topic"], x["layout"]["title"]) for x in out["extras"]] == [("By project", "Projects")]
    assert ("extras[0].layout.title", "By project") in contract.sheet_texts(s)
    assert contract.outline(s)["extras"] == ["By project", "By month"]
    assert "used twice or reserved" in errors(spec(extras=[_extra("Answer")]))
    assert "used twice or reserved" in errors(spec(extras=[_extra(), _extra()]))
    assert "topic" in errors(spec(extras=[_extra("by project")]))                       # a capital first
    bad = _extra()
    bad["layout"]["checks"][0]["expr"] = "ghost"
    assert "extras 'By project'" in errors(spec(extras=[bad]))
