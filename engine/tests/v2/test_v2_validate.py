# engine/tests/v2/test_v2_validate.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.common import BlockOut, Line, evaluate
from smartass.model.context import Ctx
from smartass.model.expr import pick
from smartass.model.frame import Fact, Frame
from smartass.model.spec import AnalysisSpec
from smartass.model.validate import GuardError, guard_result, guard_spec, sign_findings


def spec(**q_over):
    d = spec_dict(blocks=[{"type": "kpi_tiles", "query": "fm", "metrics": ["Opex"]}])
    d["queries"][0].update(q_over)
    return AnalysisSpec.model_validate(d)


def test_clean_spec_passes_without_findings():
    assert guard_spec(spec(), resolver(), CLOSED) == []


@pytest.mark.parametrize("over, match", [
    ({"entity": "un.SY.ZZ"}, "unknown entity"),
    ({"metrics": ["Opex", "Capex"]}, "unknown metric or member 'Capex'"),
    ({"metrics": ["Opex", "Draft thing"]}, "is a draft"),
    ({"periods": "Oct 26"}, "after the last closed month Sep 26"),
])
def test_hard_guards_stop_with_a_fix(over, match):
    with pytest.raises(GuardError, match=match) as e:
        guard_spec(spec(**over), resolver(), CLOSED)
    assert e.value.fix


def test_preliminary_month_is_allowed_with_a_warning():
    found = guard_spec(spec(periods="Oct 26", preliminary=True), resolver(), CLOSED)
    assert [f.severity for f in found] == ["warning"] and "preliminary" in found[0].text


def test_missing_last_year_is_a_soft_warning():
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "columns": ["value"]}])
    d["queries"][0]["compare"] = []
    found = guard_spec(AnalysisSpec.model_validate(d), resolver(), CLOSED)
    assert [f.text for f in found] == ["query fm has no last-year comparison"]


def test_check_lines_must_be_zero():
    fr = Frame([Fact("FinMax", "fm", "a", "e", "Sep 26", "Actual", "", 5.0)])
    out = BlockOut("b0", "ranking", "Top", ["Actual"],
                   lines=[Line("check", "Check: children = parent", [pick({"Account": ("a",)})], kind="check")])
    with pytest.raises(GuardError, match="off by 5"):
        guard_result([out], evaluate([out], fr))


def test_sign_findings_flag_positive_costs():
    s = spec(metrics=["Opex"])
    assert sign_findings(Ctx(s, resolver(), CLOSED, facts())) == []
    flipped = Frame([Fact(**{**f.__dict__, "value": -f.value}) for f in facts().facts])
    found = sign_findings(Ctx(s, resolver(), CLOSED, flipped))
    assert [f.severity for f in found] == ["critical"] and "Opex" in found[0].text


def plan_spec(block, plan_periods="fy", periods="Jan 26..Sep 26"):
    d = spec_dict(blocks=[block])
    d["queries"] = [{"id": "fp", "cube": "FinPlan", "metrics": ["EBITDA"], "entity": "un.SY.2BUS",
                     "periods": periods, "plan_periods": plan_periods, "compare": ["budget", "forecast"]}]
    return AnalysisSpec.model_validate(d)


@pytest.mark.parametrize("block", [
    {"type": "table", "query": "fp", "columns": ["value", "budget", "var"]},
    {"type": "table", "query": "fp", "columns": ["value", "forecast", "var_pct"]},
    {"type": "kpi_tiles", "query": "fp", "metrics": ["EBITDA"], "compare": "budget"},
    {"type": "bridge", "query": "fp", "total": "EBITDA", "steps": ["EBITDA"], "base": "forecast"},
    {"type": "drivers", "query": "fp", "metric": "EBITDA", "base": "budget"},
    {"type": "ranking", "query": "fp", "metric": "EBITDA", "compare": "budget"},
])
def test_plan_comparison_over_different_months_stops(block):
    with pytest.raises(GuardError, match="Actual covers Jan 26..Sep 26 but .* covers Jan 26..Dec 26") as e:
        guard_spec(plan_spec(block), resolver(), CLOSED)
    assert "like-for-like" in e.value.fix and "own query" in e.value.fix


def test_plan_value_columns_and_trend_over_other_months_are_allowed():
    for block in ({"type": "table", "query": "fp", "columns": ["value", "forecast"]},
                  {"type": "trend", "query": "fp", "metrics": ["EBITDA"], "series": ["value", "budget"]}):
        guard_spec(plan_spec(block), resolver(), CLOSED)


def test_plan_comparison_over_the_same_months_is_allowed():
    guard_spec(plan_spec({"type": "table", "query": "fp", "columns": ["value", "budget", "var"]},
                         plan_periods="Jan 26..Sep 26"), resolver(), CLOSED)
