# engine/tests/v2/test_v2_context.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.context import Ctx
from smartass.model.series import months, scenario
from smartass.model.spec import AnalysisSpec, Query


def ctx():
    return Ctx(AnalysisSpec.model_validate(spec_dict()), resolver(), CLOSED, facts())


def test_scenarios_per_cube():
    r = resolver()
    assert scenario("FinMax", "ly", r) == ("Actual", "")
    assert scenario("FinPlan", "value", r) == ("Actual", "Working")
    assert scenario("FinPlan", "budget", r) == ("Target", "Final")
    assert scenario("FinPlan", "forecast", r) == ("Forecast Q326", "Final")
    with pytest.raises(ValueError, match="only in FinPlan"):
        scenario("FinMax", "budget", r)


def test_ly_shifts_months_back_a_year():
    q = Query(id="fm", cube="FinMax", metrics=["Opex"], entity="E", periods="closed-ytd")
    assert months(q, "ly", CLOSED)[0] == (2025, 1)
    assert months(q, "value", CLOSED, base=[(2026, 3)]) == [(2026, 3)]


def test_pick_value_and_ly_and_formula():
    c = ctx()
    q = c.query("fm")
    assert c.value(c.pick(q, "Opex", "value")) == -85.0
    assert c.value(c.pick(q, "Opex", "ly")) == -77.0
    assert c.value(c.pick(q, "Opex", "value", "un.SY.4BUS_B")) == -75.0
    assert c.value(c.pick(q, "Opex", "value", ("un.SY.4BUS_B", "un.SY.G910"))) == -85.0
    assert c.pick(q, "Opex", "value").to_excel("Data_Q1", str) == (
        'SUM(SUMIFS(Data_Q1[Value],Data_Q1[Cube],"FinMax",Data_Q1[Query],"fm",'
        'Data_Q1[Account],{"ac.C50S","ac.C60"},Data_Q1[Entity],"un.SY.2BUS",'
        'Data_Q1[Period],"Sep 26",Data_Q1[Scenario],"Actual",'
        'Data_Q1[Balance],"Periodic"))')


def test_unknown_series_and_query():
    for cube in ("FinMax", "FinPlan"):
        with pytest.raises(ValueError, match="unknown series 'bogus'"):
            scenario(cube, "bogus", resolver())
    with pytest.raises(KeyError, match="unknown query 'zz'"):
        ctx().query("zz")
