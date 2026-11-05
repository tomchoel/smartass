# engine/tests/v2/test_v2_requests.py
from v2util import CLOSED, resolver, spec_dict

from smartass.model.blocks import BUILDERS
from smartass.model.requests import requests
from smartass.model.spec import AnalysisSpec


def test_registry_covers_the_computed_blocks():
    assert set(BUILDERS) == {"table", "kpi_tiles", "trend", "bridge", "ranking", "matrix", "drivers", "lines", "concept",
                             "chart"}


def test_one_request_per_series_with_subtree_and_all_accounts():
    reqs = requests(AnalysisSpec.model_validate(spec_dict()), resolver(), CLOSED)
    assert [(x.query, x.periods, x.scenario) for x in reqs] == [("fm", ("Sep 26",), "Actual"),
                                                                ("fm", ("Sep 25",), "Actual")]
    assert reqs[0].accounts == ("ac.C30", "ac.C50S", "ac.C60", "ac.EBITDA")
    assert reqs[0].entities == ("un.SY.2BUS", "un.SY.4BUS_B", "un.SY.G910", "un.SY.B1", "un.SY.B2")


def test_no_entity_blocks_fetch_only_the_entity():
    d = spec_dict(blocks=[{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}])
    reqs = requests(AnalysisSpec.model_validate(d), resolver(), CLOSED)
    assert reqs[0].entities == ("un.SY.2BUS",)


def test_finplan_series_have_versions():
    d = spec_dict(blocks=[{"type": "kpi_tiles", "query": "fp", "metrics": ["EBITDA"], "compare": "budget"}])
    d["queries"] = [{"id": "fp", "cube": "FinPlan", "metrics": ["EBITDA"], "entity": "un.SY.2BUS",
                     "periods": "closed-ytd", "compare": ["budget", "forecast"]}]
    reqs = requests(AnalysisSpec.model_validate(d), resolver(), CLOSED)
    assert [(x.scenario, x.version) for x in reqs] == [("Actual", "Working"), ("Target", "Final"),
                                                       ("Forecast Q326", "Final")]
    assert len(reqs[0].periods) == 9


def test_plan_periods_apply_to_budget_and_forecast_only():
    from smartass.model.analyze import plan_requests
    d = spec_dict(blocks=[{"type": "table", "query": "fp", "metrics": ["EBITDA"], "columns": ["value", "forecast"]}])
    d["queries"] = [{"id": "fp", "cube": "FinPlan", "metrics": ["EBITDA"], "entity": "un.SY.2BUS",
                     "periods": "closed-ytd", "plan_periods": "fy", "compare": ["forecast"]}]
    reqs = plan_requests(AnalysisSpec.model_validate(d), resolver(), CLOSED)   # no GuardError for Oct-Dec
    assert [len(x.periods) for x in reqs] == [9, 12]
    assert reqs[1].scenario == "Forecast Q326"


def test_closed_month_per_cube():
    from smartass.model.analyze import compute, plan_requests
    from smartass.model.frame import Fact, Frame
    d = spec_dict(blocks=[{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]},
                          {"type": "kpi_tiles", "query": "fp", "metrics": ["EBITDA"], "compare": "budget"}])
    d["queries"].append({"id": "fp", "cube": "FinPlan", "metrics": ["EBITDA"], "entity": "un.SY.2BUS",
                         "periods": "closed-ytd", "compare": ["budget"]})
    closed = {"FinMax": (2026, 9), "FinPlan": (2026, 8)}
    allr = plan_requests(AnalysisSpec.model_validate(d), resolver(), closed)
    assert [x.periods for x in allr if x.query == "fm"] == [("Sep 26",), ("Sep 25",)]
    reqs = {(x.query, x.scenario, x.version): x.periods for x in allr}
    assert reqs[("fp", "Actual", "Working")][-1] == "Aug 26"
    assert len(reqs[("fp", "Target", "Final")]) == 8
    fr = Frame([Fact("FinPlan", "fp", "EBITDA", "un.SY.2BUS", "Aug 26", "Actual", "Working", 7.0),
                Fact("FinPlan", "fp", "EBITDA", "un.SY.2BUS", "Sep 26", "Actual", "Working", 900.0)])
    res = compute(AnalysisSpec.model_validate(d), resolver(), closed, fr)
    assert res.key_values["Key figures (2) · EBITDA"]["Actual"] == 7.0       # Sep 26 is after FinPlan's closed month
