# engine/tests/v2/test_v2_requests.py
from v2util import CLOSED, resolver, spec_dict

from smartass.model.blocks import BUILDERS
from smartass.model.requests import requests
from smartass.model.spec import AnalysisSpec


def test_registry_covers_the_computed_blocks():
    assert set(BUILDERS) == {"table", "kpi_tiles", "trend", "bridge", "ranking", "matrix", "drivers"}


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
