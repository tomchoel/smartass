# engine/tests/v2/test_v2_analyze.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.analyze import compute, plan_requests
from smartass.model.spec import AnalysisSpec
from smartass.model.validate import GuardError


def test_end_to_end_base_spec():
    s = AnalysisSpec.model_validate(spec_dict())
    res = compute(s, resolver(), CLOSED, facts())
    assert [o.type for o in res.outs] == ["kpi_tiles", "table", "bridge", "ranking", "drivers"]
    assert res.key_values["Key figures · EBITDA"] == {"Actual": 65.0, "Last year": 63.0, "Variance": 2.0,
                                                     "Variance %": pytest.approx(2 / 63)}
    assert res.key_values["EBITDA bridge · Other / unexplained"] == {"Amount": 0.0}
    assert not any("Check" in k for k in res.key_values)
    keys = {m["key"] for m in res.material}
    assert {"b1.m1", "b4.top", "b4.d0", "b4.d0_0", "b4.d0_1"} <= keys      # Opex lines over 5 NOK / 5 %
    assert "b0.k0" not in keys                                             # EBITDA +2 is under 5 MNOK
    assert res.findings == []


def test_plan_requests_runs_the_guards_first():
    d = spec_dict()
    d["queries"][0]["entity"] = "un.SY.ZZ"
    with pytest.raises(GuardError, match="unknown entity"):
        plan_requests(AnalysisSpec.model_validate(d), resolver(), CLOSED)
    assert len(plan_requests(AnalysisSpec.model_validate(spec_dict()), resolver(), CLOSED)) == 2


def test_key_values_are_capped():
    s = AnalysisSpec.model_validate(spec_dict())
    res = compute(s, resolver(), CLOSED, facts(), kv_limit=3)
    assert len(res.key_values) == 3


def test_identical_block_titles_and_labels_do_not_collide_in_key_values():
    from smartass.model.frame import Fact
    d = spec_dict(blocks=[{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]},
                          {"type": "kpi_tiles", "query": "fp", "metrics": ["EBITDA"], "compare": "budget"},
                          {"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA", "EBITDA"]}])
    d["queries"].append({"id": "fp", "cube": "FinPlan", "metrics": ["EBITDA"], "entity": "un.SY.2BUS",
                         "periods": "closed", "compare": ["budget"]})
    fr = facts()
    fr.facts += [Fact("FinPlan", "fp", "EBITDA", "un.SY.2BUS", "Sep 26", "Actual", "Working", 70.0),
                 Fact("FinPlan", "fp", "EBITDA", "un.SY.2BUS", "Sep 26", "Target", "Final", 80.0)]
    res = compute(AnalysisSpec.model_validate(d), resolver(), CLOSED, fr)
    assert res.key_values["Key figures · EBITDA"]["Actual"] == 65.0
    assert res.key_values["Key figures (2) · EBITDA"]["Actual"] == 70.0
    assert res.key_values["Key figures (2) · EBITDA"]["Budget"] == 80.0
    assert "Key figures (3) · EBITDA" in res.key_values and "Key figures (3) · EBITDA (2)" in res.key_values
