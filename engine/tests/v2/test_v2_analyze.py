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
