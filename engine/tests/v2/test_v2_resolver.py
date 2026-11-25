# engine/tests/v2/test_v2_resolver.py
import pytest
from v2util import resolver

from smartass.model.resolver import UnknownMetric


def test_metric_members_per_cube_case_insensitive():
    r = resolver()
    assert r.members("opex", "FinMax") == ("ac.C50S", "ac.C60")
    assert r.members("Opex", "FinPlan") == ("OPEX",)


def test_raw_member_is_its_own_metric():
    r = resolver()
    assert r.members("ac.C30", "FinMax") == ("ac.C30",)
    assert r.sign("ac.C30") == "neutral" and r.status("ac.C30") == "approved"


def test_unknown_metric_and_missing_cube():
    r = resolver()
    with pytest.raises(UnknownMetric):
        r.members("Capex", "FinMax")
    with pytest.raises(UnknownMetric, match="FinPlan"):
        r.members("Draft thing", "FinPlan")


def test_tree_labels_and_status():
    r = resolver()
    assert r.children("un.SY.2BUS", "FinMax") == ("un.SY.4BUS_B", "un.SY.G910")
    assert r.children("un.SY.B1", "FinMax") == ()
    assert r.label("un.SY.2BUS", "FinMax") == "2BUS Business Norway"
    assert r.label("un.SY.B1", "FinMax") == "un.SY.B1"
    assert r.exists("un.SY.B2", "FinMax") and not r.exists("un.SY.ZZ", "FinMax")
    assert r.status("Draft thing") == "draft"
    assert r.materiality("Opex") == (5.0, 0.05)
    assert r.latest_forecast() == "Forecast Q326"


def test_accumulated_balance_with_sum_aggregation_is_refused():
    from smartass.model.resolver import MetricDef
    with pytest.raises(ValueError, match="Accumulated.*sum"):
        MetricDef("Bad", {"FinMax": ("ac.X",)}, balance="Accumulated", aggregation="sum")
    MetricDef("Ok", {"FinMax": ("ac.X",)}, balance="Accumulated", aggregation="last")


def test_raw_level_member_inherits_metric_balance_and_aggregation():
    r = resolver()
    # Querying raw 'ac.STA20Y' instead of metric 'FTE' must still use Accumulated / last!
    assert r.balance("ac.STA20Y") == "Accumulated"
    assert r.aggregation("ac.STA20Y") == "last"
