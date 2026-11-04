from svxml_util import EchoTransport, atlas_with_dimensions

from smartass.data.provider import SvHttpProvider
from smartass.model.requests import GridRequest
from smartview_mcp import svhttp


def provider(tmp_path, value):
    sv = svhttp.SvHttp(EchoTransport(value), home="https://plan1.example",
                       db_path=atlas_with_dimensions(tmp_path / "a.db"))
    return SvHttpProvider(sv)


def test_finmax_grid_becomes_facts(tmp_path):
    seen = []

    def value(rows, cols, pov):
        seen.append((rows, cols, pov))
        return None if rows[1] == "E2" else 10.0

    p = provider(tmp_path, value)
    req = GridRequest("fm", "FinMax", ("ac.C30", "ac.C70"), ("E1", "E2"), ("Aug 26", "Sep 26"), "Actual", "",
                      "Periodic")
    got = p.fetch(req)
    assert got.route == "http" and got.warnings == []
    assert len(got.facts) == 4                                   # E2 is #Missing everywhere: no facts
    f = got.facts[0]
    assert (f.cube, f.query, f.entity, f.scenario, f.version, f.balance, f.value) == (
        "FinMax", "fm", "E1", "Actual", "", "Periodic", 10.0)
    assert {x.period for x in got.facts} == {"Aug 26", "Sep 26"}
    assert any("Actual" in pov and "NOK" in pov and "Periodic" in pov for _, _, pov in seen)


def test_finplan_splits_years_and_relabels_periods(tmp_path):
    p = provider(tmp_path, lambda rows, cols, pov: 1.0)
    req = GridRequest("fp", "FinPlan", ("EBITDA",), ("2BUS Business",), ("Dec 25", "Jan 26"), "Target", "Final",
                      "Periodic")
    got = p.fetch(req)
    assert sorted(f.period for f in got.facts) == ["Dec 25", "Jan 26"]
    assert {(f.scenario, f.version) for f in got.facts} == {("Target", "Final")}


def test_error_cells_are_warnings_not_numbers(tmp_path):
    p = provider(tmp_path, lambda rows, cols, pov: "#NoAccess" if cols[0] == "Sep 26" else 2.0)
    req = GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Aug 26", "Sep 26"), "Actual", "", "Periodic")
    got = p.fetch(req)
    assert [f.period for f in got.facts] == ["Aug 26"]
    assert got.warnings and "#NoAccess" in got.warnings[0]


def test_closed_months_per_cube(tmp_path):
    p = provider(tmp_path, lambda *a: 1.0)
    assert p.closed_months() == {"FinMax": (2026, 9), "FinLite": (2026, 9), "FinPlan": (2026, 9)}
