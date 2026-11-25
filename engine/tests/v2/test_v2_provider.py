from svxml_util import EchoTransport, atlas_with_dimensions

from smartass.data.provider import SvHttpProvider
from smartass.model.requests import GridRequest
from smartass.epm import svhttp


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


def test_valid_empty_period_columns_say_no_data_not_did_not_know(tmp_path):
    from smartass.knowledge import atlas
    from svxml_util import FINMAX_DIMS, FINPLAN_DIMS, SOURCES, answer
    db_path = tmp_path / "atlas_periods.db"
    with atlas.connect(db_path) as db:
        for conn, dims in (("FinMax", FINMAX_DIMS), ("FinPlan", FINPLAN_DIMS)):
            for d in dims:
                atlas.store(db, conn, d, [{"name": d, "label": d, "parent": None, "depth": 0, "order": 0}], 0.1)
        atlas.store(db, "FinMax", "All Time Periods", [
            {"name": "All Time Periods", "label": "All Time Periods", "parent": None, "depth": 0, "order": 0},
            {"name": "Jan 25", "label": "Jan 25", "parent": "All Time Periods", "depth": 1, "order": 1},
            {"name": "Feb 25", "label": "Feb 25", "parent": "All Time Periods", "depth": 1, "order": 2},
            {"name": "Apr 25", "label": "Apr 25", "parent": "All Time Periods", "depth": 1, "order": 3},
        ], 0.1)

    class EmptyColsTransport:
        def post(self, base, body):
            if "GetProvisionedDataSources" in body:
                return SOURCES
            if "ConnectToProvider" in body:
                return "<res_ConnectToProvider><sID>s1</sID></res_ConnectToProvider>"
            if "OpenApplication" in body:
                return "<res_OpenApplication><sID>s1</sID></res_OpenApplication>"
            if "OpenCube" in body:
                return "<res_OpenCube></res_OpenCube>"
            # Live EPM returns None (blank) for unpopulated columns Jan 25, Feb 25 while Apr 25 has a float
            return answer([
                [None, None, "Actual", "NOK", "Periodic"],
                [None, None, "Jan 25", "Feb 25", "Apr 25"],
                ["C30 REVENUES (ac)", "T370 (un.SY)", None, None, 10.0],
            ])

    sv = svhttp.SvHttp(EmptyColsTransport(), home="https://plan1.example", db_path=db_path)
    prov = SvHttpProvider(sv)
    req = GridRequest("fm", "FinMax", ("ac.C30",), ("un.SY.T370",), ("Jan 25", "Feb 25", "Apr 25"), "Actual", "", "Periodic")
    got = prov.fetch(req)
    assert not any("did not know" in w for w in got.warnings)
    assert any("no data" in w and "Jan 25" in w and "Feb 25" in w for w in got.warnings)
