import re

import pytest
from svxml_util import answer

from smartview_mcp import atlas, svhttp, svxml
from smartview_mcp.grid import Axis
from smartview_mcp.host import InputError


def _store(db, conn, dims):
    for d in dims:
        atlas.store(db, conn, d, [{"name": d, "label": d, "parent": None, "depth": 0, "order": 0}], 0.1)


def test_atlas_dimensions_sorted_per_connection(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    _store(db, "FinMax", ["All Time Periods", "All Accounts"])
    _store(db, "FinPlan", ["Account"])
    assert atlas.dimensions(db, "FinMax") == ["All Accounts", "All Time Periods"]
    assert atlas.dimensions(db, "FinLite") == []


SOURCES = ('<res_GetProvisionedDataSources><Product id="HP"><Server name="FinMax" '
           'context="https://plan1.example/HyperionPlanning/SmartView" appName="FinMax"/></Product>'
           '<Product id="HP"><Server name="FinPlan_x" context="https://strat.example/HyperionPlanning/SmartView" '
           'appName="FinPlan"/></Product></res_GetProvisionedDataSources>')
EXPIRED = '<exception errcode="1000" type="error"><desc>session expired</desc></exception>'
INVALID = ('<exception errcode="1000" type="error"><desc>Grid is Invalid. It should have at least 2 members valid '
           'for selected plan type.</desc></exception>')


class FakeTransport:
    """Answers like the provider (shapes from the live probe; figures invented)."""

    def __init__(self, *grid_answers):
        self.sent, self.answers = [], list(grid_answers)

    def post(self, base, body):
        svxml.check_read_only(body)
        tag = re.match(r"\s*<req_(\w+)>", body).group(1)
        self.sent.append((base, tag))
        if tag == "GetProvisionedDataSources":
            return SOURCES
        if tag in ("ConnectToProvider", "OpenApplication"):
            return f"<res_{tag}><sID>sid{len(self.sent)}</sID></res_{tag}>"
        if tag == "OpenCube":
            return "<res_OpenCube></res_OpenCube>"
        return self.answers.pop(0)

    def tags(self):
        return [t for _, t in self.sent]


GRID = [[None, "2WHS Wholesale (un.SY)", "Actual"], [None, "Aug 26", "Sep 26"],
        ["C30 REVENUES (ac)", 100.0, 110.0], ["EBITDA EBITDA (ac)", 50.0, None]]


@pytest.fixture
def db_path(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    _store(db, "FinMax", ["All Accounts", "All Entities", "All Scenarios", "All Time Periods"])
    _store(db, "FinPlan", ["Account", "Entity", "Period", "Scenario", "Version", "Years"])
    return tmp_path / "a.db"


def _q(sv, conn="FinMax"):
    return sv.retrieve(conn, {"Entities": "un.SY.2WHS", "All Scenarios": "Actual"},
                       [Axis("All Accounts", ["ac.C30", "ac.EBITDA"])], [Axis("All Time Periods", ["Aug 26", "Sep 26"])],
                       True, False)


def test_retrieve_returns_the_com_contract(db_path):
    t = FakeTransport(answer(GRID))
    lay, p, secs = _q(svhttp.SvHttp(t, home="https://plan1.example", db_path=db_path))
    assert lay.pov == [("All Entities", "un.SY.2WHS"), ("All Scenarios", "Actual")]
    assert p.rows == [["ac.C30", 100.0, 110.0], ["ac.EBITDA", 50.0, None]]
    assert p.row_labels == [["C30 REVENUES (ac)"], ["EBITDA EBITDA (ac)"]] and secs >= 0
    assert t.tags() == ["GetProvisionedDataSources", "ConnectToProvider", "OpenApplication", "OpenCube", "ExecuteGrid"]
    assert t.sent[-1][0] == "https://plan1.example"


def test_session_is_reused_and_finplan_goes_to_its_pod(db_path):
    t = FakeTransport(answer(GRID), answer(GRID), answer([[None, "x"], [None, "Sep"], ["Revenue", 1.0]]))
    sv = svhttp.SvHttp(t, home="https://plan1.example", db_path=db_path)
    _q(sv)
    _q(sv)
    assert t.tags().count("ConnectToProvider") == 1
    sv.retrieve("FinPlan", {"Entity": "2WHS Wholesale"}, [Axis("Account", ["Revenue"])], [Axis("Period", ["Sep"])],
                True, False)
    assert t.sent[-1] == ("https://strat.example", "ExecuteGrid") and t.tags().count("ConnectToProvider") == 2


def test_one_reconnect_then_a_clear_error(db_path):
    t = FakeTransport(EXPIRED, answer(GRID))
    _q(svhttp.SvHttp(t, home="https://plan1.example", db_path=db_path))
    assert t.tags().count("ConnectToProvider") == 2
    t2 = FakeTransport(EXPIRED, EXPIRED)
    with pytest.raises(svhttp.SvHttpError, match="session expired"):
        _q(svhttp.SvHttp(t2, home="https://plan1.example", db_path=db_path))


def test_a_deterministic_grid_error_does_not_spend_a_reconnect(db_path):
    t = FakeTransport(INVALID, answer(GRID))
    with pytest.raises(svhttp.SvHttpError, match="Grid is Invalid"):
        _q(svhttp.SvHttp(t, home="https://plan1.example", db_path=db_path))
    assert t.tags().count("ConnectToProvider") == 1 and t.tags().count("ExecuteGrid") == 1


def test_input_errors_and_missing_atlas(db_path, tmp_path):
    sv = svhttp.SvHttp(FakeTransport(), home="https://plan1.example", db_path=db_path)
    with pytest.raises(InputError, match="unknown dimension 'All Bogus'"):
        sv.retrieve("FinMax", {"All Bogus": "x"}, [Axis("All Accounts", ["a"])], [Axis("All Time Periods", ["p"])],
                    True, False)
    with pytest.raises(svhttp.SvHttpError, match="no HTTP route"):
        sv.retrieve("Other", {}, [], [], True, False)
    empty = svhttp.SvHttp(FakeTransport(), home="https://plan1.example", db_path=tmp_path / "empty.db")
    with pytest.raises(svhttp.SvHttpError, match="atlas has no dimensions"):
        _q(empty, "FinLite")
