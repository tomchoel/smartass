import http.client
import io
import re
import sqlite3
import threading
import types
import urllib.error
from urllib.parse import urlsplit

import pytest
from svxml_util import answer

from smartass.knowledge import atlas
from smartass.epm import epmhttp, oauth, svhttp, svxml
from smartass.core import config
from smartass.epm.grid import Axis
from smartass.xl.host import InputError


def _store(db, conn, dims, error=None):
    for d in dims:
        atlas.store(db, conn, d, [{"name": d, "label": d, "parent": None, "depth": 0, "order": 0}], 0.1, error)


def test_atlas_dimensions_sorted_per_connection(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    _store(db, "FinMax", ["All Time Periods", "All Accounts"])
    _store(db, "FinPlan", ["Account"])
    assert atlas.known_dimensions(db, "FinMax") == ["All Accounts", "All Time Periods"]
    assert atlas.known_dimensions(db, "FinLite") == []
    assert not hasattr(atlas, "dimensions")          # known_dimensions is the one list (it sees failed builds too)


def _age(db, conn, days):
    with db:
        db.execute("UPDATE build SET built_at = built_at - ? WHERE connection=?", (days * 86400, conn))


def test_atlas_built_ignores_age_but_not_errors_or_empty_builds(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    _store(db, "FinMax", ["All Accounts"])
    _store(db, "FinMax", ["All Entities"], error="timed out")
    atlas.store(db, "FinMax", "All Scenarios", [], 0.1)
    _age(db, "FinMax", 40)
    assert atlas.built(db, "FinMax", "All Accounts") and not atlas.usable(db, "FinMax", "All Accounts")
    assert not atlas.built(db, "FinMax", "All Entities")
    assert not atlas.built(db, "FinMax", "All Scenarios")
    assert not atlas.built(db, "FinMax", "All Bogus")


# Like the live list: three servers with appName FinMax (prod, plan7, TEST - listed last, so keying by app picked it)
SOURCES = ('<res_GetProvisionedDataSources><Product id="HP"><Server name="FinMax" '
           'context="https://plan1.example/HyperionPlanning/SmartView" appName="FinMax"/></Product>'
           '<Product id="HP"><Server name="Plan 7 Prod" context="https://plan7.example/HyperionPlanning/SmartView" '
           'appName="FinMax"/></Product>'
           '<Product id="HP"><Server appName="FinMax" name="plan1-test" '
           'context="https://plan1-test.example/HyperionPlanning/SmartView"/></Product>'
           '<Product id="HP"><Server name="FinPlan_for_DataLoad" '
           'context="https://strat.example/HyperionPlanning/SmartView" appName="FinPlan"/></Product>'
           '</res_GetProvisionedDataSources>')
ONLY_FINMAX = ('<res_GetProvisionedDataSources><Product id="HP"><Server name="FinMax" '
               'context="https://plan1.example/HyperionPlanning/SmartView" appName="FinMax"/></Product>'
               '</res_GetProvisionedDataSources>')
NO_SOURCES = "<res_GetProvisionedDataSources></res_GetProvisionedDataSources>"
EXPIRED = '<exception errcode="1000" type="error"><desc>session expired</desc></exception>'
INVALID = ('<exception errcode="1000" type="error"><desc>Grid is Invalid. It should have at least 2 members valid '
           'for selected plan type.</desc></exception>')


class FakeTransport:
    """Answers like the provider (shapes from the live probe; figures invented). `sources` are the successive
    GetProvisionedDataSources answers (the last one repeats); `cube` is the OpenCube answer."""

    def __init__(self, *grid_answers, sources=(SOURCES,), cube="<res_OpenCube></res_OpenCube>"):
        self.sent, self.bodies, self.answers = [], [], list(grid_answers)
        self.sources, self.cube = list(sources), cube

    def post(self, base, body):
        svxml.check_read_only(body)
        tag = re.match(r"\s*<req_(\w+)>", body).group(1)
        self.sent.append((base, tag))
        self.bodies.append(body)
        if tag == "GetProvisionedDataSources":
            return self.sources.pop(0) if len(self.sources) > 1 else self.sources[0]
        if tag in ("ConnectToProvider", "OpenApplication"):
            return f"<res_{tag}><sID>sid{len(self.sent)}</sID></res_{tag}>"
        if tag == "OpenCube":
            return self.cube
        return self.answers.pop(0)

    def tags(self):
        return [t for _, t in self.sent]

    def grid_sids(self):
        return [re.search(r"<sID>(.*?)</sID>", b).group(1) for b in self.bodies if b.startswith("<req_ExecuteGrid>")]


GRID = [[None, "2WHS Wholesale (un.SY)", "Actual"], [None, "Aug 26", "Sep 26"],
        ["C30 REVENUES (ac)", 100.0, 110.0], ["EBITDA EBITDA (ac)", 50.0, None]]
SHIFTED = [[None, "2WHS Wholesale (un.SY)", "Actual"], [None, "Aug 26", "Sep 26"],
           ["EBITDA EBITDA (ac)", 50.0, None], ["C48 Contribution (ac)", 1.0, 2.0]]
FINMAX_DIMS = ["All Accounts", "All Entities", "All Scenarios", "All Time Periods"]


@pytest.fixture
def db_path(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    _store(db, "FinMax", FINMAX_DIMS)
    _store(db, "FinPlan", ["Account", "Entity", "Period", "Scenario", "Version", "Years"])
    db.close()
    return tmp_path / "a.db"


def _sv(t, db_path):
    return svhttp.SvHttp(t, home="https://plan1.example", db_path=db_path)


def _q(sv, conn="FinMax"):
    return sv.retrieve(conn, {"Entities": "un.SY.2WHS", "All Scenarios": "Actual"},
                       [Axis("All Accounts", ["ac.C30", "ac.EBITDA"])], [Axis("All Time Periods", ["Aug 26", "Sep 26"])],
                       True, False)


def test_retrieve_returns_the_com_contract(db_path):
    t = FakeTransport(answer(GRID))
    lay, p, secs = _q(_sv(t, db_path))
    assert lay.pov == [("All Entities", "un.SY.2WHS"), ("All Scenarios", "Actual")]
    assert p.rows == [["ac.C30", 100.0, 110.0], ["ac.EBITDA", 50.0, None]]
    assert p.row_labels == [["C30 REVENUES (ac)"], ["EBITDA EBITDA (ac)"]] and secs >= 0
    assert t.tags() == ["GetProvisionedDataSources", "ConnectToProvider", "OpenApplication", "OpenCube", "ExecuteGrid"]
    assert t.sent[-1][0] == "https://plan1.example"


def test_a_shifted_grid_is_an_svhttp_error_so_com_is_used(db_path):
    with pytest.raises(svhttp.SvHttpError, match="different layout.*asked ac.C30, grid shows EBITDA EBITDA"):
        _q(_sv(FakeTransport(answer(SHIFTED)), db_path))


@pytest.mark.parametrize("matrix, msg", [
    (GRID[:-1], "fewer rows"),                                               # a row went missing
    ([r[:-1] for r in GRID], "fewer columns"),                               # a column went missing
    ([GRID[0] + ["All Products"]] + [r + [None] for r in GRID[1:]], "appended POV members"),
    ([*GRID[:2], ["C30 REVENUES (ac)", "abc", 110.0], GRID[3]], "non-numeric data"),
])
def test_a_grid_the_layout_cannot_trust_falls_back_to_com(db_path, matrix, msg):
    with pytest.raises(svhttp.SvHttpError, match=msg):
        _q(_sv(FakeTransport(answer(matrix)), db_path))


def test_missing_and_no_access_labels_are_fine_data(db_path):
    m = [*GRID[:2], ["C30 REVENUES (ac)", "#Missing", "#NoAccess"], GRID[3]]
    _, p, _ = _q(_sv(FakeTransport(answer(m)), db_path))
    assert p.rows[0] == ["ac.C30", None, "#NoAccess"]


def test_session_is_reused_and_finplan_goes_to_its_pod(db_path):
    t = FakeTransport(answer(GRID), answer(GRID), answer([[None, "2WHS Wholesale", "Scenario", "Version", "Years"],
                                                           [None, "Sep", None, None, None],
                                                           ["Revenue", 1.0, None, None, None]]))
    sv = _sv(t, db_path)
    _q(sv)
    _q(sv)
    assert t.tags().count("ConnectToProvider") == 1
    sv.retrieve("FinPlan", {"Entity": "2WHS Wholesale"}, [Axis("Account", ["Revenue"])], [Axis("Period", ["Sep"])],
                True, False)
    assert t.sent[-1] == ("https://strat.example", "ExecuteGrid") and t.tags().count("ConnectToProvider") == 2


def test_the_latest_sid_of_each_answer_is_used(db_path):
    with_sid = answer(GRID).replace("<res_ExecuteGrid>", "<res_ExecuteGrid><sID>grid&amp;2</sID>")
    t = FakeTransport(with_sid, answer(GRID), cube="<res_OpenCube><sID>cube1</sID></res_OpenCube>")
    sv = _sv(t, db_path)
    _q(sv)
    _q(sv)                                         # the second answer has no sID: grid&2 stays the session
    assert t.grid_sids() == ["cube1", "grid&amp;2"] and sv._idle["FinMax"] == ["grid&2"]


def test_one_reconnect_then_a_clear_error(db_path):
    t = FakeTransport(EXPIRED, answer(GRID))
    _q(_sv(t, db_path))
    assert t.tags().count("ConnectToProvider") == 2
    assert t.tags().count("GetProvisionedDataSources") == 2      # a reconnect also re-reads the data sources
    t2 = FakeTransport(EXPIRED, EXPIRED)
    with pytest.raises(svhttp.SvHttpError, match="session expired"):
        _q(_sv(t2, db_path))


def test_a_deterministic_grid_error_does_not_spend_a_reconnect(db_path):
    t = FakeTransport(INVALID.replace("<desc>", "<desc>Error: "), answer(GRID))
    with pytest.raises(svhttp.SvHttpError, match="Grid is Invalid"):
        _q(_sv(t, db_path))
    assert t.tags().count("ConnectToProvider") == 1 and t.tags().count("ExecuteGrid") == 1


def test_data_sources_are_not_cached_empty_and_refetched_once_for_a_missing_app(db_path):
    t = FakeTransport(answer(GRID), sources=(NO_SOURCES, ONLY_FINMAX, SOURCES))
    sv = _sv(t, db_path)
    with pytest.raises(svhttp.SvHttpError, match="no Smart View data source named 'FinMax'"):
        sv.base_for("FinMax", "FinMax")
    assert sv.base_for("FinMax", "FinMax") == "https://plan1.example"        # the empty list was not kept
    assert sv.base_for("FinPlan_for_DataLoad", "FinPlan") == "https://strat.example"   # not listed: fetched again
    assert t.tags().count("GetProvisionedDataSources") == 3
    t2 = FakeTransport(sources=(ONLY_FINMAX,))
    with pytest.raises(svhttp.SvHttpError, match="named 'FinPlan_for_DataLoad'"):
        _sv(t2, db_path).base_for("FinPlan_for_DataLoad", "FinPlan")
    assert t2.tags().count("GetProvisionedDataSources") == 1


def test_the_data_source_is_chosen_by_server_name_and_must_serve_the_app(db_path, monkeypatch):
    sv = _sv(FakeTransport(), db_path)
    assert sv.base_for("FinMax", "FinMax") == "https://plan1.example"        # not plan7, not the test pod
    assert sv.base_for("Plan 7 Prod", "FinMax") == "https://plan7.example"
    with pytest.raises(svhttp.SvHttpError, match="'FinPlan_for_DataLoad' serves app FinPlan, not FinMax"):
        sv.base_for("FinPlan_for_DataLoad", "FinMax")
    monkeypatch.setitem(svhttp.CUBES, "FinPlan", ("FinMax", "FinPlan", "FinPlan"))   # a mis-set server name
    with pytest.raises(svhttp.SvHttpError, match="serves app FinMax, not FinPlan"):
        sv.retrieve("FinPlan", {}, [Axis("Account", ["Revenue"])], [Axis("Period", ["Sep"])], True, False)


def _plus(sources, *servers):
    extra = "".join(f'<Product id="HP"><Server name="{n}" appName="{a}" '
                    f'context="{b}/HyperionPlanning/SmartView"/></Product>' for n, a, b in servers)
    return sources.replace("</res_GetProvisionedDataSources>", extra + "</res_GetProvisionedDataSources>")


def test_only_the_requested_server_name_can_be_ambiguous(db_path):
    unrelated = _plus(SOURCES, ("TELCONS", "TELCONS", "https://t1.example"), ("TELCONS", "TELCONS", "https://t2.example"))
    t = FakeTransport(answer(GRID), sources=(unrelated,))
    _q(_sv(t, db_path))                                    # a duplicated other server does not block FinMax
    assert t.sent[-1] == ("https://plan1.example", "ExecuteGrid")
    clash = _plus(SOURCES, ("FinMax", "FinMax", "https://plan9.example"))
    sv = _sv(FakeTransport(sources=(clash,)), db_path)
    with pytest.raises(svhttp.SvHttpError, match="'FinMax' is listed with different pods.*plan9.example.*ambiguous"):
        sv.base_for("FinMax", "FinMax")
    assert sv.base_for("FinPlan_for_DataLoad", "FinPlan") == "https://strat.example"


def test_the_test_pod_is_refused_unless_allowed(db_path, monkeypatch):
    t = FakeTransport(answer(GRID))
    sv = _sv(t, db_path)
    monkeypatch.setitem(svhttp.CUBES, "FinMax", ("plan1-test", "FinMax", "FinMax"))
    with pytest.raises(svhttp.SvHttpError, match="refused.*test environment.*plan1-test.example"):
        _q(sv)
    assert "ConnectToProvider" not in t.tags()
    monkeypatch.setattr(config, "SVHTTP_ALLOW_TEST", True)
    _q(sv)
    assert t.sent[-1] == ("https://plan1-test.example", "ExecuteGrid")


def test_server_names_come_from_the_org_defaults():
    assert config.get(config.ORG_SETTINGS, "svhttp.finmax_server") == "FinMax"
    assert config.get(config.ORG_SETTINGS, "svhttp.finplan_server") == "FinPlan_for_DataLoad"
    assert config.get(config.ORG_SETTINGS, "svhttp.allow_test") is False
    assert [v[0] for v in svhttp.CUBES.values()] == [config.SVHTTP_FINMAX_SERVER] * 2 + [config.SVHTTP_FINPLAN_SERVER]


def test_user_errors_stay_input_errors(db_path):
    sv = _sv(FakeTransport(), db_path)
    with pytest.raises(InputError, match="used more than once"):
        sv.retrieve("FinMax", {"All Accounts": "x"}, [Axis("All Accounts", ["a"])], [Axis("All Time Periods", ["p"])],
                    True, False)
    with pytest.raises(InputError, match="no members"):
        sv.retrieve("FinMax", {}, [Axis("All Accounts", [])], [Axis("All Time Periods", ["p"])], True, False)
    with pytest.raises(InputError, match="at least one row dimension"):
        sv.retrieve("FinMax", {}, [Axis("All Accounts", ["a"])], [], True, False)


def test_an_incomplete_or_stale_atlas_falls_back_to_com(db_path, tmp_path):
    sv = _sv(FakeTransport(), db_path)
    with pytest.raises(svhttp.SvHttpError, match="atlas is incomplete or stale for FinMax.*All Bogus"):
        sv.retrieve("FinMax", {"All Bogus": "x"}, [Axis("All Accounts", ["a"])], [Axis("All Time Periods", ["p"])],
                    True, False)
    with pytest.raises(svhttp.SvHttpError, match="no HTTP route"):
        sv.retrieve("Other", {}, [], [], True, False)
    with pytest.raises(svhttp.SvHttpError, match="atlas has no dimensions"):
        _q(_sv(FakeTransport(), tmp_path / "empty.db"), "FinLite")
    db = atlas.connect(db_path)
    atlas.store(db, "FinMax", "All Scenarios", [], 0.1, "timed out")   # a failed rebuild, as `smartass atlas build` writes it
    assert db.execute("SELECT count(*) FROM member WHERE dimension='All Scenarios'").fetchone()[0] == 0   # gone ...
    db.close()
    with pytest.raises(svhttp.SvHttpError, match="atlas is incomplete or stale for FinMax.*All Scenarios"):
        _q(sv)                                                        # ... but the build table still knows it
    db = atlas.connect(db_path)
    atlas.store(db, "FinPlan", "Version", [], 0.1)                    # built without error, but empty
    db.close()
    with pytest.raises(svhttp.SvHttpError, match="atlas is incomplete or stale for FinPlan.*Version"):
        sv.retrieve("FinPlan", {}, [Axis("Account", ["Revenue"])], [Axis("Period", ["Sep"])], True, False)


def test_an_old_atlas_still_serves_the_http_route(db_path):
    db = atlas.connect(db_path)
    _age(db, "FinMax", 40)                                  # dimension names do not go stale
    db.close()
    _q(_sv(FakeTransport(answer(GRID)), db_path))


def test_the_atlas_connection_is_closed(db_path, monkeypatch):
    opened, real = [], atlas.connect
    monkeypatch.setattr(atlas, "connect", lambda p: opened.append(real(p)) or opened[-1])
    _q(_sv(FakeTransport(answer(GRID)), db_path))
    with pytest.raises(svhttp.SvHttpError):
        _q(_sv(FakeTransport(), db_path), "FinLite")
    assert len(opened) == 2
    for db in opened:
        with pytest.raises(sqlite3.ProgrammingError):
            db.execute("SELECT 1")


def test_a_busy_route_is_an_svhttp_error(db_path, monkeypatch):
    monkeypatch.setattr(config, "SVHTTP_TIMEOUT_S", 0.01)
    sv = _sv(FakeTransport(answer(GRID)), db_path)
    sv._lock.acquire()
    try:
        with pytest.raises(svhttp.SvHttpError, match="busy with another query"):
            _q(sv)
    finally:
        sv._lock.release()
    _q(sv)                                                     # free again


def test_shared_is_one_instance_across_threads(monkeypatch):
    monkeypatch.setattr(svhttp, "_shared", None)
    got = []
    ts = [threading.Thread(target=lambda: got.append(svhttp.shared())) for _ in range(4)]
    for th in ts:
        th.start()
    for th in ts:
        th.join()
    assert len({id(x) for x in got}) == 1

def _enum_xml(*months):
    body = "|".join(f"[CurrRepMonth].[Y{d.year % 100:02d}M{d.month:02d}]" for d in months)
    return f'<res_EnumMembers><overflow>0</overflow><mbrs enc="1">{body}</mbrs></res_EnumMembers>'


def _ago(n):
    import datetime as dt
    t = dt.date.today()
    k = t.year * 12 + t.month - 1 - n
    return dt.date(k // 12, k % 12 + 1, 1)


def _label(d):
    return f"{svhttp._MONTHS[d.month - 1]} {d.year % 100:02d}"


def test_closed_month_over_http(db_path):
    t = FakeTransport(_enum_xml(_ago(1)))
    assert _sv(t, db_path).closed_month() == _label(_ago(1))
    assert t.tags() == ["GetProvisionedDataSources", "ConnectToProvider", "OpenApplication", "OpenCube", "EnumMembers"]
    assert "<dim>TimePeriods</dim>" in t.bodies[-1] and ">CurrRepMonth</arg>" in t.bodies[-1]


def test_closed_month_always_asks_the_finmax_cube(db_path):
    t = FakeTransport(_enum_xml(_ago(1)))
    assert _sv(t, db_path).closed_month("FinLite") == _label(_ago(1))
    assert "<cube>FinMax</cube>" in "".join(t.bodies) or "FinMax" in "".join(t.bodies)
    assert "FinLite" not in "".join(t.bodies)


def test_closed_month_reconnects_once_but_not_for_a_bad_name(db_path):
    t = FakeTransport(EXPIRED, _enum_xml(_ago(2)))
    assert _sv(t, db_path).closed_month() == _label(_ago(2)) and t.tags().count("ConnectToProvider") == 2
    bad = '<exception errcode="4996" type="error"><desc>member invalid</desc></exception>'
    t2 = FakeTransport(bad, _enum_xml(_ago(1)))
    with pytest.raises(svhttp.SvHttpError, match="member invalid"):
        _sv(t2, db_path).closed_month()
    assert t2.tags().count("ConnectToProvider") == 1
    with pytest.raises(svhttp.SvHttpError, match="not understood"):
        _sv(FakeTransport(_enum_xml(_ago(1)).replace("[CurrRepMonth].[Y", "x[")), db_path).closed_month()
    with pytest.raises(svhttp.SvHttpError, match="FinPlan has no Current Reporting Month"):
        _sv(FakeTransport(), db_path).closed_month("FinPlan")


def test_closed_month_wants_exactly_one_child(db_path):
    with pytest.raises(svhttp.SvHttpError, match="exactly one"):
        _sv(FakeTransport(_enum_xml(_ago(1), _ago(2))), db_path).closed_month()
    with pytest.raises(svhttp.SvHttpError, match="exactly one"):
        _sv(FakeTransport('<res_EnumMembers><mbrs enc="1"></mbrs></res_EnumMembers>'), db_path).closed_month()


def test_closed_month_must_be_plausible(db_path):
    with pytest.raises(svhttp.SvHttpError, match="plausible|future"):
        _sv(FakeTransport(_enum_xml(_ago(-1))), db_path).closed_month()          # next month: in the future
    with pytest.raises(svhttp.SvHttpError, match="plausible|old"):
        _sv(FakeTransport(_enum_xml(_ago(4))), db_path).closed_month()           # more than 3 months back
    assert _sv(FakeTransport(_enum_xml(_ago(3))), db_path).closed_month() == _label(_ago(3))
    assert _sv(FakeTransport(_enum_xml(_ago(0))), db_path).closed_month() == _label(_ago(0))

def test_fetch_returns_the_raw_grid_a_source_sheet_is_written_from(db_path):
    f = _sv(FakeTransport(answer(GRID)), db_path).fetch(
        "FinMax", {"Entities": "un.SY.2WHS", "All Scenarios": "Actual"},
        [Axis("All Accounts", ["ac.C30", "ac.EBITDA"])], [Axis("All Time Periods", ["Aug 26", "Sep 26"])])
    assert [list(r[:3]) for r in f.values] == GRID
    assert f.layout.pov == [("All Entities", "un.SY.2WHS"), ("All Scenarios", "Actual")] and f.seconds >= 0


FINLITE_DIMS = ["All Accounts", "All Balance Types", "All Currencies", "All Entities", "All Scenarios",
                "All Time Periods"]                 # ops.DRILL_DEFAULTS sets All Currencies, so FinLite needs it
DRILL_ANS = ('<res_GetDrillThroughReports><authRequired>false</authRequired><drillThroughURL name="Drill Through '
             'to source" type="post">./aif/FdmeeMain.jsp?module=aif.drilldown&amp;attribute=Entities.id.un.SY'
             '</drillThroughURL></res_GetDrillThroughReports>')
ONE_CELL = [[None, "Periodic", "NOK", "SY Telenor Norge AS (un)", "Actual"], [None, "Sep 26", None, None, None],
            ["C60080 Consultancy IS/IT (ac)", -12.5, None, None, None]]
CELL = {"All Accounts": "ac.C60080", "All Entities": "un.SY", "All Time Periods": "Sep 26"}


@pytest.fixture
def lite_db(db_path):
    db = atlas.connect(db_path)
    _store(db, "FinLite", FINLITE_DIMS)
    db.close()
    return db_path


def test_drill_link_over_http_returns_the_ops_contract(lite_db):
    t = FakeTransport(answer(ONE_CELL), DRILL_ANS)
    got = _sv(t, lite_db).drill_link("FinLite", CELL)
    assert got == {"cube_value": -12.5, "url": "aif/FdmeeMain.jsp?module=aif.drilldown&attribute=Entities.id.un.SY",
                   "report": "Drill Through to source"}
    body = t.bodies[-1]
    assert "<ords>11</ords>" in body                    # (1 + nc) * width + nr = 2 * 5 + 1
    assert '<dim name="All Time Periods" col="0" enc="0">Sep 26</dim>' in body
    assert '<dim name="All Accounts" row="0" enc="0">C60080 Consultancy IS/IT (ac)</dim>' in body
    assert '<dim name="All Entities" pov="pov" enc="0">SY Telenor Norge AS (un)</dim>' in body
    assert "<cube>FinLite</cube>" in body


def test_drill_link_empty_cell_has_no_cube_value(lite_db):
    empty = [r[:] for r in ONE_CELL]
    empty[2][1] = None
    got = _sv(FakeTransport(answer(empty), DRILL_ANS), lite_db).drill_link("FinLite", CELL)
    assert got["cube_value"] is None and got["url"]


def test_drill_link_without_report(lite_db):
    t = FakeTransport(answer(ONE_CELL), "<res_GetDrillThroughReports></res_GetDrillThroughReports>")
    assert _sv(t, lite_db).drill_link("FinLite", CELL) == {"cube_value": -12.5, "url": None, "report": None}


def test_drill_link_refuses_accumulated(lite_db):
    with pytest.raises(InputError, match="Periodic"):
        _sv(FakeTransport(), lite_db).drill_link("FinLite", {**CELL, "All Balance Types": "Accumulated"})


def test_drill_link_reconnects_once_on_an_expired_session(lite_db):
    t = FakeTransport(answer(ONE_CELL), EXPIRED, DRILL_ANS)
    assert _sv(t, lite_db).drill_link("FinLite", CELL)["url"]
    assert t.tags().count("ConnectToProvider") == 2


def test_a_drill_link_error_puts_no_session_id_token_or_cookie_into_the_reply(lite_db, monkeypatch):
    """drillstep._link quotes the route's error in the tool reply: an answer that is not the provider's, carrying a
    session id, a token and a cookie, twice (the one reconnect), reaches it in words of our own only."""
    import asyncio

    from smartass.epm import drillstep
    secrets = ("sid-7f3a9c",
               "eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJ0b20ifQ.c2lnbmF0dXJlX29mX3Rva2Vu",  # allow-secret (an invented token)
               "JSESSIONID=c00k1e9")
    bad = f"<exception><sID>{secrets[0]}</sID><token>{secrets[1]}</token></exception> Set-Cookie: {secrets[2]}"
    sv = _sv(FakeTransport(answer(ONE_CELL), bad, bad), lite_db)
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")
    monkeypatch.setattr(drillstep, "_http_link", sv.drill_link)
    with pytest.raises(drillstep.ToolError) as e:
        asyncio.run(drillstep._link("FinLite", CELL))
    said = str(e.value)
    assert "not a Smart View provider answer (other XML," in said and not any(s in said for s in secrets)


PARTS = [{"All Accounts": "ac.C60080", "All Entities": e, "All Time Periods": "Sep 26"}
         for e in ("un.SY.2TEC", "un.SY.2BUS")]


def test_child_values_one_column(lite_db):
    grid = [[None, "C60080 Consultancy IS/IT (ac)", "Periodic", "NOK", "Actual"], [None, "Sep 26", None, None, None],
            ["2TEC Technology (un.SY)", -3.0, None, None, None], ["2BUS Business (un.SY)", None, None, None, None]]
    t = FakeTransport(answer(grid))
    assert _sv(t, lite_db).child_values("FinLite", PARTS) == [-3.0, None]
    assert t.tags().count("ExecuteGrid") == 1


def test_child_values_over_rest_ask_the_same_one_column_grid_of_the_second_engine(lite_db, monkeypatch):
    sv = _sv(FakeTransport(), lite_db)
    asked = []

    def fetch_rest(conn, pov, rows, cols):
        asked.append((conn, rows[0].dimension, rows[0].members, cols[0].members))
        lay = svhttp.layout(conn, FINLITE_DIMS, pov, rows, cols)
        return svhttp.Fetched(lay, [[None] * 5, [None] * 5, ["x", -3.0], ["y", None]], 0.0, "rest")
    monkeypatch.setattr(sv, "fetch_rest", fetch_rest)
    monkeypatch.setattr(sv, "fetch", lambda *a: pytest.fail("the provider route"))
    assert sv.child_values("FinLite", PARTS, rest=True) == [-3.0, None]
    assert asked == [("FinLite", "All Entities", ["un.SY.2TEC", "un.SY.2BUS"], ["Sep 26"])]


def test_child_values_none_when_parts_do_not_differ(lite_db):
    assert _sv(FakeTransport(), lite_db).child_values("FinLite", [PARTS[0], dict(PARTS[0])]) is None


# ------------------------------------------------------------------ requests run at once (step 1b)
def test_queries_run_at_once_each_on_its_own_session(db_path):
    import threading
    gate = threading.Barrier(2, timeout=5)

    class Both(FakeTransport):
        def post(self, base, body):
            if body.startswith("<req_ExecuteGrid>") and self.gated:
                gate.wait()                                # both grids must be in flight together
            return super().post(base, body)
    t = Both(answer(GRID), answer(GRID))
    t.gated = True
    sv = _sv(t, db_path)
    out, errs = [], []

    def run():
        try:
            out.append(_q(sv))
        except Exception as e:   # noqa: BLE001
            errs.append(e)
    th = [threading.Thread(target=run) for _ in range(2)]
    for x in th:
        x.start()
    for x in th:
        x.join(10)
    assert errs == [] and len(out) == 2
    assert t.tags().count("ConnectToProvider") == 2        # one session per request in flight: never shared
    assert len(set(t.grid_sids())) == 2
    t.answers.append(answer(GRID))
    t.gated = False
    _q(sv)                                                 # afterwards a pooled session is reused
    assert t.tags().count("ConnectToProvider") == 2 and len(sv._idle["FinMax"]) == 2


def test_a_session_whose_request_was_lost_is_not_reused(db_path):
    class Lost(FakeTransport):
        def post(self, base, body):
            if body.startswith("<req_ExecuteGrid>") and not self.answers:
                raise svhttp.SvHttpError("lost the request")
            return super().post(base, body)
    t = Lost()
    sv = _sv(t, db_path)
    with pytest.raises(svhttp.SvHttpError, match="lost"):
        _q(sv)
    assert sv._idle.get("FinMax", []) == []                # EPM may still work on that sID: dropped


def test_requests_per_cube_are_capped(db_path, monkeypatch):
    monkeypatch.setattr(config, "SVHTTP_SESSIONS", 1)
    monkeypatch.setattr(config, "SVHTTP_TIMEOUT_S", 0.01)
    sv = _sv(FakeTransport(answer(GRID)), db_path)
    sv._slot("FinMax").acquire()                           # the one slot is taken
    with pytest.raises(svhttp.SvHttpError, match="busy with other queries"):
        _q(sv)


def test_a_lost_request_goes_again_on_a_new_session_and_the_old_one_is_dropped(db_path):
    class LostOnce(FakeTransport):
        lost = True

        def post(self, base, body):
            if body.startswith("<req_ExecuteGrid>") and self.lost:
                self.lost = False
                self.sent.append((base, "ExecuteGrid"))
                self.bodies.append(body)
                raise svhttp.RequestLost("lost the request")
            return super().post(base, body)
    t = LostOnce(answer(GRID))
    sv = _sv(t, db_path)
    _q(sv)
    sids = t.grid_sids()
    assert len(sids) == 2 and sids[0] != sids[1]           # the second try had a session of its own
    assert sv._idle["FinMax"] == [sids[1]]                 # the lost one is not reused



def test_the_errors_the_second_engine_would_meet_as_well_are_no_fallback(db_path, tmp_path):
    """A stale atlas, an alias clash, a pod without a scope and a token refused twice are svhttp.NoFallback: the
    provider's error is the answer, REST is not tried."""
    clash = ('<exception errcode="1000" type="error"><desc>Found duplicate aliases or member name conflicting with '
             'alias under parent x.</desc></exception>')
    t = FakeTransport(clash, answer(GRID))
    with pytest.raises(svhttp.GridRefused) as e:
        _q(_sv(t, db_path))
    assert isinstance(e.value, svhttp.NoFallback) and t.tags().count("ExecuteGrid") == 1      # no reconnect
    with pytest.raises(svhttp.AtlasStale):
        _q(_sv(FakeTransport(), tmp_path / "empty.db"))
    with pytest.raises(svhttp.AtlasStale):
        _sv(FakeTransport(), db_path).retrieve("FinMax", {"All Bogus": "x"}, [Axis("All Accounts", ["a"])],
                                              [Axis("All Time Periods", ["p"])], True, False)
    assert issubclass(svhttp.NoScope, svhttp.NoFallback) and issubclass(svhttp.TokenRefused, svhttp.NoFallback)
    for t in (FakeTransport(EXPIRED, EXPIRED), FakeTransport(answer(SHIFTED))):     # what a new route may fix
        with pytest.raises(svhttp.SvHttpError) as e:
            _q(_sv(t, db_path))
        assert not isinstance(e.value, svhttp.NoFallback)


# live 2026-11-22, a member the cube no longer has (REST: 'The member ... does not exist in the Entity dimension')
MISMATCH = ('<exception errcode="1000" type="error"><desc>Member name and Alias placed on row 3 do not match.'
            '</desc></exception>')
OUT_OF_PLACE = '<exception errcode="1000" type="error"><desc>Member x is out of place.</desc></exception>'


@pytest.mark.parametrize("body", [INVALID, INVALID.replace("<desc>", "<desc>Error: "), MISMATCH, OUT_OF_PLACE])
def test_a_wrong_member_is_final_on_the_provider_route_and_goes_on_to_the_second_engine(db_path, body):
    """Deterministic (no reconnect), but not NoFallback: Planning REST's HTTP 400 names the member, so the second
    engine can leave it out (data/provider.py)."""
    t = FakeTransport(body, answer(GRID))
    with pytest.raises(svhttp.WrongMember) as e:
        _q(_sv(t, db_path))
    assert not isinstance(e.value, svhttp.NoFallback) and t.tags().count("ExecuteGrid") == 1


@pytest.mark.parametrize("desc", ["session expired: the grid is out of place in the queue",
                                  "Cannot connect: Member name and Alias placed on row 3 do not match",
                                  "Grid is Invalidated by a restart, try again", "Member x is out of place. Retry"])
def test_a_transient_error_that_only_contains_a_wrong_member_phrase_is_not_taken_for_one(db_path, desc):
    t = FakeTransport(f'<exception errcode="1000" type="error"><desc>{desc}</desc></exception>', answer(GRID))
    _q(_sv(t, db_path))                                                 # reconnected once, then answered
    assert t.tags().count("ExecuteGrid") == 2


def test_no_token_and_a_busy_data_source_lookup_are_not_sent_on_to_the_second_engine(db_path, monkeypatch):
    """REST meets the same token and waits on the same lock: the wait would only double."""
    sv = _sv(FakeTransport(), db_path)
    monkeypatch.setattr(svhttp.config, "SVHTTP_TIMEOUT_S", 0.01)
    sv._bases_lock.acquire()
    try:
        with pytest.raises(svhttp.NoFallback, match="busy looking up"):
            _q(sv)
    finally:
        sv._bases_lock.release()
    t = svhttp.BearerTransport(tokens=_Tokens(oauth.OAuthError("refresh token expired")), opener=_Wire(lambda *a: 1))
    with pytest.raises(svhttp.NoFallback, match="no EPM access token"):
        t.post("https://plan1.example", svxml.connect())


def test_no_access_token_is_its_own_kind_of_no_fallback_so_its_reply_fits_its_cause():
    """Carry-in (Task 3c review, Minor): the drill link's next step by cause - a missing token is not 'try again'."""
    t = svhttp.BearerTransport(tokens=_Tokens(oauth.OAuthError("refresh token expired")), opener=_Wire(lambda *a: 1))
    with pytest.raises(svhttp.NoToken, match="no EPM access token .*refresh token expired"):
        t.post("https://plan1.example", svxml.connect())
    assert issubclass(svhttp.NoToken, svhttp.NoFallback)


def test_the_second_engine_sends_through_the_grid_routes_session_of_the_cubes_pod(db_path):
    t = svhttp.BearerTransport(tokens=_Tokens("at-1"), opener=_Wire(lambda *a: (200, SOURCES)))
    sv = _sv(t, db_path)
    assert sv.rest_session("FinPlan") is t.session("https://strat.example")
    assert sv.rest_session("FinLite") is sv.rest_session("FinMax") is t.session("https://plan1.example")
    with pytest.raises(svhttp.SvHttpError, match="no HTTP route"):
        sv.rest_session("Vision")
    with pytest.raises(svhttp.SvHttpError, match="the bearer route's sessions"):
        _sv(FakeTransport(), db_path).rest_session("FinMax")
    empty = svhttp.BearerTransport(tokens=_Tokens("at-1"), opener=_Wire(lambda *a: (200, NO_SOURCES)))
    with pytest.raises(svhttp.SvHttpError, match="no Smart View data source"):
        _sv(empty, db_path).rest_session("FinMax")


def test_an_alias_clash_does_not_spend_a_reconnect(db_path):
    """FinPlan's 'Duplicate name ... qualified name' is deterministic: cards_fetch.grid splits the chunk itself."""
    clash = ('<exception errcode="1000" type="error"><desc>Duplicate name oWUKONORGE WUKONORGE Elim found in grid; '
             'cannot resolve. Please replace it with qualified name.</desc></exception>')
    t = FakeTransport(clash, answer(GRID))
    with pytest.raises(svhttp.SvHttpError, match="qualified name"):
        _q(_sv(t, db_path))
    assert t.tags().count("ConnectToProvider") == 1 and t.tags().count("ExecuteGrid") == 1


# ---------------------------------------------------------------- BearerTransport (OAuth access token, no browser)
class _Tokens:
    """epm/oauth.py as the transport sees it: scope_for and access_token (which may raise)."""

    def __init__(self, *tokens, scopes=("https://strat.example", "https://plan1.example"), configured=True):
        self.tokens, self.scopes, self.asked, self.is_set_up = list(tokens), scopes, [], configured

    def configured(self):
        return self.is_set_up

    def scope_for(self, base):                              # as oauth.scope_for: by the origin, port and path ignored
        return "urn:scope" if oauth.origin(base) in self.scopes else None

    def access_token(self, base, reject=None):
        self.asked.append((base, reject))
        t = self.tokens.pop(0) if len(self.tokens) > 1 else self.tokens[0]
        if isinstance(t, Exception):
            raise t
        return t


class _Answer(io.BytesIO):
    """A 2xx answer as urllib hands it back."""

    def __init__(self, status, text, headers):
        super().__init__(text.encode("utf-8"))
        self.status, self.headers = status, headers

    def info(self):
        return self.headers


class _Wire:
    """urllib's opener under the transport's EpmSessions: records every request it is given and its timeout;
    answer(origin, body, token) - the origin and the bearer as they were sent - gives (status, text) or (status,
    text, headers), or raises. A status outside 2xx is raised as urllib.error.HTTPError, as urllib's opener does."""

    def __init__(self, answer):
        self.answer, self.requests, self.timeouts = answer, [], []

    def open(self, req, timeout):
        self.requests.append(req)
        self.timeouts.append(timeout)
        u = urlsplit(req.full_url)
        token = (req.get_header("Authorization") or "").removeprefix("Bearer ")
        status, text, *more = self.answer(f"{u.scheme}://{u.hostname}", req.data.decode("utf-8"), token)
        headers = http.client.HTTPMessage()
        for k, v in (more[0] if more else {}).items():
            headers[k] = v
        if 200 <= status < 300:
            return _Answer(status, text, headers)
        raise urllib.error.HTTPError(req.full_url, status, "no", headers, io.BytesIO(text.encode("utf-8")))


def _transport(tokens, answer=lambda *a: pytest.fail("sent")):
    """A BearerTransport whose pod sessions send over a _Wire (by default one that must not be reached)."""
    return svhttp.BearerTransport(tokens=tokens, opener=_Wire(answer))


def _bearer(tokens, *answers):
    """(transport, what was sent: (origin, token)): answers are (status, text) or an exception, in order."""
    sent, queue = [], list(answers)

    def answer(origin, body, token):
        sent.append((origin, token))
        a = queue.pop(0) if len(queue) > 1 else queue[0]
        if isinstance(a, Exception):
            raise a
        return a
    return _transport(tokens, answer), sent


OK_XML = (200, "<res_ConnectToProvider><sID>s1</sID></res_ConnectToProvider>")
GATE_401 = (401, "<html><h1>401 Authorization Required</h1></html>")


def test_bearer_posts_with_the_token():
    t, sent = _bearer(_Tokens("at-1"), OK_XML)
    assert "<sID>s1</sID>" in t.post("https://strat.example/", svxml.connect())
    assert sent == [("https://strat.example", "at-1")]


def test_bearer_refuses_writes_before_it_asks_for_a_token():
    tokens = _Tokens("at-1")
    t, sent = _bearer(tokens, OK_XML)
    with pytest.raises(ValueError, match="refused"):
        t.post("https://strat.example", "<req_SubmitData></req_SubmitData>")
    assert tokens.asked == [] and sent == []


def test_the_provider_request_goes_through_the_pods_epm_session():
    """One boundary: the transport builds no request of its own. Its POST is the session's - the pod's origin and
    the allowed path, the bearer, text/xml and the session's own framing (Accept-Encoding: identity) - on one
    EpmSession per pod (one cookie jar), which is not the drill's (epmhttp.session: its key belongs to its jar)."""
    wire = _Wire(lambda *a: OK_XML)
    t = svhttp.BearerTransport(tokens=_Tokens("at-1"), opener=wire)
    for base in ("https://STRAT.example:443/", "https://strat.example", "https://plan1.example"):
        t.post(base, svxml.connect())
    s = t.session("https://strat.example")
    assert isinstance(s, epmhttp.EpmSession) and s.origin == "https://strat.example"
    assert t.session("https://strat.example") is s and t.session("https://plan1.example") is not s
    assert s is not epmhttp.session("https://strat.example")
    assert [(r.get_method(), r.full_url) for r in wire.requests] == [
        ("POST", "https://strat.example/HyperionPlanning/SmartView")] * 2 + [
        ("POST", "https://plan1.example/HyperionPlanning/SmartView")]
    r = wire.requests[0]
    assert r.get_header("Authorization") == "Bearer at-1" and r.get_header("Content-type") == "text/xml"
    assert r.get_header("Accept-encoding") == "identity" and r.data == svxml.connect().encode()


def test_the_provider_route_is_held_to_the_sessions_allowlist(monkeypatch):
    monkeypatch.delitem(epmhttp.ALLOWED, epmhttp.SMART_VIEW)
    tokens = _Tokens("at-1")
    t, sent = _bearer(tokens, OK_XML)
    with pytest.raises(svhttp.SvHttpError, match="not on the allowlist") as e:
        t.post("https://strat.example", svxml.connect())
    assert not isinstance(e.value, svhttp.RequestLost) and sent == [] and tokens.asked == []


def test_a_write_never_leaves_even_past_the_transports_own_check(monkeypatch):
    """The session checks the body itself (svxml.check_read_only): without the transport's check a write is still
    refused before a token is asked for, and the refusal does not quote the body."""
    monkeypatch.setattr(svhttp, "svxml", types.SimpleNamespace(check_read_only=lambda body: None))
    tokens = _Tokens("at-1")
    t, sent = _bearer(tokens, OK_XML)
    with pytest.raises(svhttp.SvHttpError, match="not a single read-only Smart View request") as e:
        t.post("https://strat.example", "<req_SubmitData><sID>SID-SECRET</sID></req_SubmitData>")
    assert sent == [] and tokens.asked == [] and "SID-SECRET" not in str(e.value)


def test_a_body_over_the_sessions_cap_never_leaves(monkeypatch):
    monkeypatch.setattr(epmhttp, "MAX_SMART_VIEW_BODY_BYTES", 100)
    tokens = _Tokens("at-1")
    t, sent = _bearer(tokens, OK_XML)
    with pytest.raises(svhttp.SvHttpError, match="body over 100 bytes") as e:
        t.post("https://strat.example", svxml.connect())
    assert not isinstance(e.value, svhttp.RequestLost) and sent == [] and tokens.asked == []


@pytest.mark.parametrize("refused", [GATE_401, (302, ""), (403, "<html>403</html>")])
def test_a_refused_token_gets_one_new_one(refused):
    tokens = _Tokens("at-1", "at-2")
    t, sent = _bearer(tokens, refused, OK_XML)
    assert "<sID>s1</sID>" in t.post("https://strat.example", svxml.connect())
    assert sent == [("https://strat.example", "at-1"), ("https://strat.example", "at-2")]
    assert tokens.asked == [("https://strat.example", None), ("https://strat.example", "at-1")]


@pytest.mark.parametrize("status", [401, 302, 403])
def test_a_gate_status_that_carries_provider_xml_is_not_sent_again(status):
    """Provider XML means the request reached the provider: it is the answer whatever the status, and it is never
    sent again on the same sID (the session's one new token is for a gate that let nothing through)."""
    tokens = _Tokens("at-1", "at-2")
    t, sent = _bearer(tokens, (status, OK_XML[1]))
    assert "<sID>s1</sID>" in t.post("https://strat.example", svxml.connect())
    t2, sent2 = _bearer(_Tokens("at-1", "at-2"), (status, '<exception errcode="1"><desc>x</desc></exception>'))
    assert "<exception" in t2.post("https://strat.example", svxml.connect())
    assert sent == [("https://strat.example", "at-1")] and tokens.asked == [("https://strat.example", None)]
    assert len(sent2) == 1


@pytest.mark.parametrize("refused", [GATE_401, (302, ""), (403, "<html>403</html>")])
def test_a_token_refused_twice_is_sent_twice_and_no_more(refused):
    tokens = _Tokens("at-1", "at-2", "at-3")
    t, sent = _bearer(tokens, refused)
    with pytest.raises(svhttp.SvHttpError, match=f"refused the access token twice \\(HTTP {refused[0]}\\)") as e:
        t.post("https://strat.example", svxml.connect())
    assert not isinstance(e.value, svhttp.RequestLost)
    assert [tok for _, tok in sent] == ["at-1", "at-2"] and len(tokens.asked) == 2


def test_a_pod_keeps_its_cookies_and_a_refused_token_takes_them_along():
    """One jar per pod, as a browser has: EPM's cookie comes back on the pod's next request, never on another pod's;
    a refused token drops the cookies that came with it (the session's 401 rule) before the one more try."""
    answers = [(200, OK_XML[1], {"Set-Cookie": "JSESSIONID=one; Path=/"}), OK_XML, OK_XML, GATE_401, OK_XML]
    wire = _Wire(lambda *a: answers.pop(0))
    t = svhttp.BearerTransport(tokens=_Tokens("at-1", "at-1", "at-1", "at-1", "at-2"), opener=wire)
    for base in ("https://strat.example", "https://strat.example", "https://plan1.example", "https://strat.example"):
        t.post(base, svxml.connect())
    cookies = [r.get_header("Cookie") for r in wire.requests]
    assert cookies == [None, "JSESSIONID=one", None, "JSESSIONID=one", None]
    assert [r.get_header("Authorization") for r in wire.requests[3:]] == ["Bearer at-1", "Bearer at-2"]


@pytest.mark.parametrize("limit, want", [(120.0, 108.0 + 256.0), (10.0, 9.0 + 256.0), (1000.0, epmhttp.MAX_TIMEOUT_S)])
def test_a_request_has_the_old_first_byte_limit_plus_the_read_of_the_largest_answer(monkeypatch, limit, want):
    """The session's timeout is a deadline for the whole answer; SVHTTP_TIMEOUT_S was urllib's limit per receive. So
    the deadline is the old 90 % of it (the server's time to the first byte) plus the read of the largest answer the
    session takes (MAX_ANSWER_BYTES, 64 MB) at a slow 256 kB/s, within the most an EpmSession takes."""
    monkeypatch.setattr(config, "SVHTTP_TIMEOUT_S", limit)
    wire = _Wire(lambda *a: OK_XML)
    svhttp.BearerTransport(tokens=_Tokens("at-1"), opener=wire).post("https://strat.example", svxml.connect())
    assert wire.timeouts == [pytest.approx(want)]


def test_an_answer_over_the_size_cap_is_an_svhttp_error_not_a_lost_request(monkeypatch):
    monkeypatch.setattr(epmhttp, "MAX_ANSWER_BYTES", 10)
    t, sent = _bearer(_Tokens("at-1"), OK_XML)
    with pytest.raises(svhttp.SvHttpError, match="is over") as e:
        t.post("https://strat.example", svxml.connect())
    assert not isinstance(e.value, svhttp.RequestLost) and len(sent) == 1


def test_not_signed_in_reaches_the_caller_as_it_is():                       # Review Focus 4
    t = _transport(_Tokens(oauth.NotSignedIn("not signed in")))
    with pytest.raises(oauth.NotSignedIn):
        t.post("https://strat.example", svxml.connect())


def test_a_sign_in_that_ends_after_a_refusal_reaches_the_caller_as_it_is():
    t, sent = _bearer(_Tokens("at-1", oauth.NotSignedIn("ended")), GATE_401)
    with pytest.raises(oauth.NotSignedIn):
        t.post("https://strat.example", svxml.connect())
    assert len(sent) == 1


@pytest.mark.parametrize("tokens, send, msg", [
    (_Tokens(oauth.OAuthError("the identity domain did not answer")), lambda *a: OK_XML, "identity domain"),
    (_Tokens("at-1", "at-2"), lambda *a: GATE_401, "refused the access token"),
    (_Tokens("at-1", scopes=()), lambda *a: OK_XML, "no OAuth scope is configured"),
    (_Tokens("at-1", oauth.OAuthError("another query is still renewing")), lambda *a: GATE_401, "still renewing")])
def test_everything_else_is_an_svhttp_error(tokens, send, msg):
    with pytest.raises(svhttp.SvHttpError, match=msg) as e:
        _transport(tokens, send).post("https://strat.example", svxml.connect())
    assert not isinstance(e.value, svhttp.RequestLost)             # never retried as a lost request


def test_without_any_oauth_set_up_the_caller_gets_the_set_up_text():
    t = _transport(_Tokens("at-1", scopes=(), configured=False))
    with pytest.raises(oauth.NotSetUp) as e:
        t.post("https://plan1.example", svxml.connect())
    assert str(e.value) == oauth.NOT_SET_UP


def test_a_pod_without_a_scope_says_where_to_add_it():
    tokens = _Tokens("at-1", scopes=())
    with pytest.raises(svhttp.NoScope) as e:                      # an SvHttpError: another pod's grid falls back
        _transport(tokens).post("https://plan1.example/", svxml.connect())
    assert isinstance(e.value, svhttp.SvHttpError) and tokens.asked == []
    assert str(e.value) == "no OAuth scope is configured for https://plan1.example ([oauth.scopes] in config.toml)"


def test_the_home_pod_without_a_scope_stops_the_closed_month_with_the_add_the_scope_step(db_path):
    from smartass.core import serverkit
    sv = svhttp.SvHttp(_transport(_Tokens("at-1", scopes=("https://strat.example",))), home="https://plan1.example",
                       db_path=db_path)
    with pytest.raises(svhttp.NoScope) as e:
        sv.closed_month("FinMax")
    msg = serverkit.explain("fin_analyze", e.value)
    assert "Retry once" not in msg and msg.count("Next step:") == 1
    assert "[oauth.scopes]" in msg.split("Next step:")[1] and "smartass login" in msg and "administrator" in msg


@pytest.mark.parametrize("answer_", [TimeoutError("timed out"), ConnectionResetError("reset"),
                                     http.client.IncompleteRead(b"<res_"), (502, "<html>Bad Gateway</html>")])
def test_a_request_that_may_have_reached_epm_is_lost_not_sent_again(answer_):
    t, sent = _bearer(_Tokens("at-1"), answer_)
    with pytest.raises(svhttp.RequestLost) as e:
        t.post("https://strat.example", svxml.connect())
    assert len(sent) == 1
    assert str(e.value).startswith("Smart View provider at https://strat.example: ")
    assert "at-1" not in str(e.value)


def test_a_lost_request_is_named_by_its_kind_only():
    t, _ = _bearer(_Tokens("at-SECRET"), ConnectionResetError("reset while sending Bearer at-SECRET"))
    with pytest.raises(svhttp.RequestLost) as e:
        t.post("https://strat.example", svxml.connect())
    assert str(e.value) == "Smart View provider at https://strat.example: lost the request (ConnectionResetError)"


def test_bearer_accepts_a_provider_answer_with_any_http_status():
    t, sent = _bearer(_Tokens("at-1"), (500, EXPIRED))
    assert t.post("https://strat.example", svxml.connect()) == EXPIRED


def test_a_grid_over_the_bearer_route_is_the_same_contract(db_path):
    fake = FakeTransport(answer(GRID))
    t = _transport(_Tokens("at-1"), lambda origin, body, token: (200, fake.post(origin, body)))
    lay, parsed, _ = _q(_sv(t, db_path))
    assert [r[1:] for r in parsed.rows] == [[100.0, 110.0], [50.0, None]]
    assert fake.tags() == ["GetProvisionedDataSources", "ConnectToProvider", "OpenApplication", "OpenCube", "ExecuteGrid"]


def test_the_shared_route_is_the_bearer_alone(monkeypatch):
    monkeypatch.setattr(svhttp, "_shared", None)
    sv = svhttp.shared()
    assert isinstance(sv.t, svhttp.BearerTransport) and sv.t._tokens is oauth and not hasattr(sv.t, "fallback")
    assert not hasattr(config, "SVHTTP_BEARER_RETRY_S")


def test_not_signed_in_is_not_swallowed_into_the_excel_fallback(db_path):
    """run_prepare with a provider whose transport raises NotSignedIn raises NotSignedIn (no 'routes: excel')."""
    from smartass import pipeline
    from smartass.data.provider import SvHttpProvider
    t = _transport(_Tokens(oauth.NotSignedIn("x")), lambda *a: OK_XML)
    with pytest.raises(oauth.NotSignedIn):
        pipeline.run_prepare({}, SvHttpProvider(_sv(t, db_path)))


# ---------------------------------------------------------------- M4: the bearer only over https to the pod's own origin
@pytest.mark.parametrize("base", [
    "http://strat.example", "HTTP://strat.example:443", "https://strat.example:8443", "https://strat.example:80",
    "https://user:pw-SECRET@strat.example", "https://strat.example@evil.example", "https://evil.example",
    "https://strat.example.evil.example", "https://strat.example/x", "https://strat.example?sso_token=SECRET",
    "https://strat.example/?x=SECRET", "https://strat.example#SECRET", "ftp://strat.example", "//strat.example",
    "strat.example", "https://strat.example:99999", "https://strat.example\t", "https://strat.example\n/x"])
def test_the_bearer_goes_only_over_https_to_the_pods_own_host_and_port(base, caplog):
    tokens = _Tokens("at-1")
    t, sent = _bearer(tokens, OK_XML)
    with caplog.at_level("DEBUG"), pytest.raises(svhttp.SvHttpError) as e:
        t.post(base, svxml.connect())
    assert tokens.asked == [] and sent == []                        # refused before a token was asked for
    assert not isinstance(e.value, svhttp.RequestLost)
    assert "SECRET" not in str(e.value) and "pw-" not in str(e.value) and "SECRET" not in caplog.text


@pytest.mark.parametrize("base", ["https://strat.example", "https://strat.example/", "https://STRAT.example:443",
                                  "HTTPS://strat.example:443/"])
def test_the_pods_own_origin_is_where_the_bearer_goes(base):
    t, sent = _bearer(_Tokens("at-1"), OK_XML)
    t.post(base, svxml.connect())
    assert sent == [("https://strat.example", "at-1")]
