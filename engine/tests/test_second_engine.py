"""The second engine as the fallback (phase 2, Task 3): a grid request the Smart View provider route could not fetch
is fetched once over Planning REST (epm/rest.py) - not for an error REST would meet as well, never through Excel.
(The reply's routes and warning: tests/v2/test_v2_fallback_flow.py.) EPM is faked as in test_rest.py; every member,
label and figure is invented."""
import threading

import pytest
from test_rest import ATLAS, _Epm, _session

from smartass.data import provider
from smartass.data.provider import SvHttpProvider
from smartass.epm import oauth, svhttp
from smartass.knowledge import atlas
from smartass.model.requests import GridRequest

FM = GridRequest("fm", "FinMax", ("ac.C30", "ac.C70"), ("un.SY.E1",), ("Aug 26", "Sep 26"), "Actual", "", "Periodic")
FP = GridRequest("fp", "FinPlan", ("EBX Ebitda X",), ("1TEC Tech",), ("Dec 25", "Jan 26"), "Target", "Final")


@pytest.fixture
def atlas_path(tmp_path):
    d = atlas.connect(tmp_path / "atlas.db")
    for conn, dim, members in ATLAS:
        atlas.store(d, conn, dim, [{"name": n, "label": lab, "parent": None if i == 0 else dim, "depth": int(i > 0),
                                    "order": i} for i, (n, lab) in enumerate(members)], 1.0)
    d.close()
    return tmp_path / "atlas.db"


@pytest.fixture(autouse=True)
def _no_wait(monkeypatch):
    monkeypatch.setattr(provider, "RETRY_WAIT_S", 0)
    monkeypatch.setattr("smartass.epm.rest._listed", {})


def _route(atlas_path, monkeypatch, fails, epm=None):
    """SvHttpProvider over a real SvHttp whose provider grids fail with 'fails' and whose REST session is faked."""
    sv = svhttp.SvHttp(None, home="https://plan1.example", db_path=atlas_path)
    asked = []

    def retrieve(*a):
        asked.append(a[0])
        raise fails
    monkeypatch.setattr(sv, "retrieve", retrieve)
    epm = epm or _Epm()
    monkeypatch.setattr(sv, "rest_session", lambda conn: _session(epm))
    return SvHttpProvider(sv), epm, asked


def _slices(epm):
    return [u for _, u in epm.sent if u.endswith("/exportdataslice")]


def test_a_grid_the_provider_route_could_not_fetch_is_fetched_once_over_rest(atlas_path, monkeypatch):
    prov, epm, asked = _route(atlas_path, monkeypatch, svhttp.RequestLost("Smart View provider: lost the request"))
    got = prov.fetch(FM)
    assert got.route == "rest" and len(_slices(epm)) == 1
    assert asked == ["FinMax", "FinMax"]                    # the provider's own retry of a lost request, then REST
    assert "lost the request" in got.fallback
    assert {(f.account, f.entity, f.period): f.value for f in got.facts} == {      # ac.C70 is blank: no fact
        ("ac.C30", "un.SY.E1", "Aug 26"): 140.25, ("ac.C30", "un.SY.E1", "Sep 26"): 141.25}
    assert got.warnings == []


def test_a_finplan_request_is_one_rest_grid_per_year_in_the_provider_routes_order(atlas_path, monkeypatch):
    prov, epm, _ = _route(atlas_path, monkeypatch, svhttp.SvHttpError("Smart View provider: session expired"),
                          _Epm(cell=lambda row, col, pov: 1.0))
    got = prov.fetch(FP)
    years = [b["gridDefinition"]["pov"]["members"] for b in epm.bodies if b and "gridDefinition" in b]
    assert [m for ms in years for m in ms if m[0].startswith("FY")] == [["FY25"], ["FY26"]]
    assert sorted(f.period for f in got.facts) == ["Dec 25", "Jan 26"] and got.route == "rest"


@pytest.mark.parametrize("fails", [svhttp.NoScope("no scope"), svhttp.TokenRefused("refused twice"),
                                   svhttp.AtlasStale("atlas is stale"),
                                   svhttp.GridRefused("Found duplicate aliases or member name conflicting with alias"),
                                   svhttp.NoFallback("no EPM access token for https://plan1.example: expired"),
                                   svhttp.NoFallback("the HTTP route is busy looking up the EPM data sources")])
def test_an_error_rest_would_meet_as_well_is_the_answer_and_rest_is_not_asked(atlas_path, monkeypatch, fails):
    prov, epm, _ = _route(atlas_path, monkeypatch, fails)
    with pytest.raises(type(fails)):
        prov.fetch(FM)
    assert epm.sent == []


def test_no_sign_in_is_raised_and_rest_is_not_asked(atlas_path, monkeypatch):
    prov, epm, _ = _route(atlas_path, monkeypatch, oauth.NotSignedIn("sign in"))
    with pytest.raises(oauth.NotSignedIn):
        prov.fetch(FM)
    assert epm.sent == []


def test_a_grid_neither_route_can_fetch_is_the_provider_routes_error_with_rests(atlas_path, monkeypatch):
    epm = _Epm(answer=lambda path, body: (500, "{}") if path.endswith("exportdataslice") else None)
    prov, _, _ = _route(atlas_path, monkeypatch, svhttp.SvHttpError("Smart View provider: session expired"), epm)
    with pytest.raises(svhttp.SvHttpError, match="session expired.*the second engine \\(Planning REST\\) could not "
                                                 "fetch it either \\(query fm\\): .*HTTP 500"):
        prov.fetch(FM)


def test_a_period_the_atlas_does_not_have_is_filed_under_the_period_asked_with_rests_warning(atlas_path,
                                                                                               monkeypatch):
    """REST keeps its own name for a member the atlas cannot label (a warning); the facts are filed by the layout
    asked (grid.parse), so the request's own period - not REST's name - and no crash."""
    import test_rest
    monkeypatch.setitem(test_rest.REAL, "Oct 26", "Y26M10")
    prov, _, _ = _route(atlas_path, monkeypatch, svhttp.SvHttpError("Smart View provider: session expired"))
    req = GridRequest("fm", "FinMax", ("ac.C30",), ("un.SY.E1",), ("Sep 26", "Oct 26"), "Actual", "", "Periodic")
    got = prov.fetch(req)
    assert sorted(f.period for f in got.facts) == ["Oct 26", "Sep 26"]
    assert got.warnings == ["FinMax: the atlas has no label for All Time Periods member Y26M10 (kept as its name)"]


class _Closing:
    def __init__(self, db, closed):
        self.db, self.closed = db, closed

    def __getattr__(self, name):
        return getattr(self.db, name)

    def close(self):
        self.closed.append(1)
        self.db.close()


def test_rest_reads_the_atlas_on_the_thread_that_fetches_and_closes_it(atlas_path, monkeypatch):
    opened, closed = [], []
    real = atlas.connect

    def connect(path=atlas_path):
        db = real(path)
        opened.append(threading.get_ident())
        return _Closing(db, closed)
    monkeypatch.setattr(atlas, "connect", connect)
    prov, _, _ = _route(atlas_path, monkeypatch, svhttp.SvHttpError("Smart View provider: session expired"))
    out = []
    threads = [threading.Thread(target=lambda: out.append(prov.fetch(FM).route)) for _ in range(3)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert out == ["rest"] * 3 and len(set(opened)) == 3 and len(closed) == 3


# ---------------------------------------------------------------- a wrong member over REST (Task 3b)
def _rest_only(atlas_path, monkeypatch, epm):
    return _route(atlas_path, monkeypatch, svhttp.WrongMember("Smart View provider: Grid is Invalid."), epm)[0]


def _facts(got):
    return sorted((f.account, f.entity, f.period, f.value) for f in got.facts)


def test_the_providers_wrong_member_error_goes_on_to_rest(atlas_path, monkeypatch):
    prov, epm, _ = _route(atlas_path, monkeypatch, svhttp.WrongMember("Smart View provider: Grid is Invalid."))
    assert prov.fetch(FM).route == "rest" and len(_slices(epm)) == 1


def test_a_member_rest_refuses_in_the_middle_of_the_rows_is_left_out_and_the_rest_sits_on_its_own_code(
        atlas_path, monkeypatch):
    """ac.C70 between ac.C30 and ac.C90: every remaining fact must be the one a request without ac.C70 gives (facts
    are built from the members asked: a shifted place would put a number on the wrong code, silently), ac.C70 has
    none (absent, not 0), and the request says which member was left out (one entry)."""
    import test_rest
    cell = lambda row, col, pov: float(int(row[0][-2]) * 100 + int(row[1][-1]) * 10 + int(col[0][-1]))  # noqa: E731
    asked = GridRequest("fm", "FinMax", ("ac.C30", "ac.C70", "ac.C90"), ("un.SY.E1", "un.SY.E2"),
                        ("Aug 26", "Sep 26"), "Actual", "", "Periodic")
    reduced = GridRequest("fm", "FinMax", ("ac.C30", "ac.C90"), ("un.SY.E1", "un.SY.E2"), ("Aug 26", "Sep 26"),
                          "Actual", "", "Periodic")
    epm = _Epm(cell=cell, answer=test_rest.refusing("ac.C70"))
    got = _rest_only(atlas_path, monkeypatch, epm).fetch(asked)
    alone = _Epm(cell=cell)
    want = _rest_only(atlas_path, monkeypatch, alone).fetch(reduced)
    assert _facts(got) == _facts(want) and len(got.facts) == 8
    assert not any(f.account == "ac.C70" for f in got.facts)
    assert got.dropped == [("FinMax", "All Accounts", "ac.C70")] and got.route == "rest"
    assert [b for b in epm.bodies if b and "gridDefinition" in b][-1] == alone.bodies[-1]      # without ac.C70


def test_a_period_rest_refuses_in_the_middle_of_the_columns_is_left_out_the_same_way(atlas_path, monkeypatch):
    import test_rest
    cell = lambda row, col, pov: float(int(col[0][-2:]) * 10 + len(row[0]))      # noqa: E731
    asked = GridRequest("fm", "FinMax", ("ac.C30", "ac.C90"), ("un.SY.E1",), ("Aug 26", "Jul 26", "Sep 26"),
                        "Actual", "", "Periodic")
    reduced = GridRequest("fm", "FinMax", ("ac.C30", "ac.C90"), ("un.SY.E1",), ("Aug 26", "Sep 26"), "Actual", "",
                          "Periodic")
    monkeypatch.setitem(test_rest.REAL, "Jul 26", "Y26M07")
    got = _rest_only(atlas_path, monkeypatch, _Epm(cell=cell, answer=test_rest.refusing("Jul 26"))).fetch(asked)
    want = _rest_only(atlas_path, monkeypatch, _Epm(cell=cell)).fetch(reduced)
    assert _facts(got) == _facts(want) and {f.period for f in got.facts} == {"Aug 26", "Sep 26"}
    assert got.dropped == [("FinMax", "All Time Periods", "Jul 26")]


@pytest.mark.parametrize("wrong, where", [("NOK", "a POV member"), ("un.SY.E1", "the only All Entities member")])
def test_a_member_that_cannot_be_left_out_fails_the_request_with_its_code_the_grid_and_what_to_do(
        atlas_path, monkeypatch, wrong, where):
    import test_rest
    epm = _Epm(answer=test_rest.refusing(wrong))
    with pytest.raises(svhttp.SvHttpError) as e:
        _rest_only(atlas_path, monkeypatch, epm).fetch(FM)
    text = str(e.value)
    assert wrong in text and where in text and "query fm" in text and "FinMax grid" in text
    assert "smartass atlas build" in text and len(_slices(epm)) == 1                 # never a reduced request


def test_more_than_three_wrong_members_fail_the_request_with_the_list_of_those_found(atlas_path, monkeypatch):
    import test_rest
    wrong = ("ac.C10", "ac.C20", "ac.C40", "ac.C70")
    req = GridRequest("fm", "FinMax", ("ac.C30", *wrong), ("un.SY.E1",), ("Sep 26",), "Actual", "", "Periodic")
    epm = _Epm(answer=test_rest.refusing(*wrong))
    with pytest.raises(svhttp.SvHttpError, match="more than 3 members.*ac.C10, ac.C20, ac.C40, ac.C70"):
        _rest_only(atlas_path, monkeypatch, epm).fetch(req)
    assert len(_slices(epm)) == 4


def test_a_finplan_member_refused_in_one_years_grid_is_left_out_of_every_year_once(atlas_path, monkeypatch):
    """One GridRequest is one grid per year: the member is left out of both, the request names it once, and the
    remaining facts are those of the request without it."""
    import test_rest

    def answer(path, body):                     # 2BUS Biz exists in FY25 only: EPM refuses it in FY26's grid
        if path.endswith("exportdataslice") and ["FY26"] in body["gridDefinition"]["pov"]["members"]:
            return test_rest.refusing("2BUS Biz")(path, body)
    req = GridRequest("fp", "FinPlan", ("EBX Ebitda X",), ("1TEC Tech", "2BUS Biz"), ("Dec 25", "Jan 26"), "Target",
                      "Final")
    reduced = GridRequest("fp", "FinPlan", ("EBX Ebitda X",), ("1TEC Tech",), ("Dec 25", "Jan 26"), "Target", "Final")
    cell = lambda row, col, pov: float(sum(map(ord, row[1] + col[0] + "".join(pov))))      # noqa: E731
    got = _rest_only(atlas_path, monkeypatch, _Epm(cell=cell, answer=answer)).fetch(req)
    want = _rest_only(atlas_path, monkeypatch, _Epm(cell=cell)).fetch(reduced)
    assert got.dropped == [("FinPlan", "Entity", "2BUS Biz")]
    assert _facts(got) == _facts(want) and {f.entity for f in got.facts} == {"1TEC Tech"}
    assert sorted(f.period for f in got.facts) == ["Dec 25", "Jan 26"]
