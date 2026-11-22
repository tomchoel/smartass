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
                                   svhttp.AtlasStale("atlas is stale"), svhttp.GridRefused("Grid is Invalid")])
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
                                                 "fetch it either: .*HTTP 500"):
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
