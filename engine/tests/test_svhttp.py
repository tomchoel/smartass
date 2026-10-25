import concurrent.futures as cf
import re
import sqlite3
import threading

import pytest
from svxml_util import answer

from smartview_mcp import atlas, config, drill_browser, svhttp, svxml
from smartview_mcp.grid import Axis
from smartview_mcp.host import InputError


def _store(db, conn, dims, error=None):
    for d in dims:
        atlas.store(db, conn, d, [{"name": d, "label": d, "parent": None, "depth": 0, "order": 0}], 0.1, error)


def test_atlas_dimensions_sorted_per_connection(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    _store(db, "FinMax", ["All Time Periods", "All Accounts"])
    _store(db, "FinPlan", ["Account"])
    assert atlas.dimensions(db, "FinMax") == ["All Accounts", "All Time Periods"]
    assert atlas.dimensions(db, "FinLite") == []


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


SOURCES = ('<res_GetProvisionedDataSources><Product id="HP"><Server name="FinMax" '
           'context="https://plan1.example/HyperionPlanning/SmartView" appName="FinMax"/></Product>'
           '<Product id="HP"><Server name="FinPlan_x" context="https://strat.example/HyperionPlanning/SmartView" '
           'appName="FinPlan"/></Product></res_GetProvisionedDataSources>')
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


def test_session_is_reused_and_finplan_goes_to_its_pod(db_path):
    t = FakeTransport(answer(GRID), answer(GRID), answer([[None, "x"], [None, "Sep"], ["Revenue", 1.0]]))
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
    assert t.grid_sids() == ["cube1", "grid&amp;2"] and sv._sids["FinMax"] == "grid&2"


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
    with pytest.raises(svhttp.SvHttpError, match="no Smart View data source for app FinMax"):
        sv.base_for("FinMax")
    assert sv.base_for("FinMax") == "https://plan1.example"                  # the empty list was not kept
    assert sv.base_for("FinPlan") == "https://strat.example"                 # not listed: fetched once more
    assert t.tags().count("GetProvisionedDataSources") == 3
    t2 = FakeTransport(sources=(ONLY_FINMAX,))
    with pytest.raises(svhttp.SvHttpError, match="app FinPlan"):
        _sv(t2, db_path).base_for("FinPlan")
    assert t2.tags().count("GetProvisionedDataSources") == 1


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
    _store(db, "FinMax", ["All Scenarios"], error="timed out")
    db.close()
    with pytest.raises(svhttp.SvHttpError, match="atlas is incomplete or stale for FinMax.*All Scenarios"):
        _q(sv)


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


# ------------------------------------------------------------------ BrowserTransport with fake pages
class _Page:
    def __init__(self, replies=()):
        self.replies, self.bodies, self.closed = list(replies), [], False

    def is_closed(self):
        return self.closed

    def close(self):
        self.closed = True

    def evaluate(self, js, body):
        self.bodies.append(body)
        r = self.replies.pop(0)
        if isinstance(r, BaseException):
            raise r
        return r


class _Ctx:
    def __init__(self, *pages):
        self.pages = list(pages)

    def new_page(self):
        return self.pages.pop(0)


class _Browser:
    base = "https://plan1.example"

    def __init__(self, page, ctx=None):
        self.ctx = ctx or object()
        self.s = type("S", (), {"page": page, "ctx": self.ctx, "ready": True, "deadline": 0})()
        self.logins, self.login_calls, self.login_errors, self.ensure_errors = 0, [], [], []

    def submit(self, fn):
        f = cf.Future()
        try:
            f.set_result(fn(self.s))
        except BaseException as e:   # noqa: BLE001 - like DrillBrowser's worker
            f.set_exception(e)
        return f

    def _ensure(self, s):
        if self.ensure_errors:
            raise self.ensure_errors.pop(0)
        if not s.ready:
            self.logins += 1
            s.ready = True

    def _login_once(self, page, deadline, base=None):
        self.login_calls.append((page, base))
        if self.login_errors:
            raise self.login_errors.pop(0)


OK = [200, "<res_OpenCube></res_OpenCube>"]
STRAT = "https://strat.example"


def _bt(b):
    return svhttp.BrowserTransport(lambda: b)


def test_browser_transport_refuses_writes_and_signs_in_again_once():
    page = _Page([[401, "<html>login</html>"], OK])
    b = _Browser(page)
    t = _bt(b)
    with pytest.raises(ValueError, match="refused"):
        t.post(b.base, "<req_WriteBack/>")
    assert page.bodies == []
    assert t.post(b.base, svxml.open_cube("s", "FinMax", "FinMax")) == OK[1]
    assert b.logins == 1


def test_browser_transport_gives_up_after_two_bad_answers():
    b = _Browser(_Page([[401, "x"], [401, "x"]]))
    with pytest.raises(svhttp.SvHttpError, match="HTTP 401 twice"):
        _bt(b).post(b.base, svxml.connect())


def test_a_provider_answer_is_accepted_with_any_http_status():
    exc = [500, '<exception errcode="1000"><desc>bad</desc></exception>']
    b = _Browser(_Page([exc]))
    assert _bt(b).post(b.base, svxml.connect()) == exc[1] and b.logins == 0


class _Fut(_Browser):
    def __init__(self, make):
        super().__init__(_Page())
        self.make = make

    def submit(self, fn):
        return self.make()


def _failed(exc):
    def make():
        f = cf.Future()
        f.set_exception(exc)
        return f
    return make


def test_browser_failures_become_svhttp_errors(monkeypatch):
    def cancelled():
        f = cf.Future()
        f.cancel()
        return f

    for make, text in [(cancelled, "cancel"), (_failed(drill_browser.DrillBusy("drill browser busy")), "busy"),
                       (_failed(drill_browser.DrillError("sign-in failed")), "sign-in failed"),
                       (_failed(RuntimeError("Target closed https://x/p?sso_token=SECRET\nCall log")),
                        r"browser error: Target closed https://x/p\?\.\.\.$"),
                       (_failed(RuntimeError()), "browser error: RuntimeError")]:
        b = _Fut(make)
        with pytest.raises(svhttp.SvHttpError, match=text):
            _bt(b).post(b.base, svxml.connect())
    monkeypatch.setattr(config, "SVHTTP_TIMEOUT_S", 0.01)
    pending = cf.Future()                                  # never answers
    b = _Fut(lambda: pending)
    with pytest.raises(svhttp.SvHttpError, match="did not answer within 0 s"):
        _bt(b).post(b.base, svxml.connect())
    assert pending.cancelled()                             # a queued job does not run later for nobody


def test_an_evaluate_error_gets_one_fresh_try_then_an_svhttp_error():
    b = _Browser(_Page([RuntimeError("Execution context was destroyed"), OK]))
    assert _bt(b).post(b.base, svxml.connect()) == OK[1] and b.logins == 1
    b = _Browser(_Page([RuntimeError("x"), RuntimeError()]))
    with pytest.raises(svhttp.SvHttpError, match=f"browser error at {b.base}: RuntimeError$"):
        _bt(b).post(b.base, svxml.connect())


def test_a_raw_error_from_ensure_is_redacted():
    b = _Browser(_Page([OK]))
    b.ensure_errors = [RuntimeError("crash at https://x/y?sso_token=SECRET"), RuntimeError("again ?sso_token=S")]
    with pytest.raises(svhttp.SvHttpError) as e:
        _bt(b).post(b.base, svxml.connect())
    assert "SECRET" not in str(e.value) and "again ?..." in str(e.value)
    b = _Browser(_Page([OK]))
    b.ensure_errors = [drill_browser.DrillError("Telenor SSO sign-in failed")]   # already retried inside: no 2nd try
    with pytest.raises(svhttp.SvHttpError, match="SSO sign-in failed"):
        _bt(b).post(b.base, svxml.connect())
    assert b.s.page.bodies == []


def test_other_pod_signs_in_once_more_in_a_fresh_tab():
    bad, good = _Page(), _Page([OK])
    b = _Browser(_Page(), _Ctx(bad, good))
    b.login_errors = [drill_browser.DrillError("last page: x")]
    out = _bt(b).post(STRAT, svxml.open_cube("s", "FinPlan", "FinPlan"))
    assert out == OK[1] and bad.closed
    assert b.login_calls == [(bad, STRAT), (good, STRAT)]


def test_other_pod_sign_in_failing_twice_is_an_svhttp_error_and_closes_both_tabs():
    one, two = _Page(), _Page()
    b = _Browser(_Page(), _Ctx(one, two))
    b.login_errors = [drill_browser.DrillError("one"), drill_browser.DrillError("two")]
    with pytest.raises(svhttp.SvHttpError, match="sign-in on https://strat.example failed twice: two"):
        _bt(b).post(STRAT, svxml.connect())
    assert one.closed and two.closed


def test_other_pod_tab_is_reused_and_renewed_when_ctx_or_page_changes():
    tab1, tab2, tab3 = _Page([OK, OK, OK]), _Page([OK]), _Page([OK])
    b = _Browser(_Page(), _Ctx(tab1, tab2, tab3))
    t = _bt(b)
    t.post(STRAT, svxml.connect())
    t.post(STRAT + "/", svxml.connect())
    assert len(b.login_calls) == 1 and len(tab1.bodies) == 2
    b.s.ctx = _Ctx(tab2)                                  # the browser restarted: new context, new tab
    t.post(STRAT, svxml.connect())
    assert b.login_calls[-1] == (tab2, STRAT)
    tab2.closed = True                                    # the tab died; same context as before ...
    b.s.ctx.pages = [tab3]
    t.post(STRAT, svxml.connect())                        # ... but the page is closed: fresh tab
    assert b.login_calls[-1] == (tab3, STRAT) and len(b.login_calls) == 3


def test_other_pod_failures_drop_only_that_tab():
    tab1, tab2 = _Page([[401, "<html>login</html>"]]), _Page([OK, RuntimeError("navigated")])
    tab3 = _Page([OK])
    b = _Browser(_Page(), _Ctx(tab1, tab2, tab3))
    t = _bt(b)
    assert t.post(STRAT, svxml.connect()) == OK[1]        # 401 on tab1 -> tab1 closed, tab2 signs in ...
    assert tab1.closed and b.s.ready and b.logins == 0   # ... and the own pod's session is untouched
    assert t.post(STRAT, svxml.connect()) == OK[1]        # evaluate error on tab2 -> closed, fresh tab3
    assert tab2.closed and b.s.ready and b.logins == 0


def test_the_browser_is_resolved_on_every_post_so_a_shutdown_is_survived(monkeypatch):
    current = [_Browser(_Page([OK]))]
    monkeypatch.setattr(svhttp, "_shared", None)
    monkeypatch.setattr(drill_browser, "shared_browser", lambda: current[0])
    sv = svhttp.shared()
    assert sv is svhttp.shared()
    first = current[0]
    sv.t.post(first.base, svxml.connect())
    current[0] = _Browser(_Page([OK]))                    # shutdown_shared() + a new shared_browser()
    sv.t.post(first.base, svxml.connect())
    assert len(first.s.page.bodies) == 1 and len(current[0].s.page.bodies) == 1


def test_shared_is_one_instance_across_threads(monkeypatch):
    monkeypatch.setattr(svhttp, "_shared", None)
    monkeypatch.setattr(drill_browser, "shared_browser", lambda: _Browser(_Page()))
    got = []
    ts = [threading.Thread(target=lambda: got.append(svhttp.shared())) for _ in range(4)]
    for th in ts:
        th.start()
    for th in ts:
        th.join()
    assert len({id(x) for x in got}) == 1
