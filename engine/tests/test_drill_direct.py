"""drill_direct: drill lines from EPM's own drill query, without a page per piece (figures invented)."""
import asyncio
import concurrent.futures as cf
import json

import pytest

from smartass.core import config
from smartview_mcp import drill_direct as dd
from smartview_mcp.drill_browser import CellTooWide, DrillError

URL = "aif/m/FdmeeMain.jsp?module=aif.drilldown&DRILLURL=Drill&sso_token=T&attribute=Entities.id.un.SY"


def test_sui_url_drops_drillurl_and_asks_for_json():
    u = dd.sui_url("https://epm", "./" + URL, 123)
    assert u == ("https://epm/aif/m/FdmeeMain.jsp?module=aif.drilldown&sso_token=T&attribute=Entities.id.un.SY"
                 "&FORMAT=SUI&_=123")


def test_sui_url_passes_values_through_byte_for_byte():
    u = dd.sui_url("https://epm/", "./aif/x.jsp?a=b%20c&DRILLURL=Drill&t=A+B/c=%2F&e=", 5)
    assert u == "https://epm/aif/x.jsp?a=b%20c&t=A+B/c=%2F&e=&FORMAT=SUI&_=5"


def _sui(**kw):
    return json.dumps({"DRILL_LOG": "/tmp/x", "APPLICATION_ID": 52, "PARTITIONKEY": 39, "status": 0,
                       "QUERY": "SELECT 1", "PARAMSARRAY": ["39", "un.SY"], "PARAMSTYPEARRAY": ["BigDecimal", "String"],
                       **kw})


def test_parse_sui_gives_epms_query_unchanged():
    q = dd.parse_sui(_sui())
    assert q.app_id == 52 and q.partition == 39
    assert json.loads(q.body) == {"QUERY": "SELECT 1", "PARAMS": ["39", "un.SY"],
                                  "PARAMSTYPE": ["BigDecimal", "String"]}


def test_parse_sui_no_data_is_a_note_not_an_error():
    with pytest.raises(dd.NoDrillData, match="not loaded to this intersection"):
        dd.parse_sui(json.dumps({"status": 1, "DATA": [], "details": "Drill query does not return any data. "
                                 "Data was not loaded to this intersection using Data Integration."}))


def test_parse_sui_too_wide():
    with pytest.raises(CellTooWide, match="EPMFDM-140373"):
        dd.parse_sui(json.dumps({"status": 1, "details": "EPMFDM-140373:The drill from summary operation ..."}))


def test_parse_sui_not_json_means_signed_out():
    with pytest.raises(dd.SignedOut):
        dd.parse_sui("<html>epmexit.jsp</html>")


def test_parse_sui_other_refusal():
    with pytest.raises(DrillError, match="boom"):
        dd.parse_sui(json.dumps({"status": 1, "details": "boom"}))


class FakePage:
    """Answers fetches like EPM (shapes from the spike); `pages` = getData answers in order."""

    def __init__(self, sui, pages, warm_requests=True):
        self.sui, self.pages, self.warm_requests = sui, list(pages), warm_requests
        self.handlers, self.posts, self.gets, self.closed, self.url = {}, [], [], False, ""

    def on(self, ev, fn):
        self.handlers[ev] = fn

    def remove_listener(self, ev, fn):
        self.handlers.pop(ev, None)

    def goto(self, url, **kw):
        self.url = url
        if self.warm_requests and "request" in self.handlers:
            self.handlers["request"](type("R", (), {"url": "https://epm/aif/ui/model/drilldown/getData?limit=500&offset=0",
                                                    "headers": {"page_id": "aif_drilldown",
                                                                "x-appbuilder-client-id": "cid", "cookie": "x"}})())

    def wait_for_timeout(self, ms):
        pass

    def is_closed(self):
        return self.closed

    def close(self):
        self.closed = True

    def fetch(self, url, opts):
        """What the in-tab fetch answers: (status, text)."""
        if "FORMAT=SUI" in url:
            self.gets.append((url, opts["headers"]))
            return 200, self.sui
        if "view?page=DRILL_THROUGH" in url:
            self.gets.append((url, opts["headers"]))
            return 200, json.dumps({"items": [{"viewName": "DrillFilter", "viewDefinition": "[]"}]})
        self.posts.append((url, opts["body"], opts["headers"]))
        return 200, json.dumps({"responsePayload": self.pages.pop(0)})


class FakeSession:
    def __init__(self, page):
        self.page = page                     # the plain signed-in tab
        self.made = []
        outer = self

        class Ctx:
            def new_page(self):
                outer.made.append(page)
                page.closed = False          # the fake hands the same object back as a fresh tab
                return page
        self.ctx = Ctx()


class FakeBrowser:
    """DrillBrowser's surface: jobs run at once; request() answers from the tab prep() picked (tabs[] records it)."""
    base = "https://epm"

    def __init__(self, s):
        self.s, self.logins, self.tabs = s, 0, []

    def _ensure(self, s):
        pass

    def _login(self, s):
        self.logins += 1

    def submit(self, fn):
        f = cf.Future()
        try:
            f.set_result(fn(self.s))
        except BaseException as e:   # noqa: BLE001
            f.set_exception(e)
        return f

    def request(self, prep, timeout_s):
        f = cf.Future()
        try:
            page, url, opts = prep(self.s)
            self.tabs.append(page)
            f.set_result(page.fetch(url, opts))
        except BaseException as e:   # noqa: BLE001
            f.set_exception(e)
        return f


def _items(n, amount="-1"):
    return [{"AMOUNTX": amount}] * n


def _setup(page, **kw):
    s = FakeSession(page)
    b = FakeBrowser(s)
    return dd.DirectDrill(browser=lambda: b, **kw), s, b


def test_fetch_warms_once_then_reuses_the_tab():
    page = FakePage(_sui(), [{"totalResults": 2, "hasMore": False, "items": _items(2)},
                             {"totalResults": 1, "hasMore": False, "items": _items(1)}])
    d, s, b = _setup(page)
    r1 = asyncio.run(d.fetch(URL, 100))
    r2 = asyncio.run(d.fetch(URL, 100))
    assert len(r1["items"]) == 2 and len(r2["items"]) == 1
    assert len(s.made) == 1                                   # one warm page for both pieces
    assert r1["view"]["items"][0]["viewName"] == "DrillFilter"
    assert json.loads(r1["query"].body)["QUERY"] == "SELECT 1"   # the query travels with the piece (drill_derive)
    url, body, headers = page.posts[0]
    assert f"getData?limit={config.DRILL_PAGE}&offset=0" in url and json.loads(body)["QUERY"] == "SELECT 1"
    assert headers["x-appbuilder-client-id"] == "cid" and "cookie" not in headers


def test_no_data_needs_no_warm_page():
    page = FakePage(json.dumps({"status": 1, "details": "Drill query does not return any data."}), [])
    d, s, _ = _setup(page)
    with pytest.raises(dd.NoDrillData):
        asyncio.run(d.fetch(URL, 100))
    assert s.made == [] and page.posts == []                # getData never sent


def test_warm_page_without_its_own_getdata_fails_clearly():
    page = FakePage(_sui(), [], warm_requests=False)
    d, _, _ = _setup(page, capture_timeout_s=0)
    with pytest.raises(DrillError, match="never asked for its data"):
        asyncio.run(d.fetch(URL, 100))
    assert page.posts == []


def test_getdata_only_ever_goes_from_the_warm_tab():
    page = FakePage(_sui(), [{"totalResults": 1, "hasMore": False, "items": _items(1)}])
    d, s, b = _setup(page)
    d.prewarm = lambda url: None                             # (normally the answered query warms the tab at once)
    q = asyncio.run(d.query(URL))
    assert d.tab is None                                     # the SUI query needs no warm tab
    asyncio.run(d.lines(q, URL, 100))
    assert d.tab is page and len(s.made) == 1
    d.tab.closed = True                                      # the warm tab died: getData must not use another tab
    with pytest.raises(DrillError, match="warm drill tab is gone"):
        d._warm_tab(s)


def test_signed_out_mid_paging_rewarms_once():
    page = FakePage(_sui(), [{"totalResults": 3, "hasMore": True, "items": _items(1)}])
    real = page.fetch
    calls = {"n": 0}

    def flaky(url, opts):
        if "getData" in url:
            calls["n"] += 1
            if calls["n"] == 2:
                return 200, "<html>epmexit.jsp</html>"
            if calls["n"] == 3:
                return 200, json.dumps({"responsePayload": {"totalResults": 3, "hasMore": False, "items": _items(2)}})
        return real(url, opts)
    page.fetch = flaky
    d, s, b = _setup(page)
    r = asyncio.run(d.fetch(URL, 100))
    assert len(r["items"]) == 3 and b.logins == 1 and len(s.made) == 2


def test_a_rewarm_already_done_by_another_piece_is_not_repeated():
    page = FakePage(_sui(), [])
    d, s, b = _setup(page)
    d._warm(b, s, URL)
    seen = d.gen
    d._rewarm(b, s, URL, seen)                               # the first piece that saw the sign-out: rewarms
    assert b.logins == 1 and d.gen == seen + 1
    d._rewarm(b, s, URL, seen)                               # a second piece with the same stale view: nothing
    assert b.logins == 1 and d.gen == seen + 1


def test_a_lost_drill_query_is_asked_once_more():
    page = FakePage(_sui(), [])
    answers = [(200, "<html>sign-in</html>"), (200, _sui())]
    page.fetch = lambda url, opts: answers.pop(0)
    d, _, _ = _setup(page)
    assert asyncio.run(d.query(URL)).app_id == 52
    answers[:] = [(200, "<html/>"), (200, "<html/>")]
    with pytest.raises(dd.SignedOut):
        asyncio.run(d.query(URL))


def test_pieces_run_at_once():
    """Several pieces' requests are in flight together (the browser answers them later, out of order)."""
    page = FakePage(_sui(), [])
    d, s, b = _setup(page)
    pending: list = []

    def request(prep, timeout_s):
        f = cf.Future()
        page_, url, opts = prep(s)
        pending.append((f, url))
        return f
    b.request = request
    d._warm(b, s, URL)
    d.views[(52, 39)] = {"items": []}

    async def main():
        tasks = [asyncio.ensure_future(d.lines(dd.parse_sui(_sui()), URL, 100)) for _ in range(3)]
        for _ in range(20):
            await asyncio.sleep(0)
        assert len(pending) == 3                             # all three getData requests out before any answer
        for f, _ in reversed(pending):
            f.set_result((200, json.dumps({"responsePayload": {"totalResults": 1, "hasMore": False,
                                                               "items": _items(1)}})))
        return await asyncio.gather(*tasks)
    assert [len(r["items"]) for r in asyncio.run(main())] == [1, 1, 1]


def test_prewarm_warms_once_and_lines_then_needs_no_warm():
    page = FakePage(_sui(), [{"totalResults": 1, "hasMore": False, "items": _items(1)}])
    d, s, b = _setup(page)
    d.prewarm(URL)
    d.prewarm(URL)                                           # tab alive: nothing more
    assert len(s.made) == 1 and d.tab is page
    asyncio.run(d.fetch(URL, 100))
    assert len(s.made) == 1


def test_a_failed_prewarm_is_only_logged_and_lines_warm_again():
    page = FakePage(_sui(), [{"totalResults": 1, "hasMore": False, "items": _items(1)}], warm_requests=False)
    d, s, b = _setup(page, capture_timeout_s=0)
    d.prewarm(URL)                                           # the page never asks for data: logged, not raised
    assert d.tab is None and not d._prewarming
    page.warm_requests = True
    assert len(asyncio.run(d.fetch(URL, 100))["items"]) == 1


def test_an_answered_drill_query_prewarms_and_a_refused_one_does_not():
    page = FakePage(json.dumps({"status": 1, "details": "EPMFDM-140373: too wide"}), [])
    d, s, b = _setup(page)
    with pytest.raises(CellTooWide):
        asyncio.run(d.query(URL))
    assert s.made == []                                      # a too-wide page would never ask for data
    page.sui = _sui()
    asyncio.run(d.query(URL))
    assert len(s.made) == 1 and d.tab is page


def test_a_warm_tab_of_an_old_browser_session_is_not_used():
    page = FakePage(_sui(), [])
    d, s, b = _setup(page)
    d._warm(b, s, URL)
    assert d._alive(s)
    s.ctx = type("NewCtx", (), {"new_page": lambda self: page})()   # the browser was killed and started again
    assert not d._alive(s)
    with pytest.raises(DrillError, match="warm drill tab is gone"):
        d._warm_tab(s)


def test_the_drill_sets_its_own_sign_in_time():
    page = FakePage(_sui(), [])
    d, s, b = _setup(page)
    s.deadline = 0.0                                         # an old job's deadline, long past
    asyncio.run(d.query(URL))
    import time
    assert s.deadline > time.monotonic() + config.DRILL_SETUP_S - 5


def test_a_cancelled_prewarm_does_not_switch_prewarming_off_for_good():
    page = FakePage(_sui(), [])
    d, s, b = _setup(page)

    def cancelled(fn):
        f = cf.Future()
        f.cancel()
        return f
    b.submit = cancelled
    d.prewarm(URL)
    assert not d._prewarming
