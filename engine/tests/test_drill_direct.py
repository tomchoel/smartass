"""drill_direct: drill lines from EPM's own drill query, without a page per piece (figures invented)."""
import json

import pytest

from smartview_mcp import config, drill_direct as dd
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

    def evaluate(self, js, args):
        url = args[0]
        if "FORMAT=SUI" in url:
            self.gets.append((url, args[1]))
            return {"s": 200, "t": self.sui}
        if "view?page=DRILL_THROUGH" in url:
            self.gets.append((url, args[1]))
            return {"s": 200, "t": json.dumps({"items": [{"viewName": "DrillFilter", "viewDefinition": "[]"}]})}
        self.posts.append((url, args[1], args[2]))
        return {"s": 200, "t": json.dumps({"responsePayload": self.pages.pop(0)})}


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
    base = "https://epm"

    def __init__(self):
        self.logins = 0

    def _ensure(self, s):
        pass

    def _login(self, s):
        self.logins += 1


def _items(n, amount="-1"):
    return [{"AMOUNTX": amount}] * n


def test_fetch_warms_once_then_reuses_the_tab():
    page = FakePage(_sui(), [{"totalResults": 2, "hasMore": False, "items": _items(2)},
                             {"totalResults": 1, "hasMore": False, "items": _items(1)}])
    d, s, b = dd.DirectDrill(), FakeSession(page), FakeBrowser()
    r1 = d.fetch_on(b, s, URL, 100)
    r2 = d.fetch_on(b, s, URL, 100)
    assert len(r1["items"]) == 2 and len(r2["items"]) == 1
    assert len(s.made) == 1                                   # one warm page for both pieces
    assert r1["view"]["items"][0]["viewName"] == "DrillFilter"
    url, body, headers = page.posts[0]
    assert f"getData?limit={config.DRILL_PAGE}&offset=0" in url and json.loads(body)["QUERY"] == "SELECT 1"
    assert headers["x-appbuilder-client-id"] == "cid" and "cookie" not in headers


def test_no_data_needs_no_warm_page():
    page = FakePage(json.dumps({"status": 1, "details": "Drill query does not return any data."}), [])
    s = FakeSession(page)
    with pytest.raises(dd.NoDrillData):
        dd.DirectDrill().fetch_on(FakeBrowser(), s, URL, 100)
    assert s.made == [] and page.posts == []                # getData never sent


def test_warm_page_without_its_own_getdata_fails_clearly():
    page = FakePage(_sui(), [], warm_requests=False)
    with pytest.raises(DrillError, match="never asked for its data"):
        dd.DirectDrill(capture_timeout_s=0).fetch_on(FakeBrowser(), FakeSession(page), URL, 100)
    assert page.posts == []


def test_signed_out_mid_paging_rewarms_once():
    page = FakePage(_sui(), [{"totalResults": 3, "hasMore": True, "items": _items(1)}])
    real = page.evaluate
    calls = {"n": 0}

    def flaky(js, args):
        if "getData" in args[0]:
            calls["n"] += 1
            if calls["n"] == 2:
                return {"s": 200, "t": "<html>epmexit.jsp</html>"}
            if calls["n"] == 3:
                return {"s": 200, "t": json.dumps({"responsePayload": {"totalResults": 3, "hasMore": False,
                                                                       "items": _items(2)}})}
        return real(js, args)
    page.evaluate = flaky
    b, s = FakeBrowser(), FakeSession(page)
    r = dd.DirectDrill().fetch_on(b, s, URL, 100)
    assert len(r["items"]) == 3 and b.logins == 1 and len(s.made) == 2
