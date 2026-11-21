"""drill_token: the drill over the user's token - the session key, EPM's drill query, the view and the getData pages
(figures invented). The session is faked: an object with request() that records each call and answers by path, or a
real epmhttp.EpmSession over a fake opener. Nothing here reaches EPM."""
import asyncio
import dataclasses
import http.client
import io
import json
import logging
import threading
import time
import urllib.error
from typing import NamedTuple
from urllib.parse import urlsplit

import pytest

from smartass.core import config
from smartass.epm import drill_token, drillcore, epmhttp, oauth

_REAL_SHARED = drill_token.shared      # read at import: conftest swaps drill_token.shared for a guard in every test

URL = ("aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.drilldown&DRILLURL=x&sso_token=t"
       "&attribute=Account.id.C60")
EXIT_302 = (302, "", {"location": "https://plan1.example/interop/epmexit.jsp"})
VIEW = {"items": [{"name": "AMOUNTX"}]}


def SUI(**kw):
    return 200, json.dumps({"DRILL_LOG": "/tmp/x", "APPLICATION_ID": 52, "PARTITIONKEY": 39, "status": 0,
                            "QUERY": "SELECT 1", "PARAMSARRAY": ["39", "C60"],
                            "PARAMSTYPEARRAY": ["BigDecimal", "String"], **kw})


def PAGE(items, total, more=False):
    return 200, json.dumps({"responsePayload": {"items": [{"AMOUNTX": i} for i in range(items)],
                                                "totalResults": total, "hasMore": more}})


class Call(NamedTuple):
    method: str
    url: str
    headers: dict
    body: str | None
    path: str


def _next(answers):
    return answers.pop(0) if len(answers) > 1 else answers[0]


class FakeHttp:
    """EpmSession.request as the drill sees it: records every call and answers by path, in order (the last answer
    repeats). keys: the servicetype answers' x-appbuilder-client-id (None: no such header); an answer that is an
    exception is raised. slow_post: seconds a getData call takes (its answer and the in-flight count are taken on
    entry); post_s: the same per getData body (a function). dead: keys whose getData is answered with the exit
    redirect whatever comes next in lines; "old": every key but the newest one handed out. events: every request's
    start and end, in order (kind, method, path, key, body)."""

    def __init__(self, keys=("k1",), sui=None, view=None, lines=(PAGE(1, 1),), slow_post=0.0, dead=(), post_s=None):
        self.keys, self.sui, self.dead, self.handed = list(keys), list(sui or [SUI()]), dead, []
        self.view, self.lines = list(view or [(200, json.dumps(VIEW))]), list(lines)
        self.slow_post, self.post_s, self.calls, self.events = slow_post, post_s, [], []
        self.lock = threading.Lock()
        self.in_flight = self.max_posts_in_flight = 0

    def request(self, method, url, headers=None, body=None, timeout=120.0):
        path = urlsplit(url).path
        mark = (method, path, (headers or {}).get("x-appbuilder-client-id"), body)
        with self.lock:
            self.calls.append(Call(method, url, dict(headers or {}), body, path))
            self.events.append(("start", *mark))
            if path == "/aif/ui/model/applications/servicetype":
                k = _next(self.keys)
                self.handed.append(k)
                self.events.append(("end", *mark))
                return epmhttp.Response(200, "{}", {} if k is None else {"x-appbuilder-client-id": k})
            if method == "POST":
                key = (headers or {}).get("x-appbuilder-client-id")
                dead = key != self.handed[-1] if self.dead == "old" else key in self.dead
                answer = EXIT_302 if dead else _next(self.lines)
                self.in_flight += 1
                self.max_posts_in_flight = max(self.max_posts_in_flight, self.in_flight)
            else:
                answer = _next(self.sui if path.endswith("FdmeeMain.jsp") else self.view)
        try:
            if method == "POST":
                time.sleep(self.post_s(body) if self.post_s else self.slow_post)
        finally:
            with self.lock:
                self.events.append(("end", *mark))
                if method == "POST":
                    self.in_flight -= 1
        if isinstance(answer, Exception):
            raise answer
        status, text, *headers = answer
        return epmhttp.Response(status, text, headers[0] if headers else {})


@pytest.fixture
def fake():
    return FakeHttp


def test_fetch_takes_the_key_once_then_query_view_lines(fake):
    h = fake(lines=[PAGE(items=3, total=3)])
    r = asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert [c.path for c in h.calls] == ["/aif/ui/model/applications/servicetype",
                                         "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp",
                                         "/aif/ui/model/view", "/aif/ui/model/drilldown/getData"]
    assert h.calls[0].headers == {"page_id": "smartass"}
    assert all(c.headers["x-appbuilder-client-id"] == "k1" and c.headers["page_id"] == "smartass" for c in h.calls[1:])
    assert "FORMAT=SUI" in h.calls[1].url and "DRILLURL" not in h.calls[1].url
    assert (h.calls[3].method == "POST" and "limit=5000&offset=0" in h.calls[3].url
            and h.calls[3].body == r["query"].body)
    assert len(r["items"]) == 3 and r["total"] == 3 and r["view"] == {"items": [{"name": "AMOUNTX"}]}


def test_lines_are_paged_until_has_more_ends(fake):
    h = fake(lines=[PAGE(items=5000, total=7000, more=True), PAGE(items=2000, total=7000)])
    r = asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 10_000))
    assert len(r["items"]) == 7000
    assert [c.url.split("offset=")[1] for c in h.calls if c.method == "POST"] == ["0", "5000"]


def test_a_lost_session_mid_paging_takes_a_new_key_and_resumes_at_the_same_offset(fake):   # Review Focus 2
    h = fake(keys=["k1", "k2"], lines=[PAGE(items=5000, total=7000, more=True), EXIT_302, PAGE(items=2000, total=7000)])
    r = asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 10_000))
    posts = [c for c in h.calls if c.method == "POST"]
    assert len(r["items"]) == 7000 and [p.url.split("offset=")[1] for p in posts] == ["0", "5000", "5000"]
    assert [p.headers["x-appbuilder-client-id"] for p in posts] == ["k1", "k1", "k2"]


def test_a_second_lost_session_is_an_error_that_names_the_offset(fake):
    h = fake(keys=["k1", "k2"], lines=[EXIT_302, EXIT_302])
    with pytest.raises(drillcore.DrillError, match="offset 0"):
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))


def test_a_page_instead_of_json_counts_as_a_lost_session(fake):
    h = fake(keys=["k1", "k2"], sui=[(200, "<html>sign in</html>"), SUI()], lines=[PAGE(items=1, total=1)])
    assert len(asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))["items"]) == 1
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 2


def test_twenty_pieces_at_once_take_the_key_once_and_at_most_four_are_in_flight(fake, monkeypatch):   # Review Focus 3
    monkeypatch.setattr(config, "DRILL_PARALLEL", 4)
    h = fake(lines=[PAGE(items=2, total=2)], slow_post=0.05)
    d = drill_token.TokenDrill("https://plan1.example", http=h)

    async def go():
        return await asyncio.gather(*(d.fetch(URL, 100) for _ in range(20)))
    assert all(len(r["items"]) == 2 for r in asyncio.run(go()))
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 1 and h.max_posts_in_flight <= 4


def test_a_renewal_needed_by_many_pieces_happens_once(fake, monkeypatch):
    monkeypatch.setattr(config, "DRILL_PARALLEL", 4)
    h = fake(keys=["k1", "k2"], dead={"k1"}, lines=[PAGE(items=1, total=1)], slow_post=0.02)   # k1's session is gone
    d = drill_token.TokenDrill("https://plan1.example", http=h)

    async def go():
        return await asyncio.gather(*(d.fetch(URL, 100) for _ in range(6)))
    assert all(len(r["items"]) == 1 for r in asyncio.run(go()))
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 2


def test_once_a_renewal_has_started_no_getdata_leaves_with_the_old_key(fake, monkeypatch):
    monkeypatch.setattr(config, "DRILL_PARALLEL", 4)
    h = fake(keys=["k1", "k2", "k3"], dead="old", lines=[EXIT_302, PAGE(items=1, total=1)], slow_post=0.05)
    d = drill_token.TokenDrill("https://plan1.example", http=h)

    async def go():
        return await asyncio.wait_for(asyncio.gather(*(d.fetch(URL, 100) for _ in range(20))), 20)
    assert all(len(r["items"]) == 1 for r in asyncio.run(go()))
    keys = [i for i, c in enumerate(h.calls) if c.path == "/aif/ui/model/applications/servicetype"]
    assert len(keys) == 2                       # only the ones in flight when it started carried k1 (sent before it)
    assert not [c for c in h.calls[keys[1]:] if c.headers.get("x-appbuilder-client-id") == "k1"]


def _outstanding_old_keys_when_a_new_key_is_used(events, old, new):
    """For each request that starts with the new key: how many requests with the old key are still out."""
    out, seen = 0, []
    for kind, _method, _path, key, _body in events:
        if key == old:
            out += 1 if kind == "start" else -1
        elif key == new and kind == "start":
            seen.append(out)
    return seen


def test_a_new_key_is_asked_only_after_every_request_out_with_the_old_one_has_returned(fake, monkeypatch):
    monkeypatch.setattr(config, "DRILL_PARALLEL", 4)
    h = fake(keys=["k1", "k2"], dead={"k1"}, post_s=lambda body: 0.3 if "slow" in body else 0.02)
    d = drill_token.TokenDrill("https://plan1.example", http=h)

    async def go():
        q = await d.query(URL)
        slow = dataclasses.replace(q, body='{"QUERY": "slow"}')
        return await asyncio.wait_for(asyncio.gather(d.lines(slow, URL, 100), d.lines(q, URL, 100)), 20)
    assert all(len(r["items"]) == 1 for r in asyncio.run(go()))
    ev = h.events
    slow_back = ev.index(("end", "POST", "/aif/ui/model/drilldown/getData", "k1", '{"QUERY": "slow"}'))
    keys = [i for i, e in enumerate(ev) if e[0] == "start" and e[2] == "/aif/ui/model/applications/servicetype"]
    assert len(keys) == 2 and slow_back < keys[1]       # the slow old-key request came back before the new key
    used = _outstanding_old_keys_when_a_new_key_is_used(ev, "k1", "k2")
    assert used and set(used) == {0}


def test_a_request_that_never_returns_holds_the_new_key_back_only_for_the_drain_bound(fake, monkeypatch, caplog):
    caplog.set_level(logging.DEBUG)
    monkeypatch.setattr(config, "DRILL_PARALLEL", 4)
    monkeypatch.setattr(drill_token, "_DRAIN_S", 0.2)
    h = fake(keys=["k1", "k2"], dead={"k1"}, post_s=lambda body: 1.5 if "stuck" in body else 0.02)
    d = drill_token.TokenDrill("https://plan1.example", http=h)
    done = {}

    async def piece(name, q):
        r = await d.lines(q, URL, 100)
        done[name] = time.monotonic()
        return r

    async def go():
        q = await d.query(URL)
        stuck = dataclasses.replace(q, body='{"QUERY": "stuck"}')
        t0 = time.monotonic()
        res = await asyncio.wait_for(asyncio.gather(piece("stuck", stuck), piece("other", q)), 20)
        return t0, res
    t0, res = asyncio.run(go())
    assert all(len(r["items"]) == 1 for r in res)
    assert done["other"] - t0 < 1.0 < done["stuck"] - t0     # the other piece did not wait for the stuck one
    shown = [x.getMessage() for x in caplog.records]
    assert [m for m in shown if "still out" in m]
    assert not [m for m in shown if "k1" in m or "k2" in m]


def test_a_second_loss_on_the_new_key_is_retried_once_more_on_the_third_key(fake, monkeypatch):
    monkeypatch.setattr(config, "DRILL_PARALLEL", 4)
    h = fake(keys=["k1", "k2", "k3"], dead={"k1", "k2"}, slow_post=0.02)
    d = drill_token.TokenDrill("https://plan1.example", http=h)

    async def go():
        return await asyncio.wait_for(asyncio.gather(*(d.fetch(URL, 100) for _ in range(4))), 20)
    assert all(len(r["items"]) == 1 for r in asyncio.run(go()))
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 3
    assert [c.headers["x-appbuilder-client-id"] for c in h.calls if c.method == "POST"].count("k3") == 4


def test_a_piece_never_retries_twice_on_the_same_key(fake):
    h = fake(keys=["k1", "k2"], lines=[PAGE(1, 1), EXIT_302])   # EPM hands k2 out again: the same key generation
    d = drill_token.TokenDrill("https://plan1.example", http=h)
    asyncio.run(d.fetch(URL, 100))
    before = len(h.calls)
    with pytest.raises(drillcore.DrillError, match="offset 0"):
        asyncio.run(d.fetch(URL, 100))
    assert [c.headers["x-appbuilder-client-id"] for c in h.calls[before:] if c.method == "POST"] == ["k1", "k2", "k2"]


def test_getdata_lost_after_every_new_key_stops_at_the_cap(fake):
    h = fake(keys=["k1", "k2", "k3", "k4", "k5", "k6"], lines=[PAGE(1, 1), EXIT_302])
    d = drill_token.TokenDrill("https://plan1.example", http=h)
    asyncio.run(d.fetch(URL, 100))
    before = len(h.calls)
    with pytest.raises(drillcore.DrillError, match="offset 0"):
        asyncio.run(asyncio.wait_for(d.fetch(URL, 100), 20))
    calls = h.calls[before:]
    assert [c.path for c in calls].count("/aif/ui/model/applications/servicetype") == drill_token.MAX_RENEWALS
    assert [c.headers["x-appbuilder-client-id"] for c in calls if c.method == "POST"] == ["k1", "k2", "k3", "k4"]


def test_a_failed_renewal_is_one_request_and_one_error_for_every_piece_waiting_on_it(fake, monkeypatch):
    monkeypatch.setattr(config, "DRILL_PARALLEL", 4)
    h = fake(keys=["k1", None], dead={"k1"}, slow_post=0.05)

    async def go():
        d = drill_token.TokenDrill("https://plan1.example", http=h)
        return await asyncio.gather(*(d.fetch(URL, 100) for _ in range(6)), return_exceptions=True)
    res = asyncio.run(go())
    assert all(isinstance(e, drillcore.DrillError) and "no session key" in str(e) for e in res)
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 2


def test_a_piece_that_queued_with_the_lost_key_shares_its_failed_renewal_too(fake, monkeypatch):
    monkeypatch.setattr(config, "DRILL_PARALLEL", 1)       # one getData at a time: the third piece gets its turn only
    h = fake(keys=["k1", None], dead={"k1"}, slow_post=0.1)   # after the renewal has failed for the first two

    async def go():
        d = drill_token.TokenDrill("https://plan1.example", http=h)
        return await asyncio.gather(*(d.fetch(URL, 100) for _ in range(3)), return_exceptions=True)
    res = asyncio.run(go())
    assert all(isinstance(e, drillcore.DrillError) and "no session key" in str(e) for e in res)
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 2
    assert len([c for c in h.calls if c.method == "POST"]) == 1


@pytest.mark.parametrize("sui, exc", [({"details": "EPMFDM-140373 more than 3 dimensions"}, drillcore.CellTooWide),
                                      ({"details": "The drill query does not return any data"}, drillcore.NoDrillData),
                                      ({"details": "something else"}, drillcore.DrillError)])
def test_epms_refusals_are_raised_as_before_and_no_lines_are_asked(fake, sui, exc):
    h = fake(sui=[(200, json.dumps(sui))])
    with pytest.raises(exc):
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert not [c for c in h.calls if c.method == "POST"]


def test_more_lines_than_allowed_is_too_many_lines_with_the_total(fake):
    h = fake(lines=[PAGE(items=5000, total=9000, more=True)])
    with pytest.raises(drillcore.TooManyLines) as e:
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 6000))
    assert e.value.total == 9000


def test_no_key_from_epm_is_a_clear_error_without_the_url(fake):
    h = fake(keys=[None])
    with pytest.raises(drillcore.DrillError, match="no session key") as e:
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert "sso_token" not in str(e.value)


def test_a_drill_after_a_failed_key_request_asks_for_a_key_again(fake):
    h = fake(keys=[None, "k2"])
    d = drill_token.TokenDrill("https://plan1.example", http=h)
    with pytest.raises(drillcore.DrillError, match="no session key"):
        asyncio.run(d.fetch(URL, 100))
    assert len(asyncio.run(d.fetch(URL, 100))["items"]) == 1
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 2


def test_lines_for_a_built_query_reuse_the_view(fake):                 # drill_derive's siblings
    h = fake(lines=[PAGE(items=1, total=1), PAGE(items=2, total=2)])
    d = drill_token.TokenDrill("https://plan1.example", http=h)
    first = asyncio.run(d.fetch(URL, 100))
    q2 = dataclasses.replace(first["query"], body='{"QUERY": "other"}')
    second = asyncio.run(d.lines(q2, URL, 100))
    assert len(second["items"]) == 2 and [c.path for c in h.calls].count("/aif/ui/model/view") == 1
    assert [c for c in h.calls if c.method == "POST"][-1].body == '{"QUERY": "other"}'


# ---------------------------------------------------------------- beyond the brief
def test_a_refusal_takes_no_new_key_and_is_not_asked_again(fake):
    h = fake(keys=["k1", "k2"], sui=[(200, json.dumps({"details": "something else"}))])
    with pytest.raises(drillcore.DrillError, match="EPM refused the drill"):
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 1 and len(h.calls) == 2


def test_a_drill_query_lost_after_every_new_key_stops_at_the_cap_signed_out(fake):
    h = fake(keys=["k1", "k2", "k3", "k4", "k5", "k6"], sui=[SUI(), EXIT_302])
    d = drill_token.TokenDrill("https://plan1.example", http=h)
    asyncio.run(d.fetch(URL, 100))
    before = len(h.calls)
    with pytest.raises(drillcore.SignedOut) as e:
        asyncio.run(d.fetch(URL, 100))
    assert str(e.value) == "EPM answered the drill query with a page, not JSON (signed out?)"     # as parse_sui's
    paths = [c.path for c in h.calls[before:]]
    assert paths.count("/aif/ui/model/applications/servicetype") == drill_token.MAX_RENEWALS == 3
    assert paths.count("/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp") == drill_token.MAX_RENEWALS + 1


def test_another_http_status_for_the_drill_query_is_an_error_not_a_retry(fake):
    h = fake(sui=[(500, "<html>internal error</html>")])
    with pytest.raises(drillcore.DrillError, match="HTTP 500") as e:
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert not isinstance(e.value, drillcore.SignedOut) and len(h.calls) == 2


def test_a_lost_view_takes_a_new_key_and_the_view_once_more(fake):
    h = fake(keys=["k1", "k2"], view=[(403, "forbidden"), (200, json.dumps(VIEW))])
    r = asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert r["view"] == VIEW
    views = [c for c in h.calls if c.path == "/aif/ui/model/view"]
    assert [v.headers["x-appbuilder-client-id"] for v in views] == ["k1", "k2"]
    assert "page=DRILL_THROUGH&appId=52&partitionkey=39&_=" in views[0].url


def test_the_view_lost_twice_is_signed_out(fake):
    h = fake(keys=["k1", "k2"], view=[(200, "<html>sign in</html>")])
    with pytest.raises(drillcore.SignedOut) as e:
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert str(e.value) == "EPM answered the drill view with a page, not JSON (signed out?)"
    assert not [c for c in h.calls if c.method == "POST"]


def test_a_new_key_takes_the_view_again(fake):
    h = fake(keys=["k1", "k2"], lines=[PAGE(1, 1), EXIT_302, PAGE(1, 1)])
    d = drill_token.TokenDrill("https://plan1.example", http=h)
    first = asyncio.run(d.fetch(URL, 100))
    second = asyncio.run(d.fetch(URL, 100))
    assert first["view"] == second["view"] == VIEW
    views = [c for c in h.calls if c.path == "/aif/ui/model/view"]
    assert [v.headers["x-appbuilder-client-id"] for v in views] == ["k1"]
    assert asyncio.run(d.fetch(URL, 100))["view"] == VIEW
    views = [c for c in h.calls if c.path == "/aif/ui/model/view"]
    assert [v.headers["x-appbuilder-client-id"] for v in views] == ["k1", "k2"]


def test_a_lost_request_mid_paging_is_sent_once_more_at_the_same_offset(fake):
    h = fake(keys=["k1", "k2"], lines=[PAGE(5000, 7000, more=True), epmhttp.RequestLost("IncompleteRead"),
                                       PAGE(2000, 7000)])
    r = asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 10_000))
    assert len(r["items"]) == 7000
    assert [c.url.split("offset=")[1] for c in h.calls if c.method == "POST"] == ["0", "5000", "5000"]


def test_a_lost_drill_query_is_asked_once_more(fake):
    h = fake(keys=["k1", "k2"], sui=[TimeoutError("timed out"), SUI()])
    r = asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert len(r["items"]) == 1


def test_a_lost_request_is_logged_without_its_query_or_its_error_text(fake, caplog):
    caplog.set_level(logging.DEBUG)
    h = fake(keys=["k1", "k2"], sui=[TimeoutError("timed out on ...?sso_token=SSO-SECRET"), SUI()])
    asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(SECRET_URL, 100))
    shown = [x.getMessage() for x in caplog.records]
    assert [m for m in shown if "was lost (TimeoutError)" in m]
    assert not [m for m in shown if "SSO-SECRET" in m or "FORMAT=" in m]


def test_not_signed_in_is_raised_as_it_is(fake):
    h = fake(sui=[oauth.NotSignedIn("sign in at https://login.example")])
    with pytest.raises(oauth.NotSignedIn):
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))


def test_progress_is_told_each_page(fake):
    h = fake(lines=[PAGE(5000, 7000, more=True), PAGE(2000, 7000)])
    seen = []
    asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 10_000,
                                                                               lambda n, t: seen.append((n, t))))
    assert seen == [(5000, 7000), (7000, 7000)]


def test_one_drill_per_process(monkeypatch):
    monkeypatch.setattr(drill_token, "shared", _REAL_SHARED)
    monkeypatch.setattr(drill_token, "_shared", None)
    monkeypatch.setattr(config, "EPM_BASE", "https://plan1.example")
    monkeypatch.setattr(epmhttp, "_sessions", {})
    d = drill_token.shared()
    assert d is drill_token.shared() and d.http is epmhttp.session("https://plan1.example")


# ---------------------------------------------------------------- through the real session: every request is allowed
class _Tokens:
    def scope_for(self, base):
        return "urn:scope" if base == "https://plan1.example" else None

    def access_token(self, base, reject=None):
        return "AT-SECRET"


class _Answer(io.BytesIO):
    def __init__(self, status, text, headers):
        super().__init__(text.encode("utf-8"))
        self.status, self.headers = status, _message(headers)

    def info(self):
        return self.headers


def _message(headers):
    m = http.client.HTTPMessage()
    for k, v in headers.items():
        m[k] = v
    return m


class _Opener:
    """urllib's opener as EpmSession uses it, answering by path like FakeHttp (a status outside 2xx as HTTPError)."""

    def __init__(self, keys, lines):
        self.keys, self.lines, self.sent = list(keys), list(lines), []

    def open(self, req, timeout):
        self.sent.append(req)
        path = urlsplit(req.full_url).path
        if path == "/aif/ui/model/applications/servicetype":
            status, text, headers = 200, "{}", {"X-AppBuilder-Client-Id": _next(self.keys)}
        elif path.endswith("FdmeeMain.jsp"):
            status, text, headers = *SUI(), {}
        elif path == "/aif/ui/model/view":
            status, text, headers = 200, json.dumps(VIEW), {}
        else:
            status, text, *rest = _next(self.lines)
            headers = rest[0] if rest else {}
        if 200 <= status < 300:
            return _Answer(status, text, headers)
        raise urllib.error.HTTPError(req.full_url, status, "no", _message(headers), io.BytesIO(text.encode("utf-8")))


SECRET_URL = URL.replace("sso_token=t", "sso_token=SSO-SECRET")


def _session(keys, lines):
    op = _Opener(keys, lines)
    return epmhttp.EpmSession("https://plan1.example", tokens=_Tokens(), opener=op), op


def test_every_request_fits_the_sessions_allowlist_and_the_key_reaches_the_wire(caplog):
    caplog.set_level(logging.DEBUG)
    s, op = _session(["KEY-SECRET-1", "KEY-SECRET-2"], [PAGE(5000, 7000, more=True), EXIT_302, PAGE(2000, 7000)])
    r = asyncio.run(drill_token.TokenDrill("https://plan1.example", http=s).fetch(SECRET_URL, 10_000))
    assert len(r["items"]) == 7000 and not [x for x in caplog.records if "refused" in x.getMessage()]
    assert [urlsplit(x.full_url).path.rsplit("/", 1)[1] for x in op.sent] == [
        "servicetype", "FdmeeMain.jsp", "view", "getData", "getData", "servicetype", "getData"]
    assert all(x.full_url.startswith("https://plan1.example/") for x in op.sent)
    assert [x.get_header("X-appbuilder-client-id") for x in op.sent] == [
        None, "KEY-SECRET-1", "KEY-SECRET-1", "KEY-SECRET-1", "KEY-SECRET-1", None, "KEY-SECRET-2"]
    assert all(x.get_header("Page_id") == "smartass" and x.get_header("Authorization") == "Bearer AT-SECRET"
               for x in op.sent)
    assert json.loads(op.sent[3].data) == {"QUERY": "SELECT 1", "PARAMS": ["39", "C60"],
                                           "PARAMSTYPE": ["BigDecimal", "String"]}


def test_no_key_token_or_query_in_a_log_line_or_an_error(caplog):
    caplog.set_level(logging.DEBUG)
    s, _ = _session(["KEY-SECRET-1", "KEY-SECRET-2"], [EXIT_302])
    with pytest.raises(drillcore.DrillError) as e:
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=s).fetch(SECRET_URL, 100))
    shown = [x.getMessage() for x in caplog.records] + [str(e.value), repr(e.value)]
    assert any("expired" in m for m in shown)                         # the renewal was logged
    for secret in ("KEY-SECRET", "AT-SECRET", "SSO-SECRET", "limit=", "offset="):
        assert not [m for m in shown if secret in m], secret
