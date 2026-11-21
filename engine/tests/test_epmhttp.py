"""epm/epmhttp.py: the token session's allowlist, the pod's origin, the bearer, the cookie jar, a 401 and a redirect.
The network is faked: an opener that records what it is given (or, for the redirect handler, urllib's real opener
over a fake HTTPS handler). Nothing here reaches EPM."""
import http.client
import io
import json
import logging
import time
import urllib.error
import urllib.request
import urllib.response

import pytest

from smartass.epm import epmhttp, oauth

POD = "https://plan1.example"


class _Tokens:
    """epm/oauth.py as the session sees it: scope_for and access_token (which may raise)."""

    def __init__(self, *tokens, scopes=("https://plan1.example",)):
        self.tokens, self.scopes, self.asked = list(tokens), scopes, []

    def scope_for(self, base):
        return "urn:scope" if base in self.scopes else None

    def access_token(self, base, reject=None):
        self.asked.append((base, reject))
        t = self.tokens.pop(0) if len(self.tokens) > 1 else self.tokens[0]
        if isinstance(t, Exception):
            raise t
        return t


def _message(headers):
    m = http.client.HTTPMessage()
    for k, v in headers.items():
        m[k] = v
    return m


class _Answer(io.BytesIO):
    """A 2xx answer as urllib hands it back."""

    def __init__(self, status, text, headers):
        super().__init__(text.encode("utf-8"))
        self.status, self.headers = status, _message(headers)

    def info(self):
        return self.headers


class _Opener:
    """Records every request; answers (status, text, headers) in order, the last one repeats. A status outside 2xx is
    raised as urllib.error.HTTPError, as urllib's opener does; an exception is raised as it is."""

    def __init__(self, answers):
        self.answers, self.sent, self.timeouts = list(answers), [], []

    def open(self, req, timeout):
        self.sent.append(req)
        self.timeouts.append(timeout)
        a = self.answers.pop(0) if len(self.answers) > 1 else self.answers[0]
        if isinstance(a, Exception):
            raise a
        status, text, headers = a
        if 200 <= status < 300:
            return _Answer(status, text, headers)
        raise urllib.error.HTTPError(req.full_url, status, "no", _message(headers), io.BytesIO(text.encode("utf-8")))


@pytest.fixture
def make():
    def build(tokens=("at-1",), answers=((200, "{}", {}),), base=POD):
        tok, op = _Tokens(*tokens), _Opener(answers)
        return epmhttp.EpmSession(base, tokens=tok, opener=op), op, tok
    return build


def test_only_the_allowlist_is_sent(make):                      # make(...) -> (session, opener, tokens)
    s, op, tok = make()
    for method, url, body in [("POST", "/aif/ui/model/integration/run", "{}"),
                              ("DELETE", "/aif/ui/model/view?page=DRILL_THROUGH", None),
                              ("GET", "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=x", None),
                              ("GET", "/aif/ui/model/view?page=SETUP", None),
                              ("POST", "/HyperionPlanning/SmartView", "<req_SubmitData></req_SubmitData>"),
                              ("GET", "/HyperionPlanning/rest/v3/applications", None)]:
        with pytest.raises(ValueError, match="refused"):
            s.request(method, url, body=body)
    assert op.sent == [] and tok.asked == []


def test_the_bearer_never_leaves_the_pod(make):
    s, op, tok = make()
    for url in ("https://evil.example/aif/ui/model/drilldown/getData",
                "http://plan1.example/aif/ui/model/view?page=DRILL_THROUGH",
                "//evil.example/aif/ui/model/view?page=DRILL_THROUGH"):
        with pytest.raises(epmhttp.Refused):
            s.request("GET" if "view" in url else "POST", url, body=None if "view" in url else "{}")
    assert op.sent == []


def test_a_pod_without_a_scope_is_refused(make):
    with pytest.raises(epmhttp.Refused, match="no OAuth scope"):
        epmhttp.EpmSession("https://other.example", tokens=make()[2]).request(
            "GET", "/aif/ui/model/applications/servicetype")


def test_a_request_carries_the_bearer_and_the_sessions_cookies(make):
    s, op, tok = make(answers=[(200, "{}", {"Set-Cookie": "JSESSIONID=abc; Path=/"}), (200, "{}", {})])
    s.request("GET", "/aif/ui/model/applications/servicetype", headers={"page_id": "smartass"})
    s.request("GET", "/aif/ui/model/view?page=DRILL_THROUGH&appId=1&partitionkey=2")
    assert op.sent[0].get_header("Authorization") == "Bearer at-1" and op.sent[0].get_header("Page_id") == "smartass"
    assert "JSESSIONID=abc" in (op.sent[1].get_header("Cookie") or "")


def test_a_401_gets_one_new_token_and_one_more_try(make):        # Review Focus 1
    s, op, tok = make(tokens=["at-1", "at-2"], answers=[(401, "", {}), (200, '{"ok": 1}', {})])
    r = s.request("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0", body="{}")
    assert r.status == 200 and [x.get_header("Authorization") for x in op.sent] == ["Bearer at-1", "Bearer at-2"]
    assert tok.asked == [(POD, None), (POD, "at-1")]
    s2, op2, _ = make(tokens=["at-1", "at-2"], answers=[(401, "", {})])
    assert s2.request("GET", "/aif/ui/model/applications/servicetype").status == 401 and len(op2.sent) == 2


def test_a_redirect_is_an_answer_not_followed(make):             # Review Focus 5
    s, op, _ = make(answers=[(302, "", {"Location": "https://evil.example/x"})])
    r = s.request("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0", body="{}")
    assert r.status == 302 and r.headers["location"] == "https://evil.example/x" and len(op.sent) == 1


def test_not_signed_in_is_raised_as_it_is(make):
    s, op, _ = make(tokens=[oauth.NotSignedIn("not signed in")])
    with pytest.raises(oauth.NotSignedIn):
        s.request("GET", "/aif/ui/model/applications/servicetype")


def test_one_session_per_pod(monkeypatch):
    monkeypatch.setattr(epmhttp, "_sessions", {})
    assert epmhttp.session("https://PLAN1.example:443/x") is epmhttp.session("https://plan1.example")


# ---------------------------------------------------------------- trying to break it
GET_VIEW = "/aif/ui/model/view?page=DRILL_THROUGH&appId=1&partitionkey=2"


@pytest.mark.parametrize("method, url, body, headers", [
    ("get", GET_VIEW, None, None),                                            # the method as EPM's allowlist has it
    ("GET", GET_VIEW, "{}", None),                                            # a GET carries no body
    ("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0", None, None),   # a POST carries one
    ("GET", "/aif/ui/model/view?xpage=DRILL_THROUGH", None, None),            # the whole item, not a substring
    ("GET", "/aif/ui/model/view?page=DRILL_THROUGH&page=SETUP", None, None),  # one value, not a choice
    ("GET", "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?FORMAT=SUI&FORMAT=HTML", None, None),
    ("GET", "/aif/ui/model/view?%70age=SETUP&page=DRILL_THROUGH", None, None),     # as the servlet decodes the name
    ("GET", "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?FORMAT=SUI&%46ORMAT=HTML", None, None),
    ("GET", "/aif/ui/model/view/?page=DRILL_THROUGH", None, None),            # the path exactly
    ("GET", "/AIF/ui/model/view?page=DRILL_THROUGH", None, None),
    ("GET", "/aif/ui/model/x/../view?page=DRILL_THROUGH", None, None),
    ("GET", "aif/ui/model/view?page=DRILL_THROUGH", None, None),              # a path starts at the root
    ("GET", "https://user@plan1.example/aif/ui/model/view?page=DRILL_THROUGH", None, None),
    ("GET", "https://plan1.example@evil.example/aif/ui/model/view?page=DRILL_THROUGH", None, None),
    ("GET", "https://plan1.example:8443/aif/ui/model/view?page=DRILL_THROUGH", None, None),
    ("GET", "https://plan1.example:x/aif/ui/model/view?page=DRILL_THROUGH", None, None),
    ("GET", "https://plan1.example.evil.example/aif/ui/model/view?page=DRILL_THROUGH", None, None),
    ("GET", "https:/aif/ui/model/view?page=DRILL_THROUGH", None, None),
    ("GET", "/aif/ui/model/view?page=DRILL_THROUGH\r\nHost: evil.example", None, None),   # urlsplit would drop these
    ("GET", "/aif/ui/model/view?page=DRILL_THROUGH\t", None, None),
    ("GET", "/aif/ui/model/view?page=DRILL_THROUGH&x=\u00e6", None, None),
    ("GET", "https://pl\u00e4n1.example/aif/ui/model/view?page=DRILL_THROUGH", None, None),     # an IDN host
    ("GET", "https://xn--pln1-loa.example/aif/ui/model/view?page=DRILL_THROUGH", None, None),
    ("GET", "https://plan1.example./aif/ui/model/view?page=DRILL_THROUGH", None, None),          # a trailing dot
    ("GET", "https://evil.example\\@plan1.example/aif/ui/model/view?page=DRILL_THROUGH", None, None),
    ("GET", "https://evil.example\\.plan1.example/aif/ui/model/view?page=DRILL_THROUGH", None, None),
    ("GET", "/aif\\ui/model/view?page=DRILL_THROUGH", None, None),                                 # a backslash
    ("GET", "/aif/ui/model/view;jsessionid=x?page=DRILL_THROUGH", None, None),                  # ;params
    ("GET", "/aif/ui/model%2Fview?page=DRILL_THROUGH", None, None),                             # %2F in the path
    ("GET", "/aif/ui/model/applications/servicetype?x=1", None, None),                          # query names
    ("GET", "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.drilldown&FORMAT=SUI&action=x", None,
     None),
    ("GET", "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.other&FORMAT=SUI", None, None),
    ("GET", "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?FORMAT=SUI", None, None),
    ("GET", "/aif/ui/model/view?page=DRILL_THROUGH&appId=1&appId=2", None, None),
    ("GET", "/aif/ui/model/view?page=DRILL_THROUGH&_=now", None, None),
    ("POST", "/aif/ui/model/drilldown/getData?limit=5000", "{}", None),
    ("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0&offset=1", "{}", None),
    ("POST", "/aif/ui/model/drilldown/getData?limit=all&offset=0", "{}", None),
    ("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0&delete=1", "{}", None),
    ("POST", "/aif/ui/model/drilldown/getData", "{}", None),
    ("GET", GET_VIEW, None, {"Authorization": "Bearer mine"}),                # the session's own headers
    ("GET", GET_VIEW, None, {"cookie": "JSESSIONID=other"}),
    ("GET", GET_VIEW, None, {"Host": "evil.example"}),
    ("GET", GET_VIEW, None, {"page_id": "smartass\r\nX-Evil: 1"}),           # no header can be smuggled in
    ("GET", GET_VIEW, None, {"bad name": "x"}),
])
def test_what_is_not_exactly_on_the_allowlist_is_refused_before_a_token(make, method, url, body, headers):
    s, op, tok = make()
    with pytest.raises(epmhttp.Refused, match="refused"):
        s.request(method, url, headers=headers, body=body)
    assert op.sent == [] and tok.asked == []


def test_the_request_goes_to_the_pods_origin_whatever_the_caller_wrote(make):
    s, op, _ = make(base="https://PLAN1.example:443/HyperionPlanning")
    s.request("GET", "https://Plan1.Example:443/aif/ui/model/view?page=DRILL_THROUGH&appId=%2F#frag")
    s.request("GET", "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.drilldown&FORMAT=SUI&_=1")
    assert [r.full_url for r in op.sent] == [
        "https://plan1.example/aif/ui/model/view?page=DRILL_THROUGH&appId=%2F",
        "https://plan1.example/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.drilldown&FORMAT=SUI&_=1"]
    assert [r.get_method() for r in op.sent] == ["GET", "GET"] and all(r.data is None for r in op.sent)


def test_the_cache_buster_is_a_millisecond_clock(make):                 # the drill sends int(time.time() * 1000)
    s, op, _ = make()
    now = str(int(time.time() * 1000))
    s.request("GET", f"/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.drilldown&FORMAT=SUI&_={now}")
    s.request("GET", f"/aif/ui/model/view?page=DRILL_THROUGH&appId=52&partitionkey=39&_={now}")
    assert len(op.sent) == 2 and len(now) == 13
    with pytest.raises(epmhttp.Refused):
        s.request("GET", "/aif/ui/model/view?page=DRILL_THROUGH&_=12345678901234")      # 14 digits: no clock
    assert len(op.sent) == 2


def test_smart_views_rcp_version_is_taken_on_the_drill_query(make):
    """Live 2026-11-21 (Task 7): Smart View's drill link carries rcp_version (digits and dots); every token-route
    drill query was refused for it. A version once, nothing else under that name."""
    s, op, _ = make()
    jsp = "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.drilldown&FORMAT=SUI"
    s.request("GET", f"{jsp}&rcp_version=24.06&attribute=x&_=1")
    s.request("GET", f"{jsp}&rcp_version=11.2.15.0")
    assert len(op.sent) == 2
    for bad in ("rcp_version=24.06&rcp_version=24.07", "rcp_version=24a", "rcp_version=", "rcp_version=1..2"):
        with pytest.raises(epmhttp.Refused):
            s.request("GET", f"{jsp}&{bad}")
    with pytest.raises(epmhttp.Refused):
        s.request("GET", "/aif/ui/model/view?page=DRILL_THROUGH&rcp_version=24.06")    # the drill query's only
    assert len(op.sent) == 2


def test_a_read_only_smart_view_request_and_a_getdata_post_are_sent(make):
    from smartass.epm import svxml
    s, op, _ = make(answers=[(200, "<res_ConnectToProvider/>", {"Content-Type": "text/xml"})])
    r = s.request("POST", "/HyperionPlanning/SmartView", headers={"Content-Type": "text/xml"}, body=svxml.connect())
    assert r == epmhttp.Response(200, "<res_ConnectToProvider/>", {"content-type": "text/xml"})
    s.request("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0", body='{"QUERY": "SELECT 1"}', timeout=7)
    assert [(x.get_method(), x.data) for x in op.sent] == [("POST", svxml.connect().encode()),
                                                            ("POST", b'{"QUERY": "SELECT 1"}')]


def test_the_session_keeps_its_cookies_to_itself(make):
    s, op, _ = make(answers=[(200, "{}", {"Set-Cookie": "JSESSIONID=abc; Path=/"}), (200, "{}", {})])
    s.request("GET", "/aif/ui/model/applications/servicetype")
    other, op2, _ = make()
    other.request("GET", "/aif/ui/model/applications/servicetype")
    assert op2.sent[0].get_header("Cookie") is None


def test_a_401_whose_token_cannot_be_renewed_raises_the_oauth_error(make):
    s, op, tok = make(tokens=["at-1", oauth.NotSignedIn("the sign-in has ended")], answers=[(401, "", {})])
    with pytest.raises(oauth.NotSignedIn, match="ended"):
        s.request("GET", "/aif/ui/model/applications/servicetype")
    assert len(op.sent) == 1


def test_a_403_is_an_answer_not_a_new_token(make):
    s, op, tok = make(tokens=["at-1", "at-2"], answers=[(403, "no", {})])
    assert s.request("GET", "/aif/ui/model/applications/servicetype").status == 403
    assert len(op.sent) == 1 and tok.asked == [(POD, None)]


def test_a_token_that_cannot_go_in_a_header_is_refused_without_showing_it(make):
    s, op, _ = make(tokens=["at-SECRET\r\nX-Evil: 1"])
    with pytest.raises(epmhttp.Refused) as e:
        s.request("GET", "/aif/ui/model/applications/servicetype")
    assert "SECRET" not in str(e.value) and "Authorization" in str(e.value) and op.sent == []


SECRETS = ("at-SECRET", "KEY-SECRET", "COOKIE-SECRET", "limit=", "sso_token")


def _clean(text):
    return not any(x in text for x in SECRETS)


def test_refusals_and_log_lines_never_carry_a_query_a_token_a_key_or_a_cookie(make, caplog):
    caplog.set_level(logging.DEBUG, logger="smartass.epm.epmhttp")
    s, op, _ = make(tokens=["at-SECRET"], answers=[(200, "{}", {"Set-Cookie": "JSESSIONID=COOKIE-SECRET; Path=/"})])
    key = {"x-appbuilder-client-id": "KEY-SECRET"}
    s.request("GET", "/aif/ui/model/applications/servicetype", headers={"page_id": "smartass"})
    s.request("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0", headers=key, body="{}")
    for url in ("/aif/ui/model/view?page=SETUP&sso_token=x&limit=1",
                "https://evil.example/aif/ui/model/view?page=DRILL_THROUGH&sso_token=x&limit=1",
                "/aif/ui/model/view?page=DRILL_THROUGH&sso_token=x&limit=1\n"):
        with pytest.raises(epmhttp.Refused) as e:
            s.request("GET", url, headers=key)
        assert _clean(str(e.value)), str(e.value)
    assert caplog.records and _clean(caplog.text), caplog.text
    assert "/aif/ui/model/drilldown/getData" in caplog.text                  # what was asked is still there


def test_a_network_error_is_raised_as_it_is(make):
    boom = TimeoutError("timed out")
    s, op, _ = make(answers=[boom])
    with pytest.raises(TimeoutError) as e:
        s.request("GET", "/aif/ui/model/applications/servicetype")
    assert e.value is boom


@pytest.mark.parametrize("error", [
    OSError("failed: https://plan1.example/aif/ui/model/drilldown/getData?limit=5000&offset=0"),
    ConnectionResetError("reset while sending Bearer at-SECRET"),
    urllib.error.URLError("no answer to https://plan1.example/x?limit=1 (key KEY-SECRET, cookie COOKIE-SECRET)"),
])
def test_a_network_error_that_quotes_a_secret_is_raised_without_it(make, error):
    s, op, _ = make(tokens=["at-SECRET"], answers=[(200, "{}", {"Set-Cookie": "JSESSIONID=COOKIE-SECRET; Path=/"}),
                                                     error])
    s.request("GET", "/aif/ui/model/applications/servicetype")
    with pytest.raises(OSError) as e:
        s.request("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0",
                  headers={"x-appbuilder-client-id": "KEY-SECRET"}, body="{}")
    assert isinstance(e.value, type(error)) and _clean(str(e.value)), str(e.value)
    assert e.value.__cause__ is None and e.value.__suppress_context__


class _Https(urllib.request.HTTPSHandler):
    """urllib's own opener, with the wire faked: every request is answered by a canned redirect."""

    def __init__(self, status, location):
        super().__init__()
        self.status, self.location, self.urls = status, location, []

    def https_open(self, req):
        self.urls.append(req.full_url)
        r = urllib.response.addinfourl(io.BytesIO(b""), _message({"Location": self.location}), req.full_url,
                                       self.status)
        r.msg = "Moved"
        return r


@pytest.mark.parametrize("status", [301, 302, 303, 307, 308])
@pytest.mark.parametrize("location", ["https://evil.example/x", "https://plan1.example/interop/epmexit.jsp"])
def test_the_real_opener_never_follows_a_redirect(status, location):
    wire = _Https(status, location)
    s = epmhttp.EpmSession(POD, tokens=_Tokens("at-1"), opener=epmhttp._opener(wire))
    r = s.request("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0", body="{}")
    assert r.status == status and r.headers["location"] == location
    assert wire.urls == [POD + "/aif/ui/model/drilldown/getData?limit=5000&offset=0"]


def test_the_default_opener_is_the_one_that_does_not_follow_redirects():
    s = epmhttp.EpmSession(POD)
    assert any(isinstance(h, epmhttp._NoRedirect) for h in s._opener.handlers)
    assert not any(type(h) is urllib.request.HTTPRedirectHandler for h in s._opener.handlers)


def test_sessions_are_per_pod(monkeypatch):
    monkeypatch.setattr(epmhttp, "_sessions", {})
    a, b = epmhttp.session("https://plan1.example"), epmhttp.session("https://strat.example/HyperionPlanning")
    assert a is not b and (a.origin, b.origin) == ("https://plan1.example", "https://strat.example")


GET_DATA = "/aif/ui/model/drilldown/getData?limit=5000&offset=0"


@pytest.mark.parametrize("name, value", [
    ("X-HTTP-Method-Override", "DELETE"), ("X-HTTP-Method", "PUT"), ("X-Method-Override", "DELETE"),
    ("Content-Length", "0"), ("Transfer-Encoding", "chunked"), ("Accept-Encoding", "gzip"), ("Connection", "upgrade"),
    ("TE", "trailers"), ("Upgrade", "h2c"), ("User-Agent", "x"), ("Proxy-Authorization", "Basic x"),
    ("X-Forwarded-Host", "evil.example"), ("X-Forwarded-For", "1.2.3.4"), ("Origin", "https://evil.example"),
    ("Referer", "https://evil.example/"), ("Authorization", "Bearer mine"), ("Cookie", "JSESSIONID=x"),
    ("Host", "evil.example"), ("Expect", "100-continue"), ("Range", "bytes=0-1"),
    ("X-Requested-With", "Other"),                                            # a header it takes, with another value
    ("Content-Type", "text/plain"), ("Content-Type", "application/x-www-form-urlencoded"),
])
def test_a_header_the_request_does_not_take_is_refused_before_a_token(make, name, value):
    s, op, tok = make()
    with pytest.raises(epmhttp.Refused, match="refused"):
        s.request("POST", GET_DATA, headers={"page_id": "smartass", name: value}, body="{}")
    assert op.sent == [] and tok.asked == []


def test_headers_are_taken_per_request(make):
    s, op, tok = make()
    with pytest.raises(epmhttp.Refused):                                      # Smart View takes its content type only
        s.request("POST", "/HyperionPlanning/SmartView", headers={"page_id": "x"}, body=_svxml().connect())
    with pytest.raises(epmhttp.Refused):
        s.request("POST", "/HyperionPlanning/SmartView", headers={"Content-Type": "application/json"},
                  body=_svxml().connect())
    with pytest.raises(epmhttp.Refused):                                      # one content type, not two
        s.request("POST", GET_DATA, headers={"Content-Type": "application/json", "content-type": "text/xml"},
                  body="{}")
    assert op.sent == [] and tok.asked == []
    drill = {"Accept": "application/json, text/javascript, */*; q=0.01", "Content-Type": "application/json",
             "X-Requested-With": "XMLHttpRequest", "page_id": "smartass", "x-appbuilder-client-id": "k1"}
    s.request("GET", GET_VIEW, headers=drill)                                 # what drill_token sends (Task 5)
    s.request("POST", GET_DATA, headers=drill, body='{"QUERY": "SELECT 1", "PARAMS": [], "PARAMSTYPE": []}')
    assert len(op.sent) == 2


def _svxml():
    from smartass.epm import svxml
    return svxml


def test_the_session_sets_its_own_framing_headers(make):
    s, op, _ = make()
    s.request("POST", GET_DATA, body="{}")
    s.request("POST", "/HyperionPlanning/SmartView", body=_svxml().connect())
    assert [r.get_header("Content-type") for r in op.sent] == ["application/json", "text/xml"]
    assert all(r.get_header("Accept-encoding") == "identity" for r in op.sent)


@pytest.mark.parametrize("body", ["[]", '"QUERY"', "not json", '{"QUERY": 1}', '{"QUERY": "q", "PARAMS": "x"}',
                                  '{"QUERY": "q", "DELETE": true}', '{"QUERY": "a", "QUERY": "b"}',
                                  '{"QUERY": "' + "x" * (1 << 20) + '"}'])
def test_a_getdata_body_that_is_not_a_drill_query_is_refused(make, body):
    s, op, tok = make()
    with pytest.raises(epmhttp.Refused, match="refused"):
        s.request("POST", GET_DATA, body=body)
    assert op.sent == [] and tok.asked == []


def test_a_refused_smart_view_body_is_not_quoted(make, caplog):
    caplog.set_level(logging.DEBUG, logger="smartass.epm.epmhttp")
    s, op, _ = make()
    with pytest.raises(epmhttp.Refused) as e:
        s.request("POST", "/HyperionPlanning/SmartView", body="<req_Foo><sID>SIDSECRET1234567890abcdef</sID></req_Foo>")
    assert str(e.value) == "refused: not a single read-only Smart View request"
    assert "SIDSECRET" not in caplog.text and "req_Foo" not in caplog.text and op.sent == []


def test_a_401_drops_the_cookies_that_came_with_the_refused_token(make):
    s, op, tok = make(tokens=["at-1", "at-2"], answers=[(200, "{}", {"Set-Cookie": "JSESSIONID=old; Path=/"}),
                                                       (401, "", {}), (200, "{}", {})])
    s.request("GET", "/aif/ui/model/applications/servicetype")
    s.request("GET", "/aif/ui/model/applications/servicetype")
    assert op.sent[1].get_header("Cookie") == "JSESSIONID=old" and op.sent[2].get_header("Cookie") is None
    assert op.sent[2].get_header("Authorization") == "Bearer at-2"


def test_a_response_hands_over_no_cookie_and_shows_no_key(make):
    s, op, _ = make(answers=[(302, "", {"Set-Cookie": "JSESSIONID=COOKIE-SECRET; Path=/",
                                        "x-appbuilder-client-id": "KEY-SECRET",
                                        "Location": "https://plan1.example/interop/epmexit.jsp?sso_token=T&limit=1"})])
    r = s.request("GET", "/aif/ui/model/applications/servicetype")
    assert "set-cookie" not in r.headers and r.headers["x-appbuilder-client-id"] == "KEY-SECRET"   # the caller's key
    assert _clean(repr(r)) and "KEY-SECRET" not in repr(r) and "302" in repr(r), repr(r)
    assert s._jar and [c.value for c in s._jar] == ["COOKIE-SECRET"]                          # the jar keeps it


def test_a_set_cookie_for_another_domain_is_ignored(make):
    s, op, _ = make(answers=[(200, "{}", {"Set-Cookie": "EVIL=1; Domain=evil.example; Path=/"}), (200, "{}", {})])
    s.request("GET", "/aif/ui/model/applications/servicetype")
    s.request("GET", "/aif/ui/model/applications/servicetype")
    assert list(s._jar) == [] and op.sent[1].get_header("Cookie") is None


def test_an_http_base_is_still_sent_over_https(make):
    s, op, _ = make(base="http://plan1.example")
    s.request("GET", "/aif/ui/model/applications/servicetype")
    assert op.sent[0].full_url == "https://plan1.example/aif/ui/model/applications/servicetype"


@pytest.mark.parametrize("timeout", [None, 0, -1, 601, "120", True, float("nan")])
def test_a_timeout_is_a_bounded_number_of_seconds(make, timeout):
    s, op, tok = make()
    with pytest.raises(epmhttp.Refused, match="timeout"):
        s.request("GET", "/aif/ui/model/applications/servicetype", timeout=timeout)
    assert op.sent == [] and tok.asked == []


def test_the_timeout_reaches_the_opener(make):
    s, op, _ = make()
    s.request("GET", "/aif/ui/model/applications/servicetype")
    s.request("POST", GET_DATA, body="{}", timeout=7)
    assert op.timeouts == [120.0, 7]


def test_an_answer_over_the_size_limit_stops_with_a_clear_error(make, monkeypatch):
    monkeypatch.setattr(epmhttp, "MAX_ANSWER_BYTES", 10)
    s, op, _ = make(answers=[(200, "x" * 11, {})])
    with pytest.raises(epmhttp.TooLarge, match="over") as e:
        s.request("POST", GET_DATA, body="{}")
    assert "limit=" not in str(e.value)


def test_an_answer_that_drips_past_its_timeout_is_a_timeout(make, monkeypatch):
    clock = iter(range(0, 10_000, 100))
    monkeypatch.setattr(epmhttp.time, "monotonic", lambda: float(next(clock)))
    s, op, _ = make()
    with pytest.raises(TimeoutError) as e:
        s.request("POST", GET_DATA, body="{}", timeout=50)
    assert "limit=" not in str(e.value)


@pytest.mark.parametrize("error", [http.client.IncompleteRead(b"{", 10), http.client.BadStatusLine("x"),
                                   http.client.LineTooLong("header line"),
                                   http.client.HTTPException("bad answer to ?limit=1 with at-SECRET")])
def test_http_clients_own_errors_are_a_lost_request(make, error):
    s, op, _ = make(tokens=["at-SECRET"], answers=[error])
    with pytest.raises(OSError) as e:
        s.request("POST", GET_DATA, body="{}")
    assert isinstance(e.value, epmhttp.RequestLost) and type(error).__name__ in str(e.value)
    assert _clean(str(e.value)) and e.value.__suppress_context__, str(e.value)


def test_the_error_scrub_reads_the_jar_under_its_lock(make):
    class Lock:
        def __init__(self):
            self.held, self.n = 0, 0

        def acquire(self, *a):
            self.held, self.n = self.held + 1, self.n + 1
            return True

        def release(self):
            self.held -= 1

        __enter__ = acquire

        def __exit__(self, *a):
            self.release()
    s, op, _ = make(tokens=["at-SECRET"], answers=[OSError("Bearer at-SECRET")])
    s._jar._cookies_lock = lock = Lock()
    seen = []
    real = s._jar.__iter__
    s._jar.__class__ = type("Jar", (type(s._jar),), {"__iter__": lambda self: (seen.append(lock.held), real())[1]})
    with pytest.raises(OSError):
        s.request("POST", GET_DATA, body="{}")
    assert seen and all(h == 1 for h in seen)                                  # every walk under the lock


def test_two_requests_at_once_share_one_consistent_jar(make):
    import threading
    gate = threading.Barrier(2, timeout=5)

    class Both(_Opener):
        def open(self, req, timeout):
            if len(self.sent) < 2:
                gate.wait()
            return super().open(req, timeout)
    s, _, _ = make()
    s._opener = op = Both([(200, "{}", {"Set-Cookie": "A=1; Path=/"}), (200, "{}", {"Set-Cookie": "B=2; Path=/"}),
                           (200, "{}", {})])
    errors = []

    def go():
        try:
            s.request("GET", "/aif/ui/model/applications/servicetype")
        except Exception as e:   # noqa: BLE001
            errors.append(e)
    threads = [threading.Thread(target=go) for _ in range(2)]
    for th in threads:
        th.start()
    for th in threads:
        th.join(5)
    s.request("GET", "/aif/ui/model/applications/servicetype")
    assert errors == [] and sorted((op.sent[2].get_header("Cookie") or "").split("; ")) == ["A=1", "B=2"]


# ---------------------------------------------------------------- re-review 2: the drill query, the drip, the jar
# EPM's drill query as the live probe of 2026-11-03 saw it (tests/test_drill_derive.py)
REAL_QUERY = ("SELECT ENTITY, AMOUNTX FROM TDATASEG abdv WHERE ( tpovpartition.partitionkey in ( ?,? ) ) AND"
              " abdv.catKey IN ( ? ) AND abdv.periodKey IN ( ? ) AND Entityx IN (?,?,? ) AND Accountx IN (?,? )"
              " AND Ud1x = ?")


@pytest.mark.parametrize("query", [
    REAL_QUERY, "SELECT 1", "SELECT 1 WHERE Entityx = ? AND x = ?",
    "  \n\t-- EPM's drill\n /* the cell */ select ENTITY FROM TDATASEG",          # leading space and comments
    "WITH t AS (SELECT 1 FROM dual) SELECT * FROM t",
    "SELECT 'a; DELETE FROM t' FROM dual", "SELECT 'it''s; DROP TABLE t' FROM dual",   # inside a string
    'SELECT "DELETE", "x;y" FROM dual',                                               # quoted identifiers
    "SELECT UPDATE_DATE, t.delete_flag, CREATED_BY FROM t",                           # not standalone words
    "SELECT 1 FROM t -- ; DELETE FROM t", "SELECT /*+ INDEX(t) */ 1 FROM t /* ; DROP */",   # inside a comment
])
def test_a_drill_query_that_only_reads_is_sent(make, query):
    s, op, _ = make()
    s.request("POST", GET_DATA, body=json.dumps({"QUERY": query, "PARAMS": [], "PARAMSTYPE": []}))
    assert len(op.sent) == 1


@pytest.mark.parametrize("query", [
    "DELETE FROM TDATASEG", "  \n UPDATE t SET x = 1", "SELECT 1; DELETE FROM TDATASEG", "SELECT 1;",
    "select 1 from t where x in (select 1) ; drop table t", "SELECT 1 /* x */ DELETE FROM t",
    "SELECT 1/**/;/**/DELETE FROM t", "SELECT 1 -- x\n DELETE FROM t", "SELECT 1 -- x\r DELETE FROM t",
    "/* SELECT */ DELETE FROM t", "-- SELECT\nDELETE FROM t", "sElEcT 1 FrOm t WhErE x = 1 Or DeLeTe",
    "WITH d AS (DELETE FROM t RETURNING *) SELECT * FROM d", "SELECT 'x' FROM t; GRANT DBA TO u",
    "SELECT * FROM t FOR UPDATE", "MERGE INTO t USING s ON (1=1)", "SELECT 1 FROM t WHERE TRUNCATE",
    "EXECUTE IMMEDIATE 'DROP TABLE t'", "EXEC p", "CALL p()", "BEGIN NULL; END;", "DECLARE x NUMBER",
    "SELECT ALTER", "SELECT CREATE", "SELECT REVOKE", "SELECT COMMIT", "SELECT INSERT",
    "SELECT 'unterminated", 'SELECT "unterminated', "SELECT 1 /* unterminated", "", "   ", "-- only a comment",
    "SELECTX 1", "WITHOUT 1", "(SELECT 1)",
])
def test_a_drill_query_that_could_write_is_refused_before_a_token(make, caplog, query):
    caplog.set_level(logging.DEBUG, logger="smartass.epm.epmhttp")
    s, op, tok = make()
    with pytest.raises(epmhttp.Refused) as e:
        s.request("POST", GET_DATA, body=json.dumps({"QUERY": query + "\n-- Q-SECRET", "PARAMS": []}))
    assert str(e.value) == "refused: the drill query is not one read-only SELECT"
    assert "Q-SECRET" not in caplog.text and op.sent == [] and tok.asked == []


@pytest.mark.parametrize("late, drip", [(0.0, 0.2), (0.8, None)])
def test_a_slow_answer_on_a_real_socket_stops_at_the_deadline(late, drip):
    """A local socket pair (no network), timeout 1 s: the answer's 40 bytes come one per 0.2 s, or its head comes at
    0.8 s with one byte and then nothing. Either way it ends near 1 s, not a socket timeout per byte later."""
    import socket
    import threading
    import time
    a, b = socket.socketpair()
    stop = threading.Event()

    def wire():
        try:
            if stop.wait(late):
                return
            b.sendall(b"HTTP/1.1 200 OK\r\nContent-Length: 40\r\n\r\nx")
            while drip and not stop.wait(drip):
                b.sendall(b"x")
        except OSError:
            pass
    th = threading.Thread(target=wire, daemon=True)
    th.start()

    class Wire:
        def open(self, req, timeout):
            a.settimeout(timeout)                       # as urllib does: per receive, not per answer
            r = http.client.HTTPResponse(a)
            r.begin()
            return r
    s = epmhttp.EpmSession(POD, tokens=_Tokens("at-1"), opener=Wire())
    t0 = time.monotonic()
    try:
        with pytest.raises(TimeoutError):
            s.request("GET", "/aif/ui/model/applications/servicetype", timeout=1)
        assert time.monotonic() - t0 < 1.5
    finally:
        stop.set()
        th.join(2)
        a.close()
        b.close()


@pytest.mark.parametrize("answer", [b"Content-Length: 2\r\n\r\n{}",
                                    b"Transfer-Encoding: chunked\r\n\r\n2\r\n{}\r\n0\r\n\r\n"])
def test_a_whole_answer_on_a_real_socket_is_read_after_its_socket_closes(answer):
    """Live 2026-11-21 (Task 7): every token-route request failed with WinError 10038. urllib closes the connection's
    socket once the answer's head is in (the answer's file keeps it open); the answer closes it for good when its
    body is read to the end, so the socket's timeout cannot be set for one more receive."""
    import socket
    a, b = socket.socketpair()
    b.sendall(b"HTTP/1.1 200 OK\r\n" + answer)

    class Wire:
        def open(self, req, timeout):
            a.settimeout(timeout)
            r = http.client.HTTPResponse(a)
            r.begin()
            a.close()                                   # as urllib's do_open does with h.sock
            return r
    s = epmhttp.EpmSession(POD, tokens=_Tokens("at-1"), opener=Wire())
    try:
        r = s.request("GET", "/aif/ui/model/applications/servicetype", timeout=5)
        assert (r.status, r.text) == (200, "{}")
    finally:
        b.close()


def test_a_401_whose_renewal_fails_still_drops_the_cookies(make):
    s, op, _ = make(tokens=["at-1", "at-1", oauth.NotSignedIn("ended")],
                    answers=[(200, "{}", {"Set-Cookie": "JSESSIONID=old; Path=/"}), (401, "", {})])
    s.request("GET", "/aif/ui/model/applications/servicetype")
    with pytest.raises(oauth.NotSignedIn):
        s.request("GET", "/aif/ui/model/applications/servicetype")
    assert list(s._jar) == []


@pytest.mark.parametrize("error", [OSError("Bearer at-SECRET"), http.client.HTTPException("at ?limit=1 at-SECRET")])
def test_a_scrubbed_error_keeps_no_link_to_the_original(make, error):
    s, op, _ = make(tokens=["at-SECRET"], answers=[error])
    with pytest.raises(OSError) as e:
        s.request("POST", GET_DATA, body="{}")
    assert e.value is not error and e.value.__context__ is None and e.value.__cause__ is None
    assert _clean(str(e.value))

def test_a_pod_without_a_scope_is_refused_with_the_add_the_scope_step(make):
    from smartass.core import serverkit
    s, op, tok = make(base="https://other.example")
    with pytest.raises(epmhttp.NoScope) as e:
        s.request("GET", "/aif/ui/model/applications/servicetype")
    assert isinstance(e.value, epmhttp.Refused) and op.sent == [] and tok.asked == []
    msg = serverkit.explain("fin_drill", e.value)
    assert "Retry once" not in msg and msg.count("Next step:") == 1
    assert "[oauth.scopes]" in msg.split("Next step:")[1] and "smartass login" in msg

# ---------------------------------------------------------------- Oracle's q-quoting in the drill query (final review M2)
@pytest.mark.parametrize("query", [
    "SELECT q'[;DELETE FROM t]' FROM dual", "SELECT q'{ ; DROP TABLE t }' FROM dual", "SELECT q'(x; UPDATE t)' FROM dual",
    "SELECT q'<a; GRANT b>' FROM dual", "SELECT q'!it's; DELETE!' FROM dual", "SELECT q'#--#', q'|/*|' FROM dual",
    "SELECT nq'[; DELETE]' FROM dual", "SELECT Q'[;]', NQ'{;}', Nq'(;)', nQ'<;>' FROM dual",
    "SELECT q'[a]b; DROP]' FROM dual", "SELECT x||q'[;]' FROM dual", "SELECT (q'[;]') FROM dual",
    "SELECT 'Høyhastighet; DELETE' FROM dual", "SELECT q'[Høy;]' FROM dual",       # letters beyond ASCII in a literal
    "SELECT q FROM t WHERE q = 'x'", "SELECT nq, q FROM t",                       # q and nq as plain names
])
def test_an_oracle_q_quoted_literal_with_a_keyword_or_a_semicolon_inside_is_sent(make, query):
    s, op, _ = make()
    s.request("POST", GET_DATA, body=json.dumps({"QUERY": query, "PARAMS": [], "PARAMSTYPE": []}))
    assert len(op.sent) == 1


@pytest.mark.parametrize("query", [
    "SELECT q'[']' ; DELETE FROM t --' FROM dual",                                 # the final review's probe
    "SELECT q'[x]' ; DELETE FROM t", "SELECT q'[x]' FROM t WHERE DELETE", "SELECT nq'{x}'; DROP TABLE t",
    "SELECT Q'(x)' FROM t; GRANT DBA TO u", "SELECT q'<x>' FROM t /* */ ; COMMIT",
    "SELECT q'[x", "SELECT q'[x]", "SELECT q'", "SELECT nq'{x}", "SELECT q'!x! FROM t",   # no end
    "SELECT q' x ' FROM t", "SELECT q'\tx\t' FROM t", "SELECT q'\nx\n' FROM t",          # whitespace as delimiter
    "SELECT q''x'' FROM t", "SELECT q'éxé' FROM t",                                       # a quote or non-ASCII one
    "SELECT éq'['; DELETE FROM t; --]' FROM dual", "SELECT 1q'['; DELETE FROM t; --]' FROM dual",
    "SELECT 'a'q'['; DELETE FROM t; --]' FROM dual", 'SELECT "a"q\'[\'; DELETE FROM t; --]\' FROM dual',
    "SELECT t.q'['; DELETE FROM t; --]' FROM dual",
    "SELECT\u00a0q'['; DELETE FROM t; --]' FROM dual", "SELECT 1\u2028; DELETE FROM t",   # not ASCII outside a literal
    "SELECT 1 \x00; DELETE FROM t", "SELECT ;x",
])
def test_a_q_quote_that_ends_early_or_cannot_be_read_with_certainty_is_refused(make, caplog, query):
    caplog.set_level(logging.DEBUG, logger="smartass.epm.epmhttp")
    s, op, tok = make()
    with pytest.raises(epmhttp.Refused) as e:
        s.request("POST", GET_DATA, body=json.dumps({"QUERY": query, "PARAMS": []}))
    assert str(e.value) == "refused: the drill query is not one read-only SELECT"
    assert op.sent == [] and tok.asked == []


def test_the_repos_real_drill_query_still_passes():
    assert epmhttp._read_only_sql(REAL_QUERY) and epmhttp._sql_tokens(REAL_QUERY)[:4] == ["SELECT", "ENTITY", ",", "AMOUNTX"]

def test_the_provider_route_sends_the_smart_view_rule_and_nothing_else_does():
    """POST /HyperionPlanning/SmartView is the grid route's (svhttp.BearerTransport, through its pods' EpmSessions);
    the drill (drill_token) is the only other module that imports epmhttp, and it never names that path."""
    import ast
    import warnings
    from pathlib import Path
    root = Path(epmhttp.__file__).parents[1]
    importers = {}
    for p in [*root.rglob("*.py"), *(root.parents[1] / "scripts").rglob("*.py")]:
        with warnings.catch_warnings():                       # a script's own escape sequences are not this test's
            warnings.simplefilter("ignore", SyntaxWarning)
            text = p.read_text(encoding="utf-8")
            tree = ast.parse(text)
        for node in ast.walk(tree):
            names = [a.name for a in getattr(node, "names", ())] if isinstance(node, (ast.Import, ast.ImportFrom)) else []
            module = getattr(node, "module", None) or ""
            if any(n.rsplit(".", 1)[-1] == "epmhttp" for n in [*names, module]):
                importers[p.name] = text
    assert set(importers) == {"drill_token.py", "svhttp.py"}
    assert [n for n, text in importers.items() if "SmartView" in text or "SMART_VIEW" in text] == ["svhttp.py"]
    assert "epmhttp.SMART_VIEW" in importers["svhttp.py"] and epmhttp.ALLOWED[epmhttp.SMART_VIEW].method == "POST"


def test_a_smart_view_body_has_its_own_size_cap(make):
    """A provider request is about 4 bytes a grid cell: a grid past 1 MB is sent, as the provider route always did;
    a body over MAX_SMART_VIEW_BODY_BYTES is not, and a getData body keeps MAX_BODY_BYTES."""
    from smartass.epm import svxml
    s, op, tok = make(answers=[(200, "<res_OpenApplication/>", {})])
    big = svxml.open_application("s" * (2 << 20), "FinMax")
    s.request("POST", "/HyperionPlanning/SmartView", body=big)
    assert op.sent[0].data == big.encode()
    pad = "s" * (epmhttp.MAX_SMART_VIEW_BODY_BYTES + 1 - len(svxml.open_application("", "FinMax")))
    with pytest.raises(epmhttp.Refused, match=f"over {epmhttp.MAX_SMART_VIEW_BODY_BYTES} bytes"):
        s.request("POST", "/HyperionPlanning/SmartView", body=svxml.open_application(pad, "FinMax"))
    with pytest.raises(epmhttp.Refused, match=f"over {epmhttp.MAX_BODY_BYTES} bytes"):
        s.request("POST", GET_DATA, body='{"QUERY": "' + "x" * epmhttp.MAX_BODY_BYTES + '"}')
    assert len(op.sent) == 1 and len(tok.asked) == 1


def test_a_caller_may_name_more_statuses_that_take_one_new_token(make):
    """renew_on: the grid route's gate answers (a 3xx or a 403 as well as a 401) take one new token and one more
    try, as a 401 does - its cookies go too; by default a 3xx is the answer."""
    from smartass.epm import svxml
    s, op, tok = make(tokens=["at-1", "at-2"], answers=[(302, "", {"Set-Cookie": "JSESSIONID=old; Path=/"}),
                                                       (200, "<res_ConnectToProvider/>", {})])
    r = s.request("POST", "/HyperionPlanning/SmartView", body=svxml.connect(), renew_on=(401, 302, 403))
    assert r.status == 200 and [x.get_header("Authorization") for x in op.sent] == ["Bearer at-1", "Bearer at-2"]
    assert op.sent[1].get_header("Cookie") is None and tok.asked[1] == (POD, "at-1")
    s2, op2, _ = make(tokens=["at-1", "at-2"], answers=[(302, "", {})])
    assert s2.request("POST", "/HyperionPlanning/SmartView", body=svxml.connect()).status == 302
    assert len(op2.sent) == 1
