"""epm/epmhttp.py: the token session's allowlist, the pod's origin, the bearer, the cookie jar, a 401 and a redirect.
The network is faked: an opener that records what it is given (or, for the redirect handler, urllib's real opener
over a fake HTTPS handler). Nothing here reaches EPM."""
import http.client
import io
import logging
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
        self.answers, self.sent = list(answers), []

    def open(self, req, timeout):
        self.sent.append(req)
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
    s.request("GET", "https://Plan1.Example:443/aif/ui/model/view?page=DRILL_THROUGH&a=%2F#frag")
    s.request("GET", "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.drilldown&FORMAT=SUI&_=1")
    assert [r.full_url for r in op.sent] == [
        "https://plan1.example/aif/ui/model/view?page=DRILL_THROUGH&a=%2F",
        "https://plan1.example/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.drilldown&FORMAT=SUI&_=1"]
    assert [r.get_method() for r in op.sent] == ["GET", "GET"] and all(r.data is None for r in op.sent)


def test_a_read_only_smart_view_request_and_a_getdata_post_are_sent(make):
    from smartass.epm import svxml
    s, op, _ = make(answers=[(200, "<res_ConnectToProvider/>", {"Content-Type": "text/xml"})])
    r = s.request("POST", "/HyperionPlanning/SmartView", headers={"Content-Type": "text/xml"}, body=svxml.connect())
    assert r == epmhttp.Response(200, "<res_ConnectToProvider/>", {"content-type": "text/xml"})
    s.request("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0", body='{"QUERY": "q"}', timeout=7)
    assert [(x.get_method(), x.data) for x in op.sent] == [("POST", svxml.connect().encode()),
                                                            ("POST", b'{"QUERY": "q"}')]


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
