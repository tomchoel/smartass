import logging
import os
import secrets

import pytest

from smartass.core import logsetup


def test_redact_hides_tokens_and_secrets():
    text = ("GET /x?sso_token=abcDEF123456789xyz&a=1 Authorization: Bearer eyJhbGciOiJIUzI1NiJ9.e30.sig "  # allow-secret
            'secret = "zzzzzzzzzzzzzzzzzzzz"')  # allow-secret
    out = logsetup.redact(text)
    assert "abcDEF" not in out and "eyJhbG" not in out and "zzzz" not in out
    assert "GET /x?... " in out                                   # the URL's query goes whole (drillcore.redact's way)
    assert logsetup.redact(f"drill sso_token={secrets.token_urlsafe(18)} ended") == "drill sso_token=*** ended"


def test_a_log_line_gets_the_same_scrub_as_the_providers_error_text(tmp_path):
    """One scrub (logsetup.redact) for a log line and a reply text: svxml's error text goes through it."""
    from smartass.epm import svxml
    x = secrets.token_urlsafe(24)
    text = f"Grid is Invalid. JSESSIONID = {x}; sessionId: {x} X-Auth: {x} https://h.example/p?code={x}"
    h = logsetup.configure(tmp_path / "logs" / "s.log")
    try:
        logging.getLogger("smartview.test").warning("%s", text)
        h.flush()
        logged = (tmp_path / "logs" / "s.log").read_text(encoding="utf-8")
    finally:
        logging.getLogger().removeHandler(h)
        h.close()
    assert x not in logged and x not in svxml._scrubbed(text)
    assert logsetup.redact(text) in logged and svxml._scrubbed(text) == logsetup.redact(text)


def _short() -> str:
    return secrets.token_hex(4)                   # 8 characters: under the 16 the broad name: value rule needs


# Task 6b: the shapes Task 6's review found leaking, one case each (made at run time; never a real key)
@pytest.mark.parametrize("shape, said", [
    ("login sID: '{x}' ok", "login sID: '***' ok"),                     # a %r log line: quoted value after name:
    ('got sessionId: "{x}" ok', 'got sessionId: "***" ok'),
    ("sessionId: {s} ok", "sessionId: *** ok"),                         # a short value after a NAMED secret key
    ("sID: {s} ok", "sID: *** ok"),
    ("password: {s} ok", "password: *** ok"),
    ("client_secret: {s} ok", "client_secret: *** ok"),
    ("{{'Set-Cookie': 'FOO={x}'}}", "{'Set-Cookie': '***'}"),          # a dict repr, a cookie name not known
    ("see ?code={x}&state=1 there", "see ?... there"),                  # a query after a space
    ("?code={x}", "?..."),                                              # ... or at the start
    ("next=https%3A%2F%2Fh%2Fp%3Fcode%3D{x} ok", "next=https%3A%2F%2Fh%2Fp%3F... ok"),   # a URL-encoded query
    ("sent Bearer%20{x} ok", "sent Bearer%20*** ok"),
    ("sent Bearer+{x} ok", "sent Bearer+*** ok"),
    ('<a sID=&quot;{x}&quot;>', '<a sID=&quot;***&quot;>'),             # an HTML-escaped attribute
], ids=["quoted-single", "quoted-double", "short-sessionId", "short-sID", "short-password", "short-client_secret",
        "cookie-dict", "query-after-space", "query-at-start", "query-%3F", "bearer-%20", "bearer-plus",
        "escaped-quotes"])
def test_each_shape_that_leaked_is_scrubbed(shape, said):
    x, s = secrets.token_hex(12), _short()
    text = shape.format(x=x, s=s)
    assert logsetup.redact(text) == said
    assert x not in logsetup.redact(text) and s not in logsetup.redact(text)


# Task 8: the named-key rule takes whole key names in their own case, a query after a space only with a named secret
# key in it, and the Bearer rule a token-shaped value; the shapes still found leaking (values made at run time)
@pytest.mark.parametrize("shape, said", [
    ("[('Set-Cookie', 'FOO={x}')]", "[('Set-Cookie', '***')]"),          # a header list of pairs
    ("{{'Set-Cookie': ['FOO={x}', 'BAR={x}']}}", "{'Set-Cookie': [***]}"),  # a list-valued header dict
    ("<a sID=&#34;{x}&#34;>", "<a sID=&#34;***&#34;>"),                  # quotes escaped by number
    ("<a sID=&#x27;{x}&#x27;>", "<a sID=&#x27;***&#x27;>"),                # ... and by hex (html.escape)
    ("posted device_code={x}&client_id=a", "posted device_code=***&client_id=a"),
    ('{{"device_code": "{x}", "interval": 5}}', '{"device_code": "***", "interval": 5}'),
    ("device_code: {s} ok", "device_code: *** ok"),
    ("see ?state=1&code={x} there", "see ?... there"),                  # the named key not first in the query
    ("sent Bearer {x} ok", "sent Bearer *** ok"),
    ("Password: {s} ok", "Password: *** ok"),                           # a named key in another case: still caught
    ("Client_Secret: {s} ok", "Client_Secret: *** ok"),
    ("SessionId: {s} ok", "SessionId: *** ok"),
    ("Access_Token: {s} ok", "Access_Token: *** ok"),
], ids=["cookie-pairs", "cookie-list", "escaped-34", "escaped-x27", "device_code-pair", "device_code-json",
        "device_code-short", "query-key-later", "bearer-token", "Password", "Client_Secret", "SessionId",
        "Access_Token"])
def test_each_shape_still_found_leaking_is_scrubbed(shape, said):
    x, s = secrets.token_hex(12), _short()
    text = shape.format(x=x, s=s)
    assert logsetup.redact(text) == said
    assert x not in logsetup.redact(text) and s not in logsetup.redact(text)


@pytest.mark.parametrize("text", [
    "SID: 4 (subsidiaries in the group)", "Session_id: unknown", "id_token: missing", "Oamaru: 3",
    "Bearer+Sender split", "Use ?period=Sep to filter", "Use ?period=Sep&view=2 to filter",
    "the cost bearer of the unit", "The cost bearer of ac.24074",
])
def test_ordinary_text_the_named_keys_and_the_bearer_rule_ate_passes_unchanged(text):
    assert logsetup.redact(text) == text


@pytest.mark.parametrize("text", [
    "Could not create C:\\a\\b.xlsx (sensitivity label?).",
    "atlas: NOT replaced (nothing was built (no such dimension?)); 0 built",
    "NoToken: no EPM access token for https://plan1.example: the identity domain did not answer (URLError)",
    "TokenRefused: Smart View provider at https://plan1.example: EPM refused the access token twice (HTTP 403)",
    "Your session: expired. Basic settings are kept. Why? Ask again.",
    "x-axis: months; tokens: 3; member code=ac.C60 under All Accounts",
    "Your session: expired; the token: renewed; the secret: kept. Why? Ask again. Member ac.24074 of un.SY",
    "What is 3 + 4? 7. The cookie jar: empty; key=value pairs: 3",
])
def test_ordinary_sentences_pass_the_scrub_unchanged(text):
    assert logsetup.redact(text) == text


# Task 9 (final review A, M2 and M3; B, M1): narrower on ordinary text - Basic and Authorization need a credential's
# shape - and not weaker on secrets: a named key with a space or hyphen, HTML-escaped JSON pairs, and a value with '%'
# after a cookie's name or SID / SESSION_ID (values made at run time; never a real key)
@pytest.mark.parametrize("text", [
    "Revenue from Basic subscriptions fell 4 %", "Basic Infrastructure costs", "Budget authorization: approved by CFO",
    "authorization: approved", "Authorization: pending (2026-11)", "session ID: unknown", "Session_id: unknown",
    "Basic 50%-growth-in-2026-q4 of the plan",                     # one class for Basic's shape and value: no '%'
])
def test_ordinary_text_basic_and_authorization_ate_passes_unchanged(text):
    assert logsetup.redact(text) == text


def _pct() -> str:
    return f"{secrets.token_hex(2)}%2B{secrets.token_hex(8)}"     # URL-encoded, over 16 characters, with '%' early


@pytest.mark.parametrize("shape, said", [
    ("Invalid session ID: {k} for the grid", "Invalid session ID: *** for the grid"),
    ("Session ID: {k} is invalid", "Session ID: *** is invalid"),
    ("Session-ID: {k} ok", "Session-ID: *** ok"),
    ("client secret: {s} ok", "client secret: *** ok"),
    ("Client-Secret: {s} ok", "Client-Secret: *** ok"),
    ("&#34;access_token&#34;:&#34;{k}&#34;", "&#34;access_token&#34;:&#34;***&#34;"),
    ("{{&quot;refresh_token&quot;: &quot;{k}&quot;, &quot;n&quot;: 1}}",
     "{&quot;refresh_token&quot;: &quot;***&quot;, &quot;n&quot;: 1}"),
    ("OAMAuthnCookie: {p} ok", "OAMAuthnCookie: *** ok"),
    ("ORA_OCIS_CG: {p} ok", "ORA_OCIS_CG: *** ok"),
    ("SESSION_ID: {p} ok", "SESSION_ID: *** ok"),
    ("SID: {p} ok", "SID: *** ok"),
    ("sent Basic {b} ok", "sent Basic *** ok"),
    ("Authorization: Bearer {s} ok", "Authorization: *** ok"),
    ("Authorization: SharedKey acct:{k} ok", "Authorization: *** ok"),
    ("authorization: {k} ok", "authorization: *** ok"),
], ids=["session-ID-spaced", "Session-ID-start", "Session-ID-hyphen", "client-secret-spaced", "Client-Secret-hyphen",
        "escaped-34-json", "escaped-quot-json", "OAM-percent", "ORA-percent", "SESSION_ID-percent", "SID-percent",
        "basic-base64", "authorization-bearer", "authorization-other-scheme", "authorization-bare"])
def test_each_secret_shape_of_the_final_review_is_scrubbed(shape, said):
    import base64
    k, s, p = secrets.token_hex(12), _short(), _pct()
    b = base64.b64encode(f"invented:{secrets.token_hex(2)}".encode()).decode()   # 'user:password' shape, with '='
    text = shape.format(k=k, s=s, p=p, b=b)
    assert logsetup.redact(text) == said
    assert not any(v in logsetup.redact(text) for v in (k, s, p, b))


def test_the_sign_in_reply_keeps_its_page_and_code(monkeypatch):
    """The sign-in reply is redacted (oauth.sign_in_reply): its page and code must survive the scrub."""
    from smartass.epm import oauth
    monkeypatch.setattr(oauth, "pending_logins", lambda: [{
        "pod": "https://plan1.example", "verification_uri": "https://idcs.example/ui/v1/device",
        "user_code": "QWER-TYUI", "minutes_left": 9}])
    said = oauth.sign_in_reply(oauth.NotSignedIn("x"))
    assert "Open https://idcs.example/ui/v1/device and enter the code QWER-TYUI (plan1.example, valid 9 min)" in said


def test_configure_writes_redacted_rotating_log(tmp_path):
    h = logsetup.configure(tmp_path / "logs" / "t.log", max_bytes=200, backups=2)
    try:
        lg = logging.getLogger("smartview.test")
        for i in range(20):
            lg.info("line %s sso_token=abcdefghijklmnopqrstu", i)  # allow-secret
        h.flush()
        files = sorted(p.name for p in (tmp_path / "logs").iterdir())
        assert files[0] == "t.log" and len(files) <= 3
        assert "abcdefghij" not in (tmp_path / "logs" / "t.log").read_text(encoding="utf-8")
    finally:
        logging.getLogger().removeHandler(h)
        h.close()


def _drop(h):
    logging.getLogger().removeHandler(h)
    h.close()


def test_rotation_blocked_by_other_process_keeps_records_and_backups(tmp_path):
    log = tmp_path / "s.log"
    log.write_text("", encoding="utf-8")
    (tmp_path / "s.log.1").write_text("backup one\n", encoding="utf-8")
    (tmp_path / "s.log.2").write_text("backup two\n", encoding="utf-8")
    other = open(log, "a", encoding="utf-8")         # stands in for another server process holding the log
    try:
        try:                                         # the scenario needs a handle that blocks rename (Windows)
            os.rename(log, tmp_path / "probe")
            os.rename(tmp_path / "probe", log)
            pytest.skip("an open handle does not block rename on this platform")
        except OSError:
            pass
        h = logsetup.configure(log, max_bytes=200, backups=3)
        try:
            lg = logging.getLogger("smartview.test.shared")
            for i in range(6):
                lg.info("record-%s %s", i, "x" * 60)
            h.flush()
        finally:
            _drop(h)
    finally:
        other.close()
    text = log.read_text(encoding="utf-8")
    assert all(f"record-{i} " in text for i in range(6))
    assert (tmp_path / "s.log.1").read_text(encoding="utf-8") == "backup one\n"
    assert (tmp_path / "s.log.2").read_text(encoding="utf-8") == "backup two\n"


def test_rotation_without_other_handle_still_rotates(tmp_path):
    log = tmp_path / "r.log"
    h = logsetup.configure(log, max_bytes=200, backups=10)
    try:
        lg = logging.getLogger("smartview.test.rotate")
        for i in range(6):
            lg.info("record-%s %s", i, "x" * 60)
        h.flush()
    finally:
        _drop(h)
    assert (tmp_path / "r.log.1").exists()
    assert not (tmp_path / "r.log.rotating").exists()
    every = "".join(p.read_text(encoding="utf-8") for p in tmp_path.iterdir())
    assert all(f"record-{i} " in every for i in range(6))
