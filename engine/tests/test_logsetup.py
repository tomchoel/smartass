import logging
import os
import secrets

import pytest

from smartass.core import logsetup


def test_redact_hides_tokens_and_secrets():
    logsetup.add_secret("s3cr3t-value-123")
    text = ("GET /x?sso_token=abcDEF123456789xyz&a=1 Authorization: Bearer eyJhbGciOiJIUzI1NiJ9.e30.sig "  # allow-secret
            'secret = "zzzzzzzzzzzzzzzzzzzz" and s3cr3t-value-123')  # allow-secret
    out = logsetup.redact(text)
    assert "abcDEF" not in out and "eyJhbG" not in out and "zzzz" not in out and "s3cr3t" not in out
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
