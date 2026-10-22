import logging

from smartview_mcp import logsetup


def test_redact_hides_tokens_and_secrets():
    logsetup.add_secret("s3cr3t-value-123")
    text = ("GET /x?sso_token=abcDEF123456789xyz&a=1 Authorization: Bearer eyJhbGciOiJIUzI1NiJ9.e30.sig "  # allow-secret
            'secret = "zzzzzzzzzzzzzzzzzzzz" and s3cr3t-value-123')  # allow-secret
    out = logsetup.redact(text)
    assert "abcDEF" not in out and "eyJhbG" not in out and "zzzz" not in out and "s3cr3t" not in out
    assert "sso_token=***" in out


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
