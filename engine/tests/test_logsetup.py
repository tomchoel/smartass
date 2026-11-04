import logging
import os

import pytest

from smartass.core import logsetup


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
