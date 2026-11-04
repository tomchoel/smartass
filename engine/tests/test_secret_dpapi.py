"""The HTTP API secret lives DPAPI-protected in secret.bin; a plain-text http.secret in config.toml is migrated once."""
import pytest

from smartass.core import config
from smartass.ops import setup_user


class FakeDpapi:
    """Reversible stand-in for win32crypt (no real DPAPI in tests)."""
    calls = 0

    @classmethod
    def CryptProtectData(cls, data, desc, entropy, reserved, prompt, flags):
        cls.calls += 1
        return b"DPAPI:" + bytes(reversed(data))

    @staticmethod
    def CryptUnprotectData(blob, entropy, reserved, prompt, flags):
        assert blob.startswith(b"DPAPI:")
        return "smartass", bytes(reversed(blob[6:]))


@pytest.fixture
def dpapi(monkeypatch):
    monkeypatch.setattr(config, "_dpapi", lambda: FakeDpapi)
    return FakeDpapi


def test_new_config_has_no_plain_secret(tmp_path):
    p = tmp_path / "config.toml"
    setup_user.write_user_config(p)
    text = p.read_text(encoding="utf-8")
    assert "[http]" in text and "\nsecret =" not in text


def test_secret_step_creates_a_protected_secret_once(tmp_path, dpapi):
    cfg, sec = tmp_path / "config.toml", tmp_path / "secret.bin"
    setup_user.write_user_config(cfg)
    assert setup_user.secret_step(cfg, sec) == "secret: created (secret.bin, DPAPI)"
    blob = sec.read_bytes()
    plain = config.read_secret(sec)
    assert plain and len(plain) >= 32 and plain.encode() not in blob
    assert setup_user.secret_step(cfg, sec) == "secret: unchanged"
    assert config.read_secret(sec) == plain


def test_plain_secret_in_config_is_migrated_and_removed(tmp_path, dpapi):
    cfg, sec = tmp_path / "config.toml", tmp_path / "secret.bin"
    cfg.write_text('[paths]\nanalyser = "x"\n\n[http]\nport = 47999\nsecret = "old-plain-secret"\n', encoding="utf-8")
    assert setup_user.secret_step(cfg, sec) == "secret: migrated from config.toml (secret.bin, DPAPI)"
    assert config.read_secret(sec) == "old-plain-secret"
    text = cfg.read_text(encoding="utf-8")
    assert "old-plain-secret" not in text and "port = 47999" in text and 'analyser = "x"' in text
    assert setup_user.secret_step(cfg, sec) == "secret: unchanged"


def test_migration_touches_only_the_http_section(tmp_path, dpapi):
    cfg, sec = tmp_path / "config.toml", tmp_path / "secret.bin"
    cfg.write_text('[other]\nsecret = "keep-me"\n\n[http]\nsecret = "move-me"\n', encoding="utf-8")
    setup_user.secret_step(cfg, sec)
    assert config.read_secret(sec) == "move-me"
    assert cfg.read_text(encoding="utf-8") == '[other]\nsecret = "keep-me"\n\n[http]\n'


def test_secret_resolution_order(tmp_path, dpapi):
    sec = tmp_path / "secret.bin"
    sec.write_bytes(FakeDpapi.CryptProtectData(b"from-bin", None, None, None, None, 0))
    assert config.load_http_secret({"SMARTASS_HTTP_SECRET": "from-env"}, sec, {}) == "from-env"
    assert config.load_http_secret({}, sec, {"http": {"secret": "plain"}}) == "from-bin"
    assert config.load_http_secret({}, tmp_path / "none.bin", {"http": {"secret": "plain"}}) == "plain"
    assert config.load_http_secret({}, tmp_path / "none.bin", {}) is None


def test_unreadable_secret_file_is_none(tmp_path, monkeypatch):
    class Broken:
        @staticmethod
        def CryptUnprotectData(*a):
            raise OSError("another user's blob")
    monkeypatch.setattr(config, "_dpapi", lambda: Broken)
    sec = tmp_path / "secret.bin"
    sec.write_bytes(b"junk")
    assert config.read_secret(sec) is None


def test_real_dpapi_round_trip(tmp_path):
    sec = tmp_path / "secret.bin"
    sec.write_bytes(config.protect_secret("round-trip"))
    assert b"round-trip" not in sec.read_bytes()
    assert config.read_secret(sec) == "round-trip"
