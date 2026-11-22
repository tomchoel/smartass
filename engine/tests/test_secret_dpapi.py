"""The HTTP API secret lives DPAPI-protected in secret.bin; a plain-text http.secret in config.toml is migrated once."""
import pytest

from smartass.core import config


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
