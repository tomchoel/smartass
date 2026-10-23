import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent / "live"))
import helpers  # noqa: E402


@pytest.fixture
def cookies(tmp_path, monkeypatch):
    d = tmp_path / "Network"
    d.mkdir()
    (d / "Cookies").write_text("fresh")
    monkeypatch.setattr(helpers, "_cookie_dirs", lambda: [d])
    monkeypatch.setattr(helpers, "kill_engine_edge", lambda: [])
    monkeypatch.setattr(helpers.time, "sleep", lambda s: None)
    return d


def test_expired_sso_never_overwrites_a_leftover_backup(cookies, monkeypatch):
    (cookies / "Cookies.matrix-bak").write_text("real sign-in")
    monkeypatch.setattr(helpers, "restore_matrix_bak", lambda: [])        # the restore failed (file locked)
    with pytest.raises(helpers.Limit):
        with helpers.expired_sso():
            pass
    assert (cookies / "Cookies.matrix-bak").read_text() == "real sign-in"
    assert (cookies / "Cookies").read_text() == "fresh"


def test_expired_sso_restores_a_leftover_backup_first(cookies):
    (cookies / "Cookies.matrix-bak").write_text("real sign-in")
    with helpers.expired_sso() as moved:
        assert moved == [str(cookies / "Cookies")]
        assert (cookies / "Cookies.matrix-bak").read_text() == "real sign-in"
        assert not (cookies / "Cookies").exists()
    assert (cookies / "Cookies").read_text() == "real sign-in"
    assert not list(cookies.glob("*.matrix-bak"))
