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


def test_answer_ok_fails_on_error_cell_warnings():
    import scenarios
    out = {"checks": {"c": "OK"}, "warnings": [
        "error cell 'Q6 FinMax recon'!B3 is #NoAccess: no read access or member not valid in this cube"]}
    with pytest.raises(AssertionError, match="#NoAccess"):
        scenarios._answer_ok(out, "actual_vs_plan.sy")
    scenarios._answer_ok({"checks": {"c": "OK"}, "warnings": ["could not log Q1 in 'Kilder'"]}, "ok")


def test_only_empty_excels_the_run_started_are_quit():
    before = {100, 200}
    assert helpers.leftover_excel(300, before, 0, "python.exe")           # started in the run, now empty
    assert not helpers.leftover_excel(100, before, 0, "python.exe")       # was there before the run
    assert not helpers.leftover_excel(300, before, 1, "python.exe")       # still has a workbook
    assert not helpers.leftover_excel(300, before, 0, "explorer.exe")     # the user opened it meanwhile
