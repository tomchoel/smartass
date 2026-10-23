import json
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


class FakeReg:
    """CAOptionsXML in memory (None = absent); counts writes/deletes."""

    def __init__(self, value=None):
        self.value, self.writes, self.deletes = value, 0, 0

    def read(self):
        if self.value is None:
            raise FileNotFoundError
        return self.value

    def write(self, data, kind):
        self.value, self.writes = (data, kind), self.writes + 1

    def delete(self):
        self.value, self.deletes = None, self.deletes + 1


XML = ('<options><option key="messageDisplay" value="2"/><option key="bDisplayHomePanelOnStartup" value="1"/>'
       '<option key="bCommentDlgEnable" value="0"/></options>')


@pytest.fixture
def homes(tmp_path, monkeypatch):
    from smartview_mcp import config

    monkeypatch.setattr(config, "LEGACY_HOME", None)
    real, fresh = tmp_path / "real", tmp_path / "fresh"
    real.mkdir()
    return real, fresh, real / "matrix-hkcu-options.json"


def test_owner_options_put_back_and_fresh_originals_seeded_from_hkcu(homes):
    real, fresh, marker = homes
    reg = FakeReg((XML, 1))
    with helpers.owner_sv_options_kept(fresh, reg, marker, real) as snap:
        assert json.loads(marker.read_text()) == {"snap": snap, "originals": {"104": 2, "109": True, "110": False},
                                                  "real_originals_existed": False}
        assert snap == {"present": True, "data": XML, "kind": 1}
        assert json.loads((fresh / "smartview_options_original.json").read_text()) == \
            {"104": 2, "109": True, "110": False}
        reg.write(XML.replace('"2"', '"3"'), 1)                 # the fresh engine applies its values
    assert reg.value == (XML, 1) and reg.writes == 2
    assert not marker.exists()


def test_owner_options_absent_value_is_deleted_again(homes):
    real, fresh, marker = homes
    reg = FakeReg()
    with helpers.owner_sv_options_kept(fresh, reg, marker, real):
        assert json.loads((fresh / "smartview_options_original.json").read_text()) == {}
        reg.write("<engine/>", 1)
    assert reg.value is None and reg.deletes == 1 and not marker.exists()


def test_owner_options_unchanged_are_not_rewritten_and_real_originals_win(homes):
    real, fresh, marker = homes
    (real / "smartview_options_original.json").write_text('{"104": 0, "109": true, "110": true}')
    reg = FakeReg((XML, 1))                                      # the real engine's values are applied
    with pytest.raises(RuntimeError):
        with helpers.owner_sv_options_kept(fresh, reg, marker, real):
            assert json.loads((fresh / "smartview_options_original.json").read_text()) == \
                {"104": 0, "109": True, "110": True}
            raise RuntimeError("scenario failed")
    assert reg.writes == 0 and reg.deletes == 0 and not marker.exists()


def _opts(msg, home, comment, other="A"):
    return (f'<options><option key="messageDisplay" value="{msg}"/><option key="other" value="{other}"/>'
            f'<option key="bDisplayHomePanelOnStartup" value="{home}"/><option key="bCommentDlgEnable" '
            f'value="{comment}"/></options>')


TRUE_VALUES = {"104": 2, "109": True, "110": True}


_DEAD = []                       # the entered contexts of "dead" runs: never closed, so their finally never runs


def _died_in_s9(fresh, reg, marker, real):
    """S9.lookup entered, then the matrix process died (the finally never ran)."""
    cm = helpers.owner_sv_options_kept(fresh, reg, marker, real)
    cm.__enter__()
    _DEAD.append(cm)
    assert marker.exists()


def test_leftover_marker_left_to_the_real_engine_while_its_originals_file_exists(homes):
    real, fresh, marker = homes
    (real / "smartview_options_original.json").write_text(json.dumps(TRUE_VALUES))
    reg = FakeReg((_opts(3, 0, 0), 1))                          # the real engine's values are applied
    _died_in_s9(fresh, reg, marker, real)
    reg.value, reg.writes = (_opts(3, 0, 0, other="B"), 1), 0
    assert helpers.restore_leftover_hkcu_options(reg, marker, real) is False
    assert reg.writes == 0 and reg.deletes == 0 and not marker.exists()


def test_leftover_marker_sets_only_the_three_options_and_keeps_the_type(homes):
    real, fresh, marker = homes
    reg = FakeReg((_opts(2, 1, 1), 2))                          # no engine running: HKCU holds the owner's values
    _died_in_s9(fresh, reg, marker, real)
    reg.value, reg.writes = (_opts(3, 0, 0, other="changed by the user"), 2), 0   # fresh engine's values
    assert helpers.restore_leftover_hkcu_options(reg, marker, real) is True
    assert reg.value == (_opts(2, 1, 1, other="changed by the user"), 2) and reg.writes == 1
    assert not marker.exists()


def test_leftover_marker_after_the_real_engine_exited_keeps_the_true_values(homes):
    real, fresh, marker = homes
    orig = real / "smartview_options_original.json"
    orig.write_text(json.dumps(TRUE_VALUES))
    reg = FakeReg((_opts(3, 0, 0), 1))                          # snapshot = the real engine's values
    _died_in_s9(fresh, reg, marker, real)
    reg.value = (_opts(2, 1, 1), 1)                             # the real engine exits cleanly: true values back
    orig.unlink()
    helpers.restore_leftover_hkcu_options(reg, marker, real)
    assert reg.value == (_opts(2, 1, 1), 1) and not marker.exists()
    reg.value = (_opts(3, 0, 0), 1)                             # even if the fresh engine's values were still there
    marker.write_text(json.dumps({"snap": {"present": True, "data": _opts(3, 0, 0), "kind": 1},
                                  "originals": TRUE_VALUES, "real_originals_existed": True}))
    assert helpers.restore_leftover_hkcu_options(reg, marker, real) is True
    assert reg.value == (_opts(2, 1, 1), 1)
