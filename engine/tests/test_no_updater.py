"""The updater and the release folder are deleted (light-fast phase 4 Task 2): ops/update.py and `smartass update`,
config's release settings, setup's atlas copy from the release folder, the doctor's update check and the publishing
script are gone. A user's config.toml that still has an older release's settings ([release] folder / version, [http]
port / secret) is read without error; `smartass doctor` names them (never their values) as no longer used and
`doctor --fix` takes those sections out, each said first. Every secret-looking value is made at run time."""
import importlib.util
import os
import secrets
import subprocess
import sys

import pytest
from test_neutral_imports import ENGINE, _rel, _sources

from smartass.core import config
from smartass.ops import doctor, setup_user

NAMES = ["RELEASE_FOLDER", "RELEASE_PIN", "SMARTASS_RELEASE", "copy_atlas", "latest.json", "WithAtlas",
         "smartass update", "publish_release", "update.check", "update.apply", "install.ps1"]
# The setup template every PC got up to 0.7.0: a [release] section holding only a comment.
OLD_TEMPLATE = """[paths]
# analyser = 'C:\\Users\\you\\analyser'

[packs]
enabled = []

[release]
# version = "0.2.0"     # pin a version (also downwards); empty = latest
"""


def test_the_updater_is_gone():
    assert importlib.util.find_spec("smartass.ops.update") is None
    assert not (ENGINE.parent / "scripts" / "publish_release.ps1").exists()
    assert not hasattr(config, "RELEASE_FOLDER") and not hasattr(config, "RELEASE_PIN")
    assert not hasattr(setup_user, "copy_atlas") and "[release]" not in setup_user.CONFIG_TEMPLATE
    assert "update" not in [f.__name__.lstrip("_") for f in doctor.CHECKS]
    s, err = config.load_settings(config.DEFAULTS_FILE, config.DEFAULTS_FILE.with_name("absent.toml"))
    assert err is None and "release" not in s


def test_the_command_line_has_no_update(capsys):
    from smartass.ops import cli
    with pytest.raises(SystemExit) as e:
        cli.main(["update"])
    assert e.value.code == 2 and "invalid choice" in capsys.readouterr().err


def test_no_name_of_the_updater_is_left_in_the_engine():
    assert [f"{_rel(p)}: {n}" for p in _sources() for n in NAMES if n in p.read_text(encoding="utf-8")] == []


def test_setup_copies_no_atlas(monkeypatch, tmp_path):
    """setup's steps end with the PATH step: no atlas line (`smartass login` builds the atlas)."""
    monkeypatch.setattr(config, "HOME", tmp_path)
    monkeypatch.setattr(config, "USER_CONFIG", tmp_path / "config.toml")
    monkeypatch.setattr(setup_user, "USER", tmp_path / "user")
    monkeypatch.setattr(setup_user, "link_skill", lambda src, dst, backup: "linked")
    monkeypatch.setattr(setup_user, "path_step", lambda: "path: unchanged")
    out = setup_user.run()
    assert out[-1] == "path: unchanged" and not any(line.startswith("atlas") for line in out)


def test_an_old_config_with_release_and_http_settings_is_read_without_error(tmp_path):
    secret = secrets.token_urlsafe(18)
    (tmp_path / "config.toml").write_text(
        f"[release]\nfolder = 'C:\\\\share\\\\rel'\nversion = '0.2.0'\n\n[http]\nport = 8765\nsecret = '{secret}'\n",
        encoding="utf-8")
    env = {**os.environ, "SMARTASS_HOME": str(tmp_path), "SMARTASS_RELEASE": r"C:\other"}
    code = "from smartass.core import config as c; print(repr(c.CONFIG_ERROR), hasattr(c, 'RELEASE_FOLDER'))"
    out = subprocess.run([sys.executable, "-c", code], env=env, capture_output=True, text=True, check=True).stdout
    assert out.split() == ["None", "False"]


@pytest.fixture
def home(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "HOME", tmp_path)
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(config, "LEGACY_HOME", None)
    monkeypatch.setattr(config, "USER_CONFIG", tmp_path / "config.toml")
    return tmp_path


def test_doctor_names_the_settings_no_longer_used_never_their_values(home):
    secret, folder = secrets.token_urlsafe(18), r"C:\share\smartass-release"
    (home / "config.toml").write_text(f"[paths]\nanalyser = 'C:/a'\n\n[release]\nfolder = '{folder}'\n"
                                      f"version = '0.2.0'\n\n[http]\nport = 8765\nsecret = '{secret}'\n")
    c = doctor._leftovers()
    assert c.status == "WARN" and "smartass doctor --fix" in c.fix
    assert "release.folder, release.version, http.port" in c.detail and "no longer used" in c.detail
    assert "plain [http] secret" in c.detail
    assert secret not in c.detail + c.fix and folder not in c.detail + c.fix and "8765" not in c.detail


def test_the_old_templates_empty_release_section_is_no_finding(home):
    (home / "config.toml").write_text(OLD_TEMPLATE)
    assert doctor._leftovers().status == "OK"
    doctor.fix_leftovers(say=lambda s: None)
    assert (home / "config.toml").read_text() == OLD_TEMPLATE


def test_doctor_fix_takes_the_old_sections_out_whole_saying_so_first(home):
    secret = secrets.token_urlsafe(18)
    cfg = home / "config.toml"
    cfg.write_text(f"[paths]\nanalyser = 'C:/a'\n\n[release]\n# where releases came from\nfolder = 'C:/rel'\n\n"
                   f"[http]\nport = 8765\nsecret = '{secret}'\n\n[packs]\nenabled = []\n")
    said = []

    def say(s):
        said.append((s, cfg.read_text()))
    doctor.fix_leftovers(say=say)
    assert cfg.read_text() == "[paths]\nanalyser = 'C:/a'\n\n[packs]\nenabled = []\n"
    lines = [s for s, _ in said]
    assert any("release.folder, http.port" in s and "no longer used" in s for s in lines)
    assert any("plain [http] secret" in s for s in lines)
    assert all("[release]" in text for s, text in said if "no longer used" in s)       # said before it went
    assert secret not in " ".join(lines) and "8765" not in " ".join(lines)
    assert doctor._leftovers().status == "OK"


def test_a_setting_the_fix_cannot_take_out_is_said_to_be_left(home):
    cfg = home / "config.toml"
    cfg.write_text("release.folder = 'C:/rel'\n")                       # a dotted key: no [release] line to take out
    said = []
    doctor.fix_leftovers(say=said.append)
    assert cfg.read_text() == "release.folder = 'C:/rel'\n" and doctor._leftovers().status == "WARN"
    assert any("release.folder" in s and "by hand" in s for s in said)


def test_the_updaters_attempt_file_is_a_leftover_doctor_fix_removes(home):
    p = home / "update_attempt.json"
    p.write_text('{"at": 1}')
    c = doctor._leftovers()
    assert c.status == "WARN" and str(p) in c.detail and "smartass doctor --fix" in c.fix
    said = []
    assert doctor.fix_leftovers(say=said.append) == [p] and not p.exists()
    assert any(str(p) in s and s.startswith("removing") for s in said)


def test_install_is_three_commands_in_the_readme():
    """The installer is deleted (Task 2: the smaller of rewriting it and the README's three lines); the README's install
    is uv sync, smartass setup, smartass login, in that order."""
    repo = ENGINE.parent
    assert not (repo / "installer").exists()
    readme = (repo / "README.md").read_text(encoding="utf-8")
    install = readme.split("## Install", 1)[1].split("\n## ", 1)[0]
    commands = ("uv sync --frozen", "uv run --frozen smartass setup", "uv run --frozen smartass login")
    at = [install.index(c) for c in commands]
    assert at == sorted(at) and "release folder" not in install


def test_the_config_check_says_to_run_setup(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "CONFIG_ERROR", None)
    monkeypatch.setattr(config, "USER_CONFIG", tmp_path / "config.toml")
    assert doctor._config().fix == "Run `smartass setup`."
    monkeypatch.setattr(config, "CONFIG_ERROR", "config.toml: Expected ']'")
    assert "`smartass setup`" in doctor._config().fix


# ---------------------------------------------------------------- Task 3 (b): the user's config.toml is edited safely


def _fixed(home, raw: bytes) -> tuple[bytes, list[str]]:
    cfg = home / "config.toml"
    cfg.write_bytes(raw)
    said = []
    doctor.fix_leftovers(say=said.append)
    return cfg.read_bytes(), said


def _backups(home) -> list:
    return sorted(p.name for p in home.glob("config.toml.bak-*"))


def test_a_backup_is_written_and_said_before_the_config_is_changed_and_never_overwritten(home):
    import time
    raw = b"[paths]\nanalyser = 'C:/a'\n\n[release]\nfolder = 'C:/rel'\n"
    day = time.strftime("%Y-%m-%d")
    (home / f"config.toml.bak-{day}").write_bytes(b"an earlier migration's copy")
    cfg = home / "config.toml"
    cfg.write_bytes(raw)
    said = []
    doctor.fix_leftovers(say=lambda s: said.append((s, cfg.read_bytes())))
    assert cfg.read_bytes() == b"[paths]\nanalyser = 'C:/a'\n\n"
    bak = home / f"config.toml.bak-{day}-2"
    assert bak.read_bytes() == raw and (home / f"config.toml.bak-{day}").read_bytes() == b"an earlier migration's copy"
    assert [text for s, text in said if s.startswith("backing up") and str(bak) in s] == [raw]   # said first
    assert _backups(home) == [f"config.toml.bak-{day}", f"config.toml.bak-{day}-2"]


def test_crlf_lines_are_kept_crlf_and_lf_lines_lf(home):
    crlf = b"[paths]\r\nanalyser = 'C:/a'\r\n\r\n[http]\r\nport = 8765\r\n\r\n[packs]\r\nenabled = []\r\n"
    assert _fixed(home, crlf)[0] == b"[paths]\r\nanalyser = 'C:/a'\r\n\r\n[packs]\r\nenabled = []\r\n"
    lf = crlf.replace(b"\r\n", b"\n")
    assert _fixed(home, lf)[0] == b"[paths]\nanalyser = 'C:/a'\n\n[packs]\nenabled = []\n"


def test_a_config_with_a_bom_is_left_byte_for_byte(home):
    """smartass reads config.toml as UTF-8 without a BOM: such a file is the config check's FAIL already (TOML does
    not take the mark) and the fix never touches it."""
    raw = b"\xef\xbb\xbf[paths]\nanalyser = 'C:/a'\n\n[http]\nport = 8765\n"
    assert _fixed(home, raw)[0] == raw and _backups(home) == []


def test_a_last_section_without_a_trailing_newline_goes(home):
    assert _fixed(home, b"[paths]\nanalyser = 'C:/a'\n\n[http]\nport = 8765")[0] == b"[paths]\nanalyser = 'C:/a'\n\n"


def test_an_http_line_inside_a_multi_line_string_is_not_a_section(home):
    raw = b'[paths]\nnote = """\n[http]\nport = 1\n"""\n\n[http]\nport = 8765\n'
    assert _fixed(home, raw)[0] == b'[paths]\nnote = """\n[http]\nport = 1\n"""\n\n'


def test_an_array_of_tables_after_the_section_is_kept(home):
    raw = b'[http]\nport = 1\n\n[[tables]]\nname = "keep"\n'
    assert _fixed(home, raw)[0] == b'[[tables]]\nname = "keep"\n'


def test_comment_lines_just_above_the_next_section_stay_with_it(home):
    raw = b'[http]\nport = 1\n\n# my oauth settings\n[oauth]\nidcs = "https://idcs.example"\n'
    assert _fixed(home, raw)[0] == b'# my oauth settings\n[oauth]\nidcs = "https://idcs.example"\n'


def test_when_in_doubt_the_file_is_left_as_it_is_and_the_lines_are_said(home):
    """A sub-table of [http] would be left behind by taking the section out: the result is not the old config minus
    the two tables, so nothing is written and the lines to delete by hand are said."""
    raw = b"[paths]\nanalyser = 'C:/a'\n\n[http]\nport = 1\n\n[http.extra]\nx = 1\n"
    new, said = _fixed(home, raw)
    assert new == raw and _backups(home) == []
    hand = [s for s in said if "by hand" in s]
    assert hand and "lines 4-6" in hand[0] and "lines 7-8" in hand[0], said
