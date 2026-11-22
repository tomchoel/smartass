import json
import zipfile
from pathlib import Path

import pytest

from smartass.core import config
from smartass.ops import doctor


def test_overall_and_exceptions_become_fail(monkeypatch):
    def ok():
        return doctor.Check("a", "OK", "fine")

    def boom():
        raise RuntimeError("kaput")
    monkeypatch.setattr(doctor, "CHECKS", (ok, boom))
    rs = doctor.run()
    assert [c.status for c in rs] == ["OK", "FAIL"] and "kaput" in rs[1].detail and rs[1].fix
    assert doctor.overall(rs) == "FAIL"
    assert doctor.overall([doctor.Check("a", "OK", ""), doctor.Check("b", "WARN", "")]) == "WARN"


def test_config_check(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "CONFIG_ERROR", "config.toml: Expected ']'")
    assert doctor._config().status == "FAIL"
    monkeypatch.setattr(config, "CONFIG_ERROR", None)
    monkeypatch.setattr(config, "USER_CONFIG", tmp_path / "config.toml")
    assert doctor._config().status == "WARN"
    (tmp_path / "config.toml").write_text("", encoding="utf-8")
    assert doctor._config().status == "OK"


def test_output_check_writes_and_cleans_up(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "ANALYSER_DIR", tmp_path / "a")
    assert doctor._output().status == "OK" and list((tmp_path / "a").iterdir()) == []


def test_bundle_has_doctor_versions_logs_and_redacted_config(monkeypatch, tmp_path):
    (tmp_path / "logs").mkdir()
    (tmp_path / "logs" / "server.log").write_text("x sso_token=abcdefghijklmnopqrstu\n", encoding="utf-8")  # allow-secret
    (tmp_path / "config.toml").write_text('[http]\nsecret = "verysecretvalue123456"\n', encoding="utf-8")  # allow-secret
    monkeypatch.setattr(config, "HOME", tmp_path)
    monkeypatch.setattr(config, "USER_CONFIG", tmp_path / "config.toml")
    monkeypatch.setattr(doctor, "CHECKS", (lambda: doctor.Check("a", "OK", "fine"),))
    monkeypatch.setattr(doctor, "LEFTOVERS", ())            # the bundle runs them too (the config's [http] is one)
    path = doctor.bundle()
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        assert {"doctor.json", "versions.json", "config.toml", "logs/server.log"} <= names
        assert "verysecret" not in z.read("config.toml").decode()
        assert "abcdefghij" not in z.read("logs/server.log").decode()
        assert json.loads(z.read("doctor.json"))["status"] == "OK"




def test_smartview_startup_check_warns_when_the_login_at_startup_or_diagnostics_is_back():
    def reader(values):
        return lambda path, name: values.get(name)
    ok = doctor._smartview_startup(reader({"CheckOnStartup": 0, "Profile": "0"}))
    assert ok.status == "OK" and ok.name == "smartview_startup"
    on = doctor._smartview_startup(reader({"CheckOnStartup": 1, "Profile": "0"}))
    assert on.status == "WARN" and "Check for updates when Microsoft Office starts" in on.fix
    diag = doctor._smartview_startup(reader({"CheckOnStartup": 0, "Profile": "1"}))
    assert diag.status == "WARN" and "diagnostics" in diag.detail.lower()
    assert doctor._smartview_startup(reader({})).status == "OK"      # not set: Smart View's own defaults


def test_the_checks_of_excel_as_engine_smart_view_and_vba_access_are_gone():
    """Task 6: what a PC needs now is Excel for answers, the sign-in and the atlas - not Smart View, not VBA access."""
    names = [f.__name__ for f in doctor.CHECKS + doctor.LEFTOVERS]
    assert "_vba_access" not in names and "_smartview" not in names
    assert {"_excel", "_oauth", "_atlas", "_smartview_startup", "_leftovers"} <= set(names)


def test_the_doctors_checks_are_the_designs_list():
    """Light-fast phase 4 Task 3: what `smartass doctor` checks, in order; fin_status runs all but the leftovers."""
    names = [f.__name__.lstrip("_") for f in doctor.CHECKS + doctor.LEFTOVERS]
    assert names == ["config", "excel", "smartview_startup", "epm", "oauth", "atlas", "knowledge", "output",
                     "leftovers"]


def _home(monkeypatch, tmp_path):
    from smartass.core import config
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(config, "LEGACY_HOME", None)
    return tmp_path


def test_nothing_left_behind_is_ok(monkeypatch, tmp_path):
    _home(monkeypatch, tmp_path)
    assert doctor._leftovers().status == "OK" and doctor.fix_leftovers(say=lambda s: None) == []


def test_a_crashed_runs_smart_view_options_are_reported_with_their_values_and_how_to_set_them(monkeypatch, tmp_path):
    """An older release's run that crashed with Smart View's options changed: nothing can restore them now (the hidden
    Excel is gone), so the doctor says what they were and where to set them by hand."""
    home = _home(monkeypatch, tmp_path)
    (home / "smartview_options_original.json").write_text(json.dumps({"104": 2, "109": True, "110": False}))
    c = doctor._leftovers()
    assert c.status == "WARN" and "crashed with Smart View's options changed" in c.detail
    assert "which messages Smart View shows" in c.detail and "= 2" in c.detail and "= True" in c.detail
    assert "Smart View ribbon > Options" in c.fix and "smartass doctor --fix" in c.fix


def test_doctor_fix_says_each_file_before_it_removes_it(monkeypatch, tmp_path):
    home = _home(monkeypatch, tmp_path)
    opts = home / "smartview_options_original.json"
    opts.write_text(json.dumps({"104": 2}))
    bridge = home / "svbridge_0123456789ab.xlam"
    bridge.write_bytes(b"")
    said = []
    removed = doctor.fix_leftovers(say=said.append)
    assert sorted(removed) == sorted([opts, bridge]) and not opts.exists() and not bridge.exists()
    assert said[0].startswith("Smart View options an older smartass left changed") and "= 2" in said[0]
    assert said.index(f"removing {opts}") == 1 and f"removing the unused VBA bridge add-in {bridge}" in said
    assert doctor._leftovers().status == "OK"


def test_a_bridge_still_loaded_in_an_excel_is_left_and_said(monkeypatch, tmp_path):
    home = _home(monkeypatch, tmp_path)
    bridge = home / "svbridge_0123456789ab.xlam"
    bridge.write_bytes(b"")
    real_unlink = Path.unlink

    def locked(self, *a, **k):
        if self == bridge:
            raise PermissionError(13, "in use")
        return real_unlink(self, *a, **k)
    monkeypatch.setattr(Path, "unlink", locked)
    said = []
    assert doctor.fix_leftovers(say=said.append) == [] and bridge.exists()
    assert any("could not remove" in s and "close that Excel" in s for s in said)


# Smart View's stored options as the owner's PC holds them (22.11): a STRING value CAOptionsXML, about fifty options.
def _stored(message="2", home_panel="1", comment="0", extra="") -> str:
    return ('<?xml version="1.0" encoding="UTF-8"?><options>'
            '<option key="bAdjustColumnWidth" value="1"/>'
            f'<option key="messageDisplay" value="{message}"/>'
            f'<option key="bDisplayHomePanelOnStartup" value="{home_panel}"/>'
            f'<option key="bCommentDlgEnable" value="{comment}"/>{extra}</options>')


def _svkit_record(monkeypatch, tmp_path, stored):
    """The old svkit home with the record an older engine always kept there ({"104": 2, "109": true, "110": false},
    as on the owner's PC since October), and Smart View's stored options from a fake registry reader."""
    from smartass.core import config
    from smartass.xl import dialogs
    _home(monkeypatch, tmp_path / "home")
    (tmp_path / "home").mkdir()
    legacy = tmp_path / "svkit"
    legacy.mkdir()
    old = legacy / "smartview_options_original.json"
    old.write_text(json.dumps({"104": 2, "109": True, "110": False}))
    monkeypatch.setattr(config, "LEGACY_HOME", legacy)
    reads = []

    def read(path, name):
        reads.append((path, name))
        return stored(path, name) if callable(stored) else stored
    monkeypatch.setattr(dialogs, "read_hkcu", read)
    return old, reads


def test_the_old_svkit_homes_record_is_no_finding_while_smart_views_options_equal_it(monkeypatch, tmp_path):
    """The owner's PC, 22.11: the old svkit home's record and Smart View's stored options agree - the false alarm
    the doctor raised for ever. No warning, and doctor --fix says nothing of it."""
    old, reads = _svkit_record(monkeypatch, tmp_path, _stored())
    assert doctor._leftovers().status == "OK"
    said = []
    assert doctor.fix_leftovers(say=said.append) == [] and old.exists()
    assert not any("svkit" in s or "Smart View" in s for s in said), said
    assert reads and set(reads) == {(r"Software\Hyperion Solutions\HyperionSmartView\Options", "CAOptionsXML")}


@pytest.mark.parametrize("stored", [_stored(message="4"), _stored(home_panel="0"), _stored(comment="1")])
def test_the_old_svkit_homes_record_is_reported_when_smart_views_options_differ_but_never_changed(
        monkeypatch, tmp_path, stored):
    old, _ = _svkit_record(monkeypatch, tmp_path, stored)
    c = doctor._leftovers()
    assert c.status == "WARN" and str(old) in c.detail and "differ" in c.detail
    assert "Smart View ribbon > Options" in c.fix and "smartass doctor --fix" not in c.fix     # --fix cannot clear it
    assert f"delete {old} by hand" in c.fix
    said = []
    assert doctor.fix_leftovers(say=said.append) == [] and old.exists()
    assert any("left as it is" in s for s in said)


@pytest.mark.parametrize("stored", [None, "", 7, "not xml <", "<options/>", _stored(extra="").replace(
    '<option key="messageDisplay" value="2"/>', ""), _stored(message="two")])
def test_smart_views_options_that_cannot_be_read_or_compared_raise_no_warning(monkeypatch, tmp_path, stored):
    """No value, not XML, the option missing, a value that is not a number: nothing proves a difference."""
    _svkit_record(monkeypatch, tmp_path, stored)
    assert doctor._leftovers().status == "OK"


def test_a_registry_reader_that_fails_raises_no_warning(monkeypatch, tmp_path):
    def broken(path, name):
        raise OSError("access denied")
    _svkit_record(monkeypatch, tmp_path, broken)
    assert doctor._leftovers().status == "OK"


def test_smart_views_options_are_read_without_com(monkeypatch, tmp_path):
    import win32com.client
    _svkit_record(monkeypatch, tmp_path, _stored(message="4"))
    monkeypatch.setattr(win32com.client, "Dispatch", lambda *a, **k: pytest.fail("COM was used"))
    monkeypatch.setattr(win32com.client, "GetObject", lambda *a, **k: pytest.fail("COM was used"), raising=False)
    assert doctor._leftovers().status == "WARN"


def test_the_homes_own_record_is_a_crashed_run_whatever_smart_view_holds(monkeypatch, tmp_path):
    """In THIS home the file is only there when a run crashed with the options changed: reported as before."""
    from smartass.xl import dialogs
    home = _home(monkeypatch, tmp_path)
    (home / "smartview_options_original.json").write_text(json.dumps({"104": 2, "109": True, "110": False}))
    monkeypatch.setattr(dialogs, "read_hkcu", lambda path, name: _stored())
    c = doctor._leftovers()
    assert c.status == "WARN" and "crashed with Smart View's options changed" in c.detail


def test_cli_doctor_fix_runs_the_fix_then_the_checks(monkeypatch, capsys):
    from smartass.ops import cli
    order = []
    monkeypatch.setattr(doctor, "fix_leftovers", lambda say=print: order.append("fix"))
    monkeypatch.setattr(doctor, "run", lambda leftovers=False: order.append("run") or [])
    assert cli.main(["doctor", "--fix"]) == 0 and order == ["fix", "run"]


def test_cli_doctor_fix_json_prints_only_json_and_the_fix_lines_to_stderr(monkeypatch, capsys):
    import json

    from smartass.ops import cli
    monkeypatch.setattr(doctor, "fix_leftovers", lambda say=print: say("removing C:/old/svbridge_1.xlam"))
    monkeypatch.setattr(doctor, "run", lambda leftovers=False: [doctor.Check("config", "OK", "fine")])
    assert cli.main(["doctor", "--fix", "--json"]) == 0
    out = capsys.readouterr()
    assert json.loads(out.out)["checks"][0]["name"] == "config" and "removing" in out.err
    assert cli.main(["doctor", "--fix"]) == 0 and "removing" in capsys.readouterr().out      # without --json: stdout


def test_the_com_writers_are_gone():
    import smartass, pathlib
    root = pathlib.Path(smartass.__file__).parent
    assert not (root / "xl" / "owned.py").exists()
    src = (root / "pipeline.py").read_text(encoding="utf-8")
    assert "RENDER" not in src and "owned" not in src and "_backup(" not in src
    assert not hasattr(__import__("smartass.render.excel", fromlist=["x"]), "write_answer")

def test_nothing_in_the_engine_needs_a_browser():
    import pathlib, smartass
    root = pathlib.Path(smartass.__file__).parent
    hits = [p.name for p in root.rglob("*.py")
            if "playwright" in p.read_text(encoding="utf-8") or "drill_browser" in p.read_text(encoding="utf-8")]
    assert hits == []
    assert "playwright" not in (root.parent / "pyproject.toml").read_text(encoding="utf-8")
    assert "drill_browser" not in [c.__name__.lstrip("_") for c in doctor.CHECKS]
