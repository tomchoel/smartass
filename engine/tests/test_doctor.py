import json
import zipfile
from pathlib import Path

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
    names = [f.__name__ for f in doctor.CHECKS]
    assert "_vba_access" not in names and "_smartview" not in names
    assert {"_excel", "_oauth", "_atlas", "_smartview_startup", "_leftovers"} <= set(names)


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


def test_the_old_svkit_homes_copy_is_reported_but_never_changed(monkeypatch, tmp_path):
    from smartass.core import config
    _home(monkeypatch, tmp_path / "home")
    (tmp_path / "home").mkdir()
    legacy = tmp_path / "svkit"
    legacy.mkdir()
    old = legacy / "smartview_options_original.json"
    old.write_text(json.dumps({"109": True}))
    monkeypatch.setattr(config, "LEGACY_HOME", legacy)
    said = []
    assert doctor._leftovers().status == "WARN" and doctor.fix_leftovers(say=said.append) == [] and old.exists()
    assert any("left as it is" in s for s in said)


def test_cli_doctor_fix_runs_the_fix_then_the_checks(monkeypatch, capsys):
    from smartass.ops import cli
    order = []
    monkeypatch.setattr(doctor, "fix_leftovers", lambda: order.append("fix"))
    monkeypatch.setattr(doctor, "run", lambda: order.append("run") or [])
    assert cli.main(["doctor", "--fix"]) == 0 and order == ["fix", "run"]


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
