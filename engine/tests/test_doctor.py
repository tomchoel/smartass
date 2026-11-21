import json
import zipfile

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


def test_vba_access_check_names_the_trust_center_fix():
    key = r"Software\Microsoft\Office\16.0\Excel\Security"
    seen = []

    def reader(values):
        def read(path, name):
            seen.append((path, name))
            return values.get((path, name))
        return read
    ok = doctor._vba_access(reader({(key, "AccessVBOM"): 1}))
    assert ok.status == "OK" and ok.name == "vba_access"
    assert (key, "AccessVBOM") in seen
    for values in ({(key, "AccessVBOM"): 0}, {}):
        off = doctor._vba_access(reader(values))
        assert off.status == "FAIL"
        assert "Trust Center" in off.fix and "Trust access to the VBA project object model" in off.fix
    policy = r"Software\Policies\Microsoft\Office\16.0\Excel\Security"
    assert doctor._vba_access(reader({(policy, "AccessVBOM"): 0, (key, "AccessVBOM"): 1})).status == "FAIL"


def test_vba_access_is_a_doctor_check():
    assert doctor._vba_access in doctor.CHECKS


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
