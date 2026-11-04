import json

from smartass.core import config
from smartass.xl.host import ExcelHost


def _host(calls, current):
    h = ExcelHost.__new__(ExcelHost)                      # no Excel: only _run is needed here

    def run(fn, *args):
        calls.append((fn, *args))
        return ["OK", current[args[0]]] if fn == "B_GetOption" else None
    h._run = run
    return h


def _paths(tmp_path, monkeypatch):
    home, legacy = tmp_path / "smartass", tmp_path / "svkit"
    monkeypatch.setattr(config, "HOME", home)
    monkeypatch.setattr(config, "SVKIT", home)
    monkeypatch.setattr(config, "LEGACY_HOME", legacy)
    return home / "smartview_options_original.json", legacy / "smartview_options_original.json"


def test_originals_come_from_the_legacy_svkit_file_when_home_has_none(tmp_path, monkeypatch):
    home_file, legacy_file = _paths(tmp_path, monkeypatch)
    legacy_file.parent.mkdir()
    real = {str(k): f"user{k}" for k in config.GLOBAL_OPTIONS}
    legacy_file.write_text(json.dumps(real))
    calls = []
    h = _host(calls, {k: "svkit-applied" for k in config.GLOBAL_OPTIONS})
    h._apply_global_options()
    assert h.saved_options == {int(k): v for k, v in real.items()}
    assert json.loads(home_file.read_text()) == real
    assert json.loads(legacy_file.read_text()) == real          # the old engine still needs it
    assert not [c for c in calls if c[0] == "B_GetOption"]


def test_originals_are_read_from_excel_when_no_file_exists(tmp_path, monkeypatch):
    home_file, _ = _paths(tmp_path, monkeypatch)
    calls = []
    h = _host(calls, {k: f"cur{k}" for k in config.GLOBAL_OPTIONS})
    h._apply_global_options()
    assert h.saved_options == {k: f"cur{k}" for k in config.GLOBAL_OPTIONS}
    assert json.loads(home_file.read_text()) == {str(k): f"cur{k}" for k in config.GLOBAL_OPTIONS}
