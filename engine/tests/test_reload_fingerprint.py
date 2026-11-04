"""run_server reloads its child only when engine/VERSION changes or on fin_status(restart=True) - never on a .py edit."""
import importlib.util
import re
from pathlib import Path

ENGINE = Path(__file__).parents[1]


def _relay_module():
    spec = importlib.util.spec_from_file_location("run_server", ENGINE / "run_server.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _tree(tmp_path):
    (tmp_path / "smartass" / "xl").mkdir(parents=True)
    sub = tmp_path / "smartass" / "xl" / "jobs.py"
    sub.write_text("y = 1\n", encoding="utf-8")
    (tmp_path / "VERSION").write_text("0.2.0\n", encoding="utf-8")
    return sub


def test_a_py_edit_does_not_change_the_fingerprint(tmp_path, monkeypatch):
    rs = _relay_module()
    sub = _tree(tmp_path)
    monkeypatch.setattr(rs, "HERE", tmp_path)
    monkeypatch.setattr(rs, "CHILD", tmp_path / "child_server.py")
    before = rs.fingerprint()
    sub.write_text("y = 22\n", encoding="utf-8")
    (tmp_path / "child_server.py").write_text("# edited\n", encoding="utf-8")
    assert rs.fingerprint() == before


def test_a_version_bump_changes_the_fingerprint(tmp_path, monkeypatch):
    rs = _relay_module()
    _tree(tmp_path)
    monkeypatch.setattr(rs, "HERE", tmp_path)
    before = rs.fingerprint()
    (tmp_path / "VERSION").write_text("0.2.1\n", encoding="utf-8")
    assert rs.fingerprint() != before


def test_fin_status_restart_is_an_explicit_reload():
    rs = _relay_module()

    def call(name, args):
        return {"jsonrpc": "2.0", "id": 7, "method": "tools/call", "params": {"name": name, "arguments": args}}

    assert rs.wants_restart(call("fin_status", {"restart": True}))
    assert not rs.wants_restart(call("fin_status", {}))
    assert not rs.wants_restart(call("fin_status", {"restart": False}))
    assert not rs.wants_restart(call("fin_analyze", {"restart": True}))
    assert not rs.wants_restart({"jsonrpc": "2.0", "id": 1, "method": "tools/list"})


def test_relay_reloads_on_explicit_restart_even_with_the_same_version(monkeypatch):
    rs = _relay_module()
    r = rs.Relay()
    started, stopped = [], []

    class FakeChild:
        def poll(self):
            return None

    r.child, r.fp, r.init_req = FakeChild(), rs.fingerprint(), {"id": 0}
    monkeypatch.setattr(r, "start_child", lambda replay: started.append(replay))
    monkeypatch.setattr(r, "stop_child", lambda reason: stopped.append(reason))
    monkeypatch.setattr(r, "notify_client", lambda method: None)
    r.maybe_reload()
    assert not started                               # same VERSION: no reload
    r.maybe_reload(force=True)
    assert started == [True] and stopped


def test_version_file_matches_package_and_pyproject():
    import smartass
    version = (ENGINE / "VERSION").read_text(encoding="utf-8").strip()
    pyproject = re.search(r'(?m)^version\s*=\s*"([^"]+)"', (ENGINE / "pyproject.toml").read_text()).group(1)
    assert version == pyproject == smartass.__version__


def test_bump_script_exists():
    text = (ENGINE.parent / "scripts" / "bump.ps1").read_text(encoding="utf-8")
    assert "VERSION" in text and "pyproject.toml" in text
