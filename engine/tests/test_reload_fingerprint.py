"""run_server's reload fingerprint covers every code file, including subpackages (core/, xl/ ...)."""
import importlib.util
from pathlib import Path


def _relay_module():
    spec = importlib.util.spec_from_file_location("run_server", Path(__file__).parents[1] / "run_server.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_a_change_in_a_subpackage_changes_the_fingerprint(tmp_path, monkeypatch):
    rs = _relay_module()
    (tmp_path / "smartass" / "xl").mkdir(parents=True)
    (tmp_path / "smartass" / "top.py").write_text("x = 1\n", encoding="utf-8")
    sub = tmp_path / "smartass" / "xl" / "jobs.py"
    sub.write_text("y = 1\n", encoding="utf-8")
    monkeypatch.setattr(rs, "HERE", tmp_path)
    monkeypatch.setattr(rs, "CHILD", tmp_path / "child_server.py")
    before = rs.fingerprint()
    assert any("xl" in name for name, _, _ in before)
    sub.write_text("y = 22\n", encoding="utf-8")
    assert rs.fingerprint() != before
