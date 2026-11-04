from smartview_mcp import config as v1config
from smartview_mcp import session as v1session

from smartass import legacy


def test_configure_v2_moves_files_and_session(monkeypatch, tmp_path):
    monkeypatch.setattr(v1config, "ANALYSER_DIR", tmp_path / "analyser")
    monkeypatch.setattr(v1session, "_STATE", tmp_path / "sessions.json")
    monkeypatch.setattr(v1session, "_info", None)
    monkeypatch.setattr(legacy, "_configured", False)
    legacy.configure_v2()
    assert v1config.ANALYSER_DIR == tmp_path / "analyser" / "v2"
    assert v1session._STATE.name == "sessions_v2.json"
    assert v1session.info()["id"].startswith("v2-")
    legacy.configure_v2()                                   # idempotent
    assert v1config.ANALYSER_DIR == tmp_path / "analyser" / "v2"
