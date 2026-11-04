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


def test_two_v2_sessions_survive_set_workbook(monkeypatch, tmp_path):
    import json

    import psutil

    class P:
        def __init__(self, pid, cmd):
            self.pid, self._cmd = pid, cmd

        def create_time(self):
            return 1000.5

        def cmdline(self):
            return self._cmd

    procs = [P(11, ["python", "run_server_v2.py"]), P(22, ["python", "run_server_v2.py"]),
             P(33, ["python", "run_server.py"])]
    monkeypatch.setattr(v1config, "ANALYSER_DIR", tmp_path / "analyser")
    monkeypatch.setattr(v1config, "SVKIT", tmp_path)
    monkeypatch.setattr(v1session, "_STATE", tmp_path / "sessions_v2.json")
    monkeypatch.setattr(v1session, "_info", None)
    monkeypatch.setattr(legacy, "_configured", False)
    monkeypatch.setattr(v1session, "_alive_relay", v1session._alive_relay)     # restored after the test
    monkeypatch.setattr(v1session, "set_workbook", v1session.set_workbook)
    monkeypatch.setattr(psutil, "process_iter", lambda: procs)
    legacy.configure_v2()
    (tmp_path / "sessions_v2.json").write_text(json.dumps({
        "v2-11-1000": {"workbook": "a.xlsx"}, "v2-22-1000": {"workbook": "b.xlsx"},
        "v2-99-1000": {"workbook": "dead.xlsx"}}), encoding="utf-8")
    monkeypatch.setattr(v1session, "_info", {"id": "v2-33-1000", "cwd": "c", "project": "p"})
    v1session.set_workbook("c.xlsx")
    saved = json.loads((tmp_path / "sessions_v2.json").read_text(encoding="utf-8"))
    assert set(saved) == {"v2-11-1000", "v2-22-1000", "v2-33-1000"}        # the dead relay 99 is forgotten
