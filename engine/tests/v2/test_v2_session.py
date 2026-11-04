# engine/tests/v2/test_v2_session.py
import json

import psutil

from smartass.core import config, session


class P:
    def __init__(self, pid, cmd):
        self.pid, self._cmd = pid, cmd

    def create_time(self):
        return 1000.5

    def cmdline(self):
        return self._cmd


def test_state_file_is_sessions_json_and_the_relay_is_run_server():
    assert session._STATE.name == "sessions.json"
    assert session._STATE.parent == config.HOME
    assert session._alive_relay(P(1, ["python.exe", r"C:\dev\smartass\engine\run_server.py"]))
    assert not session._alive_relay(P(2, ["python.exe", "run_server_v2.py"]))
    assert not session._alive_relay(P(3, ["python.exe", "child_server.py"]))


def test_info_takes_the_relay_ancestor(monkeypatch):
    class Me(P):
        def parents(self):
            return [P(7, ["cmd.exe"]), Relay(42, ["python", "run_server.py"])]

    class Relay(P):
        def cwd(self):
            return r"C:\work\My proj"
    monkeypatch.setattr(session, "_info", None)
    monkeypatch.setattr(psutil, "Process", lambda: Me(99, ["python", "child_server.py"]))
    assert session.info() == {"id": "42-1000", "cwd": r"C:\work\My proj", "project": "My proj"}


def test_two_sessions_survive_set_workbook(monkeypatch, tmp_path):
    procs = [P(11, ["python", "run_server.py"]), P(22, ["python", "run_server.py"]), P(44, ["python", "other.py"])]
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(session, "_STATE", tmp_path / "sessions.json")
    monkeypatch.setattr(psutil, "process_iter", lambda: procs)
    (tmp_path / "sessions.json").write_text(json.dumps({
        "11-1000": {"workbook": "a.xlsx"}, "22-1000": {"workbook": "b.xlsx"},
        "99-1000": {"workbook": "dead.xlsx"}, "_last:x": {"workbook": "x.xlsx"}}), encoding="utf-8")
    monkeypatch.setattr(session, "_info", {"id": "33-1000", "cwd": "c", "project": "p"})
    session.set_workbook("c.xlsx")
    saved = json.loads((tmp_path / "sessions.json").read_text(encoding="utf-8"))
    assert set(saved) == {"11-1000", "22-1000", "33-1000", "_last:x", "_last:c"}   # the dead relay 99 is forgotten
    assert not list(tmp_path.glob("*.tmp"))                                         # written atomically


def test_the_write_is_atomic(monkeypatch, tmp_path):
    """The state file is replaced in one step (os.replace of a temp file), never written in place."""
    import os
    monkeypatch.setattr(session, "_STATE", tmp_path / "sessions.json")
    monkeypatch.setattr(psutil, "process_iter", lambda: [])
    monkeypatch.setattr(session, "_info", {"id": "1-1", "cwd": "c", "project": "p"})
    seen = []
    real = os.replace

    def spy(src, dst):
        seen.append((os.path.basename(src), os.path.basename(dst)))
        return real(src, dst)
    monkeypatch.setattr(session.os, "replace", spy)
    session.set_workbook("w.xlsx")
    assert seen and seen[0][1] == "sessions.json" and seen[0][0].endswith(".tmp")
