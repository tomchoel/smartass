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
    fake_table(monkeypatch, [Proc(99, ["python", "child_server.py"], 7), Proc(7, ["cmd.exe"], 42),
                             Proc(42, ["python", "run_server.py"], 5, cwd=r"C:\work\My proj"), Proc(5, ["claude.exe"], 1)])
    assert session.info() == {"id": "42-1000", "cwd": r"C:\work\My proj", "project": "My proj",
                              "client": "5-1000"}


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


class Proc(P):
    """A process of a fake process table: its parent pid, create time, command line and working folder."""
    def __init__(self, pid, cmd, ppid, born=1000.5, cwd=None):
        super().__init__(pid, cmd)
        self.ppid_, self.born, self._cwd = ppid, born, cwd

    def create_time(self):
        return self.born

    def cwd(self):
        if self._cwd is None:
            raise psutil.AccessDenied(self.pid)
        return self._cwd


def fake_table(monkeypatch, procs, me=99):
    """psutil's view of the machine replaced by procs: Process() is me, Process(pid) a process of the table (gone:
    NoSuchProcess); session._ppids counts the snapshots of the whole table taken, and Process(pid) the processes
    looked at."""
    table = {p.pid: p for p in procs}
    seen = {"snapshots": 0, "looked at": []}

    def process(pid=None):
        pid = me if pid is None else pid
        seen["looked at"].append(pid)
        if pid not in table:
            raise psutil.NoSuchProcess(pid)
        return table[pid]

    def ppids():
        seen["snapshots"] += 1
        return {p.pid: p.ppid_ for p in procs}
    monkeypatch.setattr(session, "_info", None)
    monkeypatch.setattr(psutil, "Process", process)
    monkeypatch.setattr(session, "_ppids", ppids)
    return seen


def test_info_walks_the_parents_from_one_snapshot_of_the_process_table(monkeypatch):
    """psutil's Process.parents() snapshots every process on the machine once per ancestor (8 x 190 ms = 1.5 s on the
    first answer of every server start, 20.11.2026): the walk takes one snapshot, and looks at no process above the
    relay's parent."""
    seen = fake_table(monkeypatch, [
        Proc(99, ["python", "child_server.py"], 7), Proc(7, ["cmd.exe"], 42),
        Proc(42, ["python", "run_server.py"], 5, cwd=r"C:\work\My proj"), Proc(5, ["claude.exe"], 3),
        Proc(3, ["explorer.exe"], 1)])
    assert session.info() == {"id": "42-1000", "cwd": r"C:\work\My proj", "project": "My proj", "client": "5-1000"}
    assert seen["snapshots"] == 1
    assert 3 not in seen["looked at"] and 1 not in seen["looked at"]
    session.info()
    assert seen["snapshots"] == 1                                   # kept for the process


def test_a_parent_pid_reused_by_a_younger_process_ends_the_walk(monkeypatch, tmp_path):
    """As psutil's parent(): a process created after its child is not its parent (the pid was reused)."""
    monkeypatch.chdir(tmp_path)
    fake_table(monkeypatch, [
        Proc(99, ["python", "child_server.py"], 7, born=2000.0, cwd=str(tmp_path / "Here")),
        Proc(7, ["python", "run_server.py"], 5, born=2500.0, cwd=r"C:\elsewhere")])
    assert session.info() == {"id": "99-2000", "cwd": str(tmp_path / "Here"), "project": "Here", "client": ""}


def test_without_a_relay_the_server_itself_is_the_session(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    seen = fake_table(monkeypatch, [Proc(99, ["python", "child_server.py"], 7), Proc(7, ["cmd.exe"], 0),
                                    Proc(0, ["System Idle Process"], 0)])
    assert session.info() == {"id": "99-1000", "cwd": str(tmp_path), "project": tmp_path.name, "client": ""}
    assert seen["snapshots"] == 1                                   # pid 0, its own parent, ends the walk


def test_a_relay_whose_parent_is_gone_has_no_client(monkeypatch):
    fake_table(monkeypatch, [Proc(99, ["python", "child_server.py"], 42),
                             Proc(42, ["python", "run_server.py"], 5, cwd=r"C:\p")])
    assert session.info()["client"] == ""
