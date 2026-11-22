"""run_server's stop of a server that does not exit (final review B, I1): its process tree is ended, but never an Excel
in it - an Excel the server launched (xl/target._launch_excel) is the user's Excel, with the user's workbooks - nor
anything under that Excel. Fake process trees only: nothing here lists or ends a real process."""
import importlib.util
import subprocess
from pathlib import Path

import psutil

ENGINE = Path(__file__).parents[1]


def _relay_module():
    spec = importlib.util.spec_from_file_location("run_server", ENGINE / "run_server.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class Proc:
    """A process of a fake tree: its name, its direct children, and what happened to it (ended)."""

    def __init__(self, pid, name, kids=(), ended=None, gone=False):
        self.pid, self._name, self._kids, self.gone = pid, name, list(kids), gone
        self.ended = ended if ended is not None else []

    def name(self):
        if self.gone:
            raise psutil.NoSuchProcess(self.pid)
        return self._name

    def children(self, recursive=False):
        if self.gone:
            raise psutil.NoSuchProcess(self.pid)
        if recursive:                                   # psutil's flat list: what a filter on the name alone sees
            return [d for k in self._kids for d in [k, *k.children(recursive=True)]]
        return list(self._kids)

    def kill(self):
        if self.gone:
            raise psutil.NoSuchProcess(self.pid)
        self.ended.append(self.pid)


def _tree():
    ended = []

    def p(pid, name, *kids, gone=False):
        return Proc(pid, name, kids, ended, gone)
    excel = p(40, "EXCEL.EXE", p(41, "splwow64.exe"), p(42, "msedgewebview2.exe", p(43, "msedgewebview2.exe")))
    server = p(10, "python.exe", p(20, "python.exe", p(21, "conhost.exe")), excel, p(30, "node.exe"))
    return server, ended


def test_the_tree_is_ended_without_an_excel_and_without_what_runs_under_it():
    rs = _relay_module()
    server, ended = _tree()
    rs.end_tree(server)
    assert sorted(ended) == [10, 20, 21, 30] and ended[-1] == 10      # the server itself last, as before
    assert not {40, 41, 42, 43} & set(ended)


def test_an_excel_deeper_in_the_tree_is_spared_too_and_its_name_in_any_case():
    rs = _relay_module()
    ended = []
    excel = Proc(50, "excel.exe", [Proc(51, "conhost.exe", ended=ended)], ended)
    server = Proc(10, "python.exe", [Proc(20, "python.exe", [excel], ended)], ended)
    rs.end_tree(server)
    assert sorted(ended) == [10, 20]


def test_a_process_gone_meanwhile_does_not_stop_the_others_being_ended():
    rs = _relay_module()
    ended = []
    server = Proc(10, "python.exe", [Proc(20, "python.exe", gone=True, ended=ended), Proc(30, "node.exe", ended=ended),
                                     Proc(40, "EXCEL.EXE", [Proc(41, "x.exe", ended=ended)], ended)], ended)
    rs.end_tree(server)
    assert sorted(ended) == [10, 30]


def test_a_server_that_does_not_stop_in_time_has_its_tree_ended_but_not_its_excel(monkeypatch):
    rs = _relay_module()
    server, ended = _tree()
    closed = []

    class Child:
        pid = 10
        stdin = type("In", (), {"close": lambda self: closed.append(True)})()

        def wait(self, timeout):
            raise subprocess.TimeoutExpired("python", timeout)

        def kill(self):
            raise AssertionError("the tree's own end is used, not a kill of the child alone")
    monkeypatch.setattr(psutil, "Process", lambda pid: server if pid == 10 else None)
    monkeypatch.setattr(rs, "log", lambda msg: None)
    r = rs.Relay()
    r.child = Child()
    r.stop_child("test")
    assert closed == [True] and sorted(ended) == [10, 20, 21, 30]
