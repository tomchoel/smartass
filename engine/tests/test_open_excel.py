import pytest

from smartview_mcp import userexcel
from smartview_mcp.host import HostError


class FakeProc:
    def __init__(self, exits_after):
        self.pid, self.polls, self.exits_after = 4242, 0, exits_after

    def poll(self):
        self.polls += 1
        return 0 if self.exits_after is not None and self.polls >= self.exits_after else None


class FakeWb:
    def __init__(self, full_name):
        self.FullName = full_name


class FakeApp:
    def __init__(self, books):
        self.Workbooks = books


@pytest.fixture
def fast(monkeypatch):
    monkeypatch.setattr(userexcel.time, "sleep", lambda s: None)
    monkeypatch.setattr(userexcel, "hold_turn", lambda *a, **k: True)
    monkeypatch.setattr(userexcel, "wait_login_cleared", lambda *a, **k: None)


def test_open_path_fails_fast_when_the_new_excel_exits(tmp_path, monkeypatch, fast):
    f = tmp_path / "a.xlsx"
    f.write_bytes(b"x")
    proc = FakeProc(exits_after=2)
    monkeypatch.setattr(userexcel, "_launch_excel", lambda p: proc)
    monkeypatch.setattr(userexcel, "interactive_excels", lambda: [])
    with pytest.raises(HostError, match="closed before it opened"):
        userexcel._open_path(None, f, [])
    assert proc.polls == 2


def test_open_path_waits_for_a_slow_cold_excel(tmp_path, monkeypatch, fast):
    f = tmp_path / "a.xlsx"
    f.write_bytes(b"x")
    calls = {"n": 0}

    def excels():
        calls["n"] += 1
        return [(7, 1, FakeApp([FakeWb(str(f))]))] if calls["n"] > 120 else []   # appears after ~2 minutes

    monkeypatch.setattr(userexcel, "_launch_excel", lambda p: FakeProc(exits_after=None))
    monkeypatch.setattr(userexcel, "interactive_excels", excels)
    clock = iter(range(0, 10_000))
    monkeypatch.setattr(userexcel.time, "monotonic", lambda: next(clock))
    t = userexcel._open_path(None, f, [])
    assert t.pid == 7
