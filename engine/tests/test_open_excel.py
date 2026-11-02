import json
from pathlib import Path

import pytest
import pywintypes

from smartview_mcp import config, dialogs, session, styling
from smartview_mcp.excel import instances, jobs, login, target
from smartview_mcp.dialogs import Dialog
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
    monkeypatch.setattr(target.time, "sleep", lambda s: None)
    monkeypatch.setattr(jobs, "hold_turn", lambda *a, **k: True)
    monkeypatch.setattr(login, "wait_login_cleared", lambda *a, **k: None)


def test_open_path_fails_fast_when_the_new_excel_exits(tmp_path, monkeypatch, fast):
    f = tmp_path / "a.xlsx"
    f.write_bytes(b"x")
    proc = FakeProc(exits_after=2)
    monkeypatch.setattr(target, "_launch_excel", lambda p: proc)
    monkeypatch.setattr(instances, "interactive_excels", lambda: [])
    with pytest.raises(HostError, match="closed before it opened"):
        target._open_path(None, f, [], [])
    assert proc.polls == 2


def test_open_path_waits_for_a_slow_cold_excel(tmp_path, monkeypatch, fast):
    f = tmp_path / "a.xlsx"
    f.write_bytes(b"x")
    calls = {"n": 0}

    def excels():
        calls["n"] += 1
        return [(7, 1, FakeApp([FakeWb(str(f))]))] if calls["n"] > 120 else []   # appears after ~2 minutes

    monkeypatch.setattr(target, "_launch_excel", lambda p: FakeProc(exits_after=None))
    monkeypatch.setattr(instances, "interactive_excels", excels)
    clock = iter(range(0, 10_000))
    monkeypatch.setattr(target.time, "monotonic", lambda: next(clock))
    t = target._open_path(None, f, [], [])
    assert t.pid == 7


# ------------------------------------------------------------------ several Excels (fake windows and COM objects)
SAVEAS_FAILED = pywintypes.com_error(-2147352567, "Exception occurred.", (0, "Microsoft Excel", "SaveAs method of "
                                     "Workbook class failed", "xlmain11.chm", 0, -2146827284), None)


class Books(list):
    def __init__(self, items=(), add=None):
        super().__init__(items)
        self._add = add

    def Add(self, *a):
        wb = self._add()
        self.append(wb)
        return wb


class NewWb:
    def __init__(self, fail: bool):
        self.Name, self.fail, self.closed, self.saved_as = "tn_template_x1", fail, [], None
        self.sheet = type("Sheet", (), {"Name": "Sheet1"})()

    def Worksheets(self, i):
        return self.sheet

    def SaveAs(self, path, fmt):
        if self.fail:
            raise SAVEAS_FAILED
        self.saved_as = path
        open(path, "wb").close()

    def Close(self, save):
        self.closed.append(save)


def _excels(monkeypatch, windows: dict, apps: dict, engine=()):
    """windows: {pid: [captions]} in z-order; apps: {pid: FakeApp or None (does not answer COM)}."""
    config.PID_FILE.write_text(json.dumps([{"pid": p, "ctime": 1.0} for p in engine]))
    monkeypatch.setattr(instances, "_excel_windows",
                        lambda: [(pid, pid * 10, [(pid * 100 + i, c) for i, c in enumerate(caps)])
                                 for pid, caps in windows.items()])
    monkeypatch.setattr(instances, "_app_from_window", lambda ex7: apps.get(ex7 // 100))
    monkeypatch.setattr(instances, "_is_embedding", lambda pid: False)
    monkeypatch.setattr(login, "login_open", lambda pid: [])


def test_target_open_in_the_second_excel_is_used_there(tmp_path, monkeypatch, fast):
    f = tmp_path / "Book A.xlsx"
    f.write_bytes(b"x")
    first = FakeApp(Books([FakeWb(str(tmp_path / "other.xlsx"))]))
    second = FakeApp(Books([FakeWb(str(tmp_path / "x.xlsx")), FakeWb(str(f))]))
    _excels(monkeypatch, {11: ["other"], 22: ["x", "Book A"]}, {11: first, 22: second})
    monkeypatch.setattr(jobs, "com", lambda *a, **k: pytest.fail("must not open it in another Excel"))
    inst = instances.excel_instances()
    excels = [(i.pid, i.hwnd, i.app) for i in inst if i.app is not None]
    t = target._open_path(None, f, excels, inst)
    assert (t.pid, t.app, t.wb.FullName) == (22, second, str(f))


def test_engine_hidden_excel_is_never_a_user_excel(tmp_path, monkeypatch):
    f = tmp_path / "Book A.xlsx"
    engine, user = FakeApp(Books([FakeWb(str(f))])), FakeApp(Books([FakeWb(str(f))]))
    _excels(monkeypatch, {33: ["Book A"], 44: ["Book A"]}, {33: engine, 44: user}, engine=[33])
    assert [i.pid for i in instances.excel_instances()] == [44]
    assert [e[0] for e in instances.interactive_excels()] == [44]


@pytest.mark.parametrize("found,words", [([Dialog(5, "NUIDialog", "Microsoft Excel", "matrix", "message")], "dialog"),
                                         ([], "edit mode")])
def test_target_in_a_busy_excel_names_dialog_or_edit_mode(tmp_path, monkeypatch, fast, found, words):
    f = tmp_path / "2026-10-23_1318 MATRIX S8 scratch.xlsx"
    f.write_bytes(b"x")
    user = FakeApp(Books([FakeWb(str(tmp_path / "mine.xlsx"))]))
    _excels(monkeypatch, {11: ["mine"], 22: ["2026-10-23_1318 MATRIX S8 scratch  -  Read-Only"]}, {11: user, 22: None})
    monkeypatch.setattr(target, "BUSY_WAIT_S", 0)
    monkeypatch.setattr(dialogs, "scan", lambda pid: found if pid == 22 else [])
    monkeypatch.setattr(jobs, "com", lambda *a, **k: pytest.fail("must not open it in another Excel"))
    inst = instances.excel_instances()
    with pytest.raises(HostError, match=words) as e:
        target._open_path(None, f, [(11, 110, user)], inst)
    assert "pid 22" in str(e.value) and f.name in str(e.value)


def test_busy_excel_that_answers_in_time_is_used(tmp_path, monkeypatch, fast):
    f = tmp_path / "Book A.xlsx"
    f.write_bytes(b"x")
    later = FakeApp(Books([FakeWb(str(f))]))
    apps = {22: None}
    _excels(monkeypatch, {22: ["Book A"]}, apps)
    inst = instances.excel_instances()
    apps[22] = later                                # e.g. the user pressed Esc
    assert target._open_path(None, f, [], inst).app is later


def test_busy_excel_with_a_same_named_workbook_elsewhere_falls_back_to_open(tmp_path, monkeypatch, fast):
    f = tmp_path / "Book A.xlsx"
    f.write_bytes(b"x")
    (tmp_path / "other").mkdir()
    holder = FakeApp(Books([FakeWb(str(tmp_path / "other" / "Book A.xlsx"))]))
    apps = {22: None}
    _excels(monkeypatch, {11: ["mine"], 22: ["Book A"]}, apps)
    inst = instances.excel_instances()
    apps[22] = holder                               # answers COM, but its 'Book A' is another file
    opened = FakeWb(str(f))
    user = FakeApp(Books())
    user.Workbooks.Open = lambda p: opened
    monkeypatch.setattr(instances, "_file_in_use", lambda p: False)
    monkeypatch.setattr(jobs, "com", lambda fn, **k: fn())
    t = target._open_path(None, f, [(11, 110, user)], inst)
    assert (t.pid, t.wb) == (11, opened)


def test_read_only_file_is_not_in_use_but_a_locked_one_is(tmp_path):
    import win32file
    ro = tmp_path / "ro.xlsx"
    ro.write_bytes(b"x")
    ro.chmod(0o444)                                 # read-only attribute: access denied for write, Excel opens it R/O
    try:
        assert instances._file_in_use(ro) is False
    finally:
        ro.chmod(0o666)
    locked = tmp_path / "locked.xlsx"
    locked.write_bytes(b"x")
    h = win32file.CreateFile(str(locked), win32file.GENERIC_READ | win32file.GENERIC_WRITE, win32file.FILE_SHARE_READ,
                             None, win32file.OPEN_EXISTING, 0, None)   # like Excel holding it open
    try:
        assert instances._file_in_use(locked) is True
    finally:
        h.Close()
    assert instances._file_in_use(locked) is False


def test_locked_file_open_in_no_reachable_excel_is_named(tmp_path, monkeypatch, fast):
    f = tmp_path / "Book A.xlsx"
    f.write_bytes(b"x")
    monkeypatch.setattr(instances, "_file_in_use", lambda p: True)
    monkeypatch.setattr(jobs, "com", lambda *a, **k: pytest.fail("must not try Workbooks.Open"))
    with pytest.raises(HostError, match="open in another program"):
        target._open_path(None, f, [(11, 110, FakeApp(Books()))], [])


def test_failed_saveas_of_a_new_workbook_closes_it_unsaved(tmp_path, monkeypatch, fast):
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    monkeypatch.setattr(styling, "ensure_template", lambda app: tmp_path / "tn_template_x.xltx")
    monkeypatch.setattr(instances, "_label", lambda wb: None)
    new = NewWb(fail=True)
    app = FakeApp(Books([FakeWb(str(tmp_path / "mine.xlsx"))], add=lambda: new))
    with pytest.raises(HostError, match="could not save the new workbook"):
        target._new_target(None, "Q", [(11, 110, app)])
    assert new.closed == [False]


def test_new_workbook_never_reuses_a_file_name_from_the_same_minute(tmp_path, monkeypatch, fast):
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    monkeypatch.setattr(styling, "ensure_template", lambda app: tmp_path / "tn_template_x.xltx")
    monkeypatch.setattr(instances, "_label", lambda wb: None)
    made = []

    def add():
        made.append(NewWb(fail=False))
        return made[-1]
    app = FakeApp(Books(add=add))
    t1 = target._new_target(None, "Same question", [(11, 110, app)])
    t2 = target._new_target(None, "Same question", [(11, 110, app)])
    names = [Path(w.saved_as).name for w in made]
    assert t1.created and t2.created and names[0] != names[1] and names[1].endswith(" (2).xlsx")
