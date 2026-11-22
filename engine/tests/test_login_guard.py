"""Smart View's Login in the user's Excel (Task 6 ruling): smartass uses no Smart View any more, but an Excel whose
Login is open is never touched - COM into it blocks. Detection by window enumeration only (xl/dialogs.py); nothing
waits for or clicks the Login (the SSO watcher is gone): the job is refused at once, naming that Excel. On fakes:
no window, no Excel."""
import pytest

from smartass.core.errors import HostError
from smartass.xl import dialogs, instances, jobs, target


class NoCom:
    """An Excel object any COM call on which fails the test."""

    def __getattr__(self, name):
        raise AssertionError(f"COM into an Excel whose Login is open: {name}")


def _windows(monkeypatch, logins: set[int]):
    monkeypatch.setattr(instances, "_excel_windows", lambda: [(11, 110, [(1100, "Budget.xlsx")]),
                                                              (22, 220, [(2200, "Book1")])])
    monkeypatch.setattr(instances, "_is_embedding", lambda pid: False)
    monkeypatch.setattr(instances, "engine_excel_pids", lambda: set())
    monkeypatch.setattr(dialogs, "login_open", lambda pid: [5] if pid in logins else [])
    asked = []
    monkeypatch.setattr(instances, "_app_from_window", lambda ex7: asked.append(ex7) or f"app{ex7 // 100}")
    return asked


def test_a_login_open_in_any_excel_refuses_at_once_naming_it_and_touching_none(monkeypatch):
    _windows(monkeypatch, {22})
    with pytest.raises(HostError) as e:
        instances.excel_instances()
    said = str(e.value)
    assert "Smart View's Login is open in the Excel showing 'Book1' (pid 22)" in said
    assert "sign in there or close that Login window, then ask again" in said and "Check for updates" in said


def test_an_excel_with_no_workbook_window_is_named_by_its_main_windows_title(monkeypatch):
    """A fresh Excel showing only the Login has no workbook window yet: the refusal reads its main window's title
    (dialogs.titles, no COM) instead of naming the pid alone."""
    monkeypatch.setattr(instances, "_excel_windows", lambda: [(33, 330, [])])
    monkeypatch.setattr(instances, "_is_embedding", lambda pid: False)
    monkeypatch.setattr(instances, "engine_excel_pids", lambda: set())
    monkeypatch.setattr(dialogs, "login_open", lambda pid: [5])
    monkeypatch.setattr(dialogs, "windows", lambda pid=None, parent=None: [(330, "XLMAIN", "Excel"),
                                                                           (5, "#32770", "Login")])
    with pytest.raises(HostError, match=r"Smart View's Login is open in the Excel showing 'Excel' \(pid 33\)"):
        instances.excel_instances()


def test_without_a_login_the_excels_are_listed_as_before(monkeypatch):
    asked = _windows(monkeypatch, set())
    got = instances.excel_instances()
    assert [(i.pid, i.app, i.captions) for i in got] == [(11, "app11", ["Budget.xlsx"]), (22, "app22", ["Book1"])]
    assert asked == [1100, 2200]


def test_a_login_never_falls_through_to_launching_another_excel(monkeypatch, tmp_path):
    """No 'no Excel found, launch one' and no 'file locked': the refusal is the job's answer."""
    _windows(monkeypatch, {11})
    monkeypatch.setattr(target, "_launch_excel", lambda path: pytest.fail("an Excel was launched"))
    for t in ("new", "active", str(tmp_path / "x.xlsx")):
        with pytest.raises(HostError, match="Smart View's Login is open"):
            target.resolve_target(t)


def test_the_target_is_refused_before_its_first_com_call_when_a_login_opened_meanwhile(monkeypatch):
    """Named by its main window's title, read without COM (no COM into that Excel either way)."""
    monkeypatch.setattr(dialogs, "login_open", lambda pid: [5])
    monkeypatch.setattr(dialogs, "windows", lambda pid=None, parent=None: [(70, "XLMAIN", "Budget.xlsx - Excel"),
                                                                           (5, "#32770", "Login")])
    with pytest.raises(HostError, match=r"Smart View's Login is open in the Excel showing 'Budget.xlsx - Excel' "
                                        r"\(pid 7\)"):
        target._finish_target(target.Target(NoCom(), NoCom(), 7, 70), allow_archive=False)


def test_an_excel_whose_title_cannot_be_read_is_named_by_its_pid(monkeypatch):
    monkeypatch.setattr(dialogs, "windows", lambda pid=None, parent=None: [])
    assert str(dialogs.login_refusal(7)).startswith("Smart View's Login is open in Excel (pid 7);")


def test_without_a_login_the_target_is_made_ready_as_before(monkeypatch):
    monkeypatch.setattr(dialogs, "login_open", lambda pid: [])
    monkeypatch.setattr(target.Target, "path", property(lambda self: "C:/a/b.xlsx"))
    monkeypatch.setattr(target.Target, "full_name", property(lambda self: "C:/a/b.xlsx"))
    monkeypatch.setattr(target.Target, "name", property(lambda self: "b.xlsx"))
    ready = []
    monkeypatch.setattr(target, "_wait_ready", lambda app, **k: ready.append(k))
    t = target.Target(object(), object(), 7, 70)
    assert target._finish_target(t, allow_archive=False) is t and ready == [{"pid": 7, "name": "b.xlsx"}]


class Clock:
    def __init__(self):
        self.t, self.sleeps = 0.0, 0

    def now(self):
        return self.t

    def sleep(self, s):
        self.t += s
        self.sleeps += 1


def test_after_a_launch_a_login_that_comes_is_refused_at_once(monkeypatch):
    c = Clock()
    monkeypatch.setattr(dialogs, "login_open", lambda pid: [5] if c.t >= 4 else [])
    monkeypatch.setattr(dialogs, "windows", lambda pid=None, parent=None: [(90, "XLMAIN", "Book1 - Excel")])
    with pytest.raises(HostError, match=r"Login is open in the Excel showing 'Book1 - Excel' \(pid 9\)"):
        dialogs.watch_after_start(9, clock=c.now, sleep=c.sleep)
    assert 4 <= c.t < 5                                  # not the whole grace, and no wait for a click


def test_after_a_launch_without_a_login_the_whole_grace_is_watched_as_before(monkeypatch):
    """Today's 12 s after smartass launched an Excel (Smart View's start-up Login is seen before the first COM)."""
    c = Clock()
    monkeypatch.setattr(dialogs, "login_open", lambda pid: [])
    checks = []
    dialogs.watch_after_start(9, clock=c.now, sleep=c.sleep, alive=lambda: checks.append(c.t))
    assert dialogs.LAUNCH_GRACE_S == 12 and 12 <= c.t < 13 and len(checks) == c.sleeps + 1


def test_a_launched_excel_whose_workbook_opened_is_watched_before_it_is_used(monkeypatch, tmp_path):
    """_open_path: the Excel launched on the file shows it; its Login then comes - the job is refused."""
    f = tmp_path / "a.xlsx"
    f.write_bytes(b"x")

    class Book:
        FullName = str(f)

    class App:
        Workbooks = [Book()]

    class Proc:
        pid = 4242

        def poll(self):
            return None
    monkeypatch.setattr(target, "_launch_excel", lambda path: Proc())
    monkeypatch.setattr(target.time, "sleep", lambda s: None)
    monkeypatch.setattr(jobs, "hold_turn", lambda *a, **k: True)
    monkeypatch.setattr(instances, "interactive_excels", lambda: [(9, 90, App())])
    watched = []

    def watch(pid, **k):
        watched.append(pid)
        raise dialogs.login_refusal(pid, [])
    monkeypatch.setattr(dialogs, "watch_after_start", watch)
    with pytest.raises(HostError, match="Login is open"):
        target._open_path(f, [], [])
    assert watched == [9]


# ---- Task 6b: a Login after the 12 s grace, where Smart View's start-up check is on (read from the registry, no COM)

def _shown_3_s_after_the_launch(monkeypatch, value):
    """An Excel launched at t=0 whose workbook showed at t=3; Smart View's Login (if any) comes at t=21. value: what
    the registry gives for the start-up check (an exception: it cannot be read)."""
    c = Clock()
    c.t = 3.0

    def read(path, name):
        assert (path, name) == dialogs.STARTUP_CHECK
        if isinstance(value, Exception):
            raise value
        return value
    monkeypatch.setattr(dialogs, "read_hkcu", read)
    monkeypatch.setattr(dialogs, "login_open", lambda pid: [5] if c.t >= 21 else [])
    monkeypatch.setattr(dialogs, "windows", lambda pid=None, parent=None: [])
    return c


def test_with_the_start_up_check_on_a_login_21_s_after_the_launch_is_refused_at_once(monkeypatch):
    """The watch runs to STARTUP_LOGIN_S after the launch (not 12 s after the workbook showed): the Login seen there
    is the refusal naming that Excel, not a COM call that blocks until the job's deadline."""
    c = _shown_3_s_after_the_launch(monkeypatch, 1)
    with pytest.raises(HostError, match=r"Smart View's Login is open in Excel \(pid 9\)"):
        dialogs.watch_after_start(9, grace=dialogs.launch_grace(0.0, clock=c.now), clock=c.now, sleep=c.sleep)
    assert 21 <= c.t < 21.5


@pytest.mark.parametrize("value", [0, None, OSError("access denied")], ids=["off", "not set", "unreadable"])
def test_with_the_start_up_check_off_or_unreadable_the_watch_is_todays_12_s(monkeypatch, value):
    c = _shown_3_s_after_the_launch(monkeypatch, value)
    dialogs.watch_after_start(9, grace=dialogs.launch_grace(0.0, clock=c.now), clock=c.now, sleep=c.sleep)
    assert 15 <= c.t < 15.5                              # 12 s from when the workbook showed, as before


def test_with_the_start_up_check_on_the_watch_is_never_shorter_than_todays_grace(monkeypatch):
    c = _shown_3_s_after_the_launch(monkeypatch, 1)
    assert dialogs.STARTUP_LOGIN_S == 25 and dialogs.launch_grace(0.0, clock=c.now) == 22
    c.t = 20.0                                           # a slow start: the workbook showed 20 s after the launch
    assert dialogs.launch_grace(0.0, clock=c.now) == dialogs.LAUNCH_GRACE_S


def test_the_suite_never_reads_this_pcs_smart_view_setting():
    assert dialogs.startup_check_on() is None


class _Tpl:
    Name = "tn_template1"


@pytest.mark.parametrize("value, grace", [(1, (20, 25)), (0, (12, 12)), (None, (12, 12))], ids=["on", "off", "unset"])
@pytest.mark.parametrize("launch", ["_spawn_excel", "_open_path"])
def test_both_launch_sites_watch_from_the_launch(monkeypatch, tmp_path, value, grace, launch):
    """_spawn_excel and _open_path give the watch launch_grace from the moment they launched Excel."""
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"x")

    class Book:
        Name, FullName = "tn_template1", str(book)

    class App:
        Workbooks = [Book()]

    class Proc:
        pid = 4242

        def poll(self):
            return None
    monkeypatch.setattr(dialogs, "read_hkcu", lambda path, name: value)
    monkeypatch.setattr(target, "_launch_excel", lambda path: Proc())
    monkeypatch.setattr(target.time, "sleep", lambda s: None)
    monkeypatch.setattr(jobs, "hold_turn", lambda *a, **k: True)
    monkeypatch.setattr(instances, "interactive_excels", lambda: [(9, 90, App())])
    watched = []
    monkeypatch.setattr(dialogs, "watch_after_start", lambda pid, **k: watched.append((pid, k["grace"])))
    if launch == "_spawn_excel":
        target._spawn_excel(tmp_path / "tn_template.xltx", set())
    else:
        target._open_path(book, [], [])
    assert [p for p, _g in watched] == [9] and grace[0] <= watched[0][1] <= grace[1]
