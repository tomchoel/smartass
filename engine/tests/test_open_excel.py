import json
import threading
import time
from pathlib import Path

import pytest
import pywintypes

from smartass.core import config, session
from smartass.xl import dialogs, styling
from smartass.xl import instances, jobs, target
from smartass.xl.dialogs import Dialog
from smartass.core.errors import HostError


class FakeProc:
    def __init__(self, exits_after):
        self.pid, self.polls, self.exits_after = 4242, 0, exits_after

    def poll(self):
        self.polls += 1
        return 0 if self.exits_after is not None and self.polls >= self.exits_after else None


class FakeWb:
    def __init__(self, full_name):
        self.FullName = full_name
        self.Name = Path(full_name).name


class FakeApp:
    def __init__(self, books):
        self.Workbooks = books


@pytest.fixture
def fast(monkeypatch):
    monkeypatch.setattr(target.time, "sleep", lambda s: None)
    monkeypatch.setattr(jobs, "hold_turn", lambda *a, **k: True)
    monkeypatch.setattr(dialogs, "watch_after_start", lambda *a, **k: None)
    monkeypatch.setattr(dialogs, "refuse_login", lambda *a, **k: None)


def test_open_path_fails_fast_when_the_new_excel_exits(tmp_path, monkeypatch, fast):
    f = tmp_path / "a.xlsx"
    f.write_bytes(b"x")
    proc = FakeProc(exits_after=2)
    monkeypatch.setattr(target, "_launch_excel", lambda p: proc)
    monkeypatch.setattr(instances, "interactive_excels", lambda: [])
    with pytest.raises(HostError, match="closed before it opened"):
        target._open_path(f, [], [])
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
    t = target._open_path(f, [], [])
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
    monkeypatch.setattr(dialogs, "login_open", lambda pid: [])


def test_target_open_in_the_second_excel_is_used_there(tmp_path, monkeypatch, fast):
    f = tmp_path / "Book A.xlsx"
    f.write_bytes(b"x")
    first = FakeApp(Books([FakeWb(str(tmp_path / "other.xlsx"))]))
    second = FakeApp(Books([FakeWb(str(tmp_path / "x.xlsx")), FakeWb(str(f))]))
    _excels(monkeypatch, {11: ["other"], 22: ["x", "Book A"]}, {11: first, 22: second})
    monkeypatch.setattr(jobs, "com", lambda *a, **k: pytest.fail("must not open it in another Excel"))
    inst = instances.excel_instances()
    excels = [(i.pid, i.hwnd, i.app) for i in inst if i.app is not None]
    t = target._open_path(f, excels, inst)
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
        target._open_path(f, [(11, 110, user)], inst)
    assert "pid 22" in str(e.value) and f.name in str(e.value)


def test_busy_excel_that_answers_in_time_is_used(tmp_path, monkeypatch, fast):
    f = tmp_path / "Book A.xlsx"
    f.write_bytes(b"x")
    later = FakeApp(Books([FakeWb(str(f))]))
    apps = {22: None}
    _excels(monkeypatch, {22: ["Book A"]}, apps)
    inst = instances.excel_instances()
    apps[22] = later                                # e.g. the user pressed Esc
    assert target._open_path(f, [], inst).app is later


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
    t = target._open_path(f, [(11, 110, user)], inst)
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
        target._open_path(f, [(11, 110, FakeApp(Books()))], [])


def test_failed_saveas_of_a_new_workbook_closes_it_unsaved(tmp_path, monkeypatch, fast):
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    monkeypatch.setattr(styling, "ensure_template", lambda app: tmp_path / "tn_template_x.xltx")
    monkeypatch.setattr(instances, "_label", lambda wb: None)
    new = NewWb(fail=True)
    app = FakeApp(Books([FakeWb(str(tmp_path / "mine.xlsx"))], add=lambda: new))
    with pytest.raises(HostError, match="could not save the new workbook"):
        target._new_target("Q", [(11, 110, app)])
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
    t1 = target._new_target("Same question", [(11, 110, app)])
    t2 = target._new_target("Same question", [(11, 110, app)])
    names = [Path(w.saved_as).name for w in made]
    assert t1.created and t2.created and names[0] != names[1] and names[1].endswith(" (2).xlsx")


# ------------------------------------------------------------------ an open workbook targeted by its name
class NamedWb:
    def __init__(self, name, full_name=None):
        self.Name, self.FullName = name, full_name or name


def _by_name(monkeypatch, windows: dict, apps: dict):
    monkeypatch.setattr(target, "_finish_target", lambda t, allow_archive: t)
    _excels(monkeypatch, windows, apps)


def test_unsaved_open_workbook_is_targeted_by_name(tmp_path, monkeypatch, fast):
    book1 = NamedWb("Book1")
    first = FakeApp(Books([NamedWb("other.xlsx", str(tmp_path / "other.xlsx"))]))
    second = FakeApp(Books([book1]))
    _by_name(monkeypatch, {11: ["other"], 22: ["Book1"]}, {11: first, 22: second})
    t = target.resolve_target("book1")
    assert (t.pid, t.wb) == (22, book1)


def test_open_workbook_is_targeted_by_name_without_extension(tmp_path, monkeypatch, fast):
    wb = NamedWb("Pack.xlsx", str(tmp_path / "Pack.xlsx"))
    _by_name(monkeypatch, {11: ["Pack.xlsx"]}, {11: FakeApp(Books([wb]))})
    assert target.resolve_target("Pack").wb is wb


def test_two_open_workbooks_with_the_same_name_are_ambiguous(tmp_path, monkeypatch, fast):
    from smartass.core.errors import InputError
    a = NamedWb("Pack.xlsx", str(tmp_path / "a" / "Pack.xlsx"))
    b = NamedWb("Pack.xlsx", str(tmp_path / "b" / "Pack.xlsx"))
    _by_name(monkeypatch, {11: ["Pack.xlsx"], 22: ["Pack.xlsx"]}, {11: FakeApp(Books([a])), 22: FakeApp(Books([b]))})
    with pytest.raises(InputError, match="ambiguous") as e:
        target.resolve_target("Pack.xlsx")
    assert str(tmp_path / "a" / "Pack.xlsx") in str(e.value) and str(tmp_path / "b" / "Pack.xlsx") in str(e.value)


def test_unknown_name_is_refused_naming_the_open_workbooks(tmp_path, monkeypatch, fast):
    from smartass.core.errors import InputError
    _by_name(monkeypatch, {11: ["Book1"]}, {11: FakeApp(Books([NamedWb("Book1"), NamedWb("x.xlam")]))})
    with pytest.raises(InputError, match="absolute path") as e:
        target.resolve_target("Book9")
    assert "Book1" in str(e.value) and "x.xlam" not in str(e.value)


def test_archive_guard_applies_to_a_workbook_targeted_by_name(tmp_path, monkeypatch, fast):
    from smartass.core.errors import InputError
    arch = config.ARCHIVE_ROOT / "2024" / "Model.xlsx"
    _excels(monkeypatch, {11: ["Model.xlsx"]}, {11: FakeApp(Books([NamedWb("Model.xlsx", str(arch))]))})
    with pytest.raises(InputError, match="archive"):
        target.resolve_target("Model.xlsx")



# ------------------------------------------------------------------ phase 3c Task 1: a new workbook the job builds
class Opener(Books):
    """Workbooks of an Excel that opens files (Open) and refuses Add (a new workbook the job builds is never added)."""

    def __init__(self, items=(), fail=None, read_only=False):
        super().__init__(items, add=lambda: pytest.fail("a workbook the job builds is never added"))
        self.fail, self.opened, self.read_only = fail, [], read_only

    def Open(self, path):
        if self.fail:
            raise self.fail
        wb = OpenedWb(path, self)
        wb.ReadOnly = self.read_only
        self.opened.append(wb)
        self.append(wb)
        return wb


class OpenedWb:
    def __init__(self, path, books):
        self.Name, self.FullName, self.books, self.closed = Path(path).name, str(path), books, []
        self.ReadOnly = False

    def Close(self, save):
        self.closed.append(save)
        self.books.remove(self)


def test_a_new_workbook_the_job_builds_gets_a_reserved_free_name_and_nothing_in_excel(tmp_path, monkeypatch, fast):
    """build_new: the Excel is chosen as for any new workbook, and a name free on disk and among that Excel's open
    workbooks is reserved (an empty file made exclusively, so a parallel job picks the next); nothing is added,
    opened or saved in Excel - the job writes the file, then open_new opens it."""
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    first = target._new_path("Q")
    books = Opener([FakeWb(str(tmp_path / "elsewhere" / first.name))])
    books[0].Name = first.name
    t = target._new_target("Q", [(11, 110, FakeApp(books))], build_new=True)
    assert t.wb is None and t.created and t.made.exists() and t.made.stat().st_size == 0
    assert t.made.name.endswith(" (2).xlsx") and t.book_name == t.made.name and t.path == str(t.made)
    assert not books.opened


def test_two_new_workbooks_reserved_at_the_same_moment_get_two_files(tmp_path, monkeypatch, fast):
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    same = target._new_path("Q")
    monkeypatch.setattr(target, "_new_path", lambda title, taken=frozenset(): same if same.name.lower() not in taken
                        else same.with_name(same.stem + " (2).xlsx"))
    a = target._new_target("Q", [(11, 110, FakeApp(Opener()))], build_new=True)
    b = target._new_target("Q", [(11, 110, FakeApp(Opener()))], build_new=True)
    assert a.made != b.made and a.made.exists() and b.made.exists()


def test_open_new_opens_the_built_file_in_the_chosen_excel(tmp_path, monkeypatch, fast):
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    books = Opener()
    t = target._new_target("Q", [(11, 110, FakeApp(books))], build_new=True)
    target.open_new(t)
    assert books.opened == [t.wb] and t.wb.FullName == str(t.made)


def test_a_new_workbook_is_discarded_closed_unsaved_and_its_reserved_file_deleted(tmp_path, monkeypatch, fast):
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    books = Opener()
    t = target._new_target("Q", [(11, 110, FakeApp(books))], build_new=True)
    target.open_new(t)
    wb = t.wb
    note = target.discard_new(t)
    assert wb.closed == [False] and not t.made.exists() and wb not in books
    assert note == f"the new workbook {t.made.name} was closed without saving and its file deleted"


def test_a_new_workbook_never_opened_is_only_deleted(tmp_path, monkeypatch, fast):
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    t = target._new_target("Q", [(11, 110, FakeApp(Opener()))], build_new=True)
    t.made.write_bytes(b"PK half a file")
    assert target.discard_new(t) == f"the new workbook {t.made.name} was not opened; its file was deleted"
    assert not t.made.exists()


def test_a_file_that_cannot_be_deleted_is_named_in_the_note(tmp_path, monkeypatch, fast):
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    t = target._new_target("Q", [(11, 110, FakeApp(Opener()))], build_new=True)
    tries = []

    def unlink(self, missing_ok=False):
        tries.append(self)
        raise PermissionError("held by Excel")
    monkeypatch.setattr(Path, "unlink", unlink)
    note = target.discard_new(t)
    assert len(tries) == target.DELETE_TRIES and note == (f"the new workbook {t.made.name} was not opened; its file "
                                                          f"could not be deleted ({t.made}): delete it yourself")


def test_a_failed_open_of_a_new_workbook_is_the_error_and_leaves_its_file(tmp_path, monkeypatch, fast):
    """open_new raises what Excel said; the job's clean-up (discard_new) deletes the file."""
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    books = Opener(fail=SAVEAS_FAILED)
    t = target._new_target("Q", [(11, 110, FakeApp(books))], build_new=True)
    with pytest.raises(pywintypes.com_error):
        target.open_new(t)
    assert t.wb is None and target.discard_new(t).endswith("its file was deleted") and not t.made.exists()


def test_a_busy_excel_rejecting_the_read_only_check_never_opens_the_new_workbook_twice(tmp_path, monkeypatch, fast):
    """Task D2 review, Minor 3: Workbooks.Open and wb.ReadOnly were one retried unit, so a busy rejection of the read
    (likely right after Excel opened a OneDrive file) ran the Open again."""
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    rejected = [2]

    class BusyAfterOpen(OpenedWb):
        @property
        def ReadOnly(self):
            if rejected[0]:
                rejected[0] -= 1
                raise pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)
            return False

        @ReadOnly.setter
        def ReadOnly(self, value):
            pass

    class Books(Opener):
        def Open(self, path):
            wb = BusyAfterOpen(path, self)
            self.opened.append(wb)
            self.append(wb)
            return wb
    books = Books()
    t = target._new_target("Q", [(11, 110, FakeApp(books))], build_new=True)
    target.open_new(t)
    assert len(books.opened) == 1 and t.wb is books.opened[0] and rejected == [0]


def test_a_new_workbook_excel_opens_read_only_is_refused_and_can_be_discarded(tmp_path, monkeypatch, fast):
    """A file Excel cannot save (no sensitivity label it knows: live, 20.11.2026, it opens read-only and its Save
    writes nothing, without an error) is an error at the open - as a SaveAs that made no file is - and the workbook
    is known to the clean-up, which closes it and deletes the file."""
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    books = Opener(read_only=True)
    t = target._new_target("Q", [(11, 110, FakeApp(books))], build_new=True)
    with pytest.raises(HostError, match="read-only"):
        target.open_new(t)
    wb = t.wb
    assert target.discard_new(t).endswith("closed without saving and its file deleted") and wb.closed == [False]
    assert not t.made.exists()


def test_without_an_excel_a_new_workbook_is_made_as_today_even_when_the_job_would_build_it(tmp_path, monkeypatch,
                                                                                           fast):
    """No Excel at all (the seed gives or a build makes the file without Excel, then an Excel is launched on it) and an
    Excel started for a parallel agent (on the TN template) keep today's way: the Target comes back open."""
    made = tmp_path / "made.xlsx"
    monkeypatch.setattr(target, "_new_analysis_file", lambda title: made)
    monkeypatch.setattr(target, "_open_path", lambda p, excels: target.Target(None, FakeWb(str(p)), 1, 1))
    t = target._new_target("Q", [], build_new=True)
    assert t.made is None and t.wb.FullName == str(made) and t.created


def test_resolve_target_leaves_the_session_alone_for_a_workbook_the_job_builds(tmp_path, monkeypatch, fast):
    """The session's workbook is set once the answer is in it (exceljob.excel_job): a job that fails deletes the
    file. A failure while the Excel is made ready (the bridge) deletes the reserved file too."""
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    monkeypatch.setattr(session, "set_workbook", lambda p: pytest.fail("the session is set after the answer"))
    app = FakeApp(Opener())
    monkeypatch.setattr(instances, "excel_instances", lambda: [instances.XlInstance(11, 110, [], app)])
    monkeypatch.setattr(target, "_finish_target", lambda t, allow_archive: t)
    t = target.resolve_target("new", build_new=True)
    assert t.wb is None and t.made.exists()

    def broken(t, allow_archive):
        raise HostError("the bridge add-in did not load")
    monkeypatch.setattr(target, "_finish_target", broken)
    with pytest.raises(HostError, match="did not load; the new workbook .* was not opened; its file was deleted"):
        target.resolve_target("new", build_new=True)
    assert [p.name for p in tmp_path.iterdir()] == [t.made.name]



def test_a_new_workbook_whose_open_was_abandoned_is_found_by_its_name_and_closed(tmp_path, monkeypatch, fast):
    """Review minor 1: a Workbooks.Open the job gave up on (a timeout) may still open the file: with no workbook in
    hand the clean-up looks for it among that Excel's workbooks by its name (reserved: no other open workbook has
    it), closes it without saving, then deletes the file."""
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    books = Opener()
    t = target._new_target("Q", [(11, 110, FakeApp(books))], build_new=True)
    late = books.Open(str(t.made))                      # Excel opened it after the call was abandoned
    assert t.wb is None
    note = target.discard_new(t)
    assert late.closed == [False] and late not in books and not t.made.exists()
    assert note == f"the new workbook {t.made.name} was closed without saving and its file deleted"


def test_another_sessions_workbook_of_the_same_name_is_never_closed_by_the_clean_up(tmp_path, monkeypatch, fast):
    """Final review M8: two sessions in different project folders reserved the same title in the same minute, and
    this job's Open failed on the name the other session's workbook holds: the clean-up finds the new workbook by its
    full path, so the other one stays open; only this job's file is deleted."""
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    books = Opener()
    t = target._new_target("Q", [(11, 110, FakeApp(books))], build_new=True)
    theirs = books.Open(str(tmp_path / "other project" / t.made.name))
    assert target.discard_new(t) == f"the new workbook {t.made.name} was not opened; its file was deleted"
    assert theirs.closed == [] and theirs in books and not t.made.exists()


def test_the_resolve_takes_no_runner():
    """Task 6 left the resolve's h (the job's runner) unused: the target functions take none (light-fast phase 2,
    Task 7)."""
    import inspect
    for fn in (target.resolve_target, target._resolve, target._new_target, target._open_path, target._finish_target):
        assert "h" not in inspect.signature(fn).parameters, fn.__name__


def test_a_note_keeps_the_shape_of_a_com_error():
    """Review minor 5: a COM error keeps its four args (hresult, text, excepinfo, argerror); the note goes into its
    text. Any other error's text gets it at the end."""
    e = pywintypes.com_error(-2147352567, "Exception occurred.", (0, "Microsoft Excel", "Open failed", None, 0, -1),
                             None)
    target._noted(e, "the new workbook x.xlsx was not opened; its file was deleted")
    assert len(e.args) == 4 and e.args[0] == -2147352567 and e.args[2][2] == "Open failed"
    assert e.args[1] == "Exception occurred.; the new workbook x.xlsx was not opened; its file was deleted"
    h = HostError("the bridge did not load")
    target._noted(h, "deleted")
    assert h.args == ("the bridge did not load; deleted",)


def test_a_new_question_waits_a_moment_for_a_turn_being_handed_over_before_starting_another_excel(tmp_path,
                                                                                                  monkeypatch):
    """Task D2 review, Minor 2: a job keeps its turn while the objects it made are released after its answer (tens of
    ms); a new question's request (target 'new') that came then found that Excel taken and started another, visible
    Excel. It waits a moment (jobs.FREE_WAIT_S) for a turn that frees."""
    monkeypatch.setattr(config, "SVKIT", tmp_path / "svkit")
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)
    tpl = tmp_path / "tn.xltx"
    tpl.write_bytes(b"x")
    monkeypatch.setattr(styling, "template_path", lambda *a: tpl)
    monkeypatch.setattr(target, "_spawn_excel", lambda tpl, known: pytest.fail("started another Excel"))
    monkeypatch.setattr(jobs, "_first_contact", lambda app: None)
    held = threading.Event()

    def previous_job():                              # answered; its objects are being released, its turn still held
        tt = jobs.Turn(11)
        assert tt.try_acquire()
        held.set()
        time.sleep(0.2)
        tt.close()
    th = threading.Thread(target=previous_job)
    th.start()
    assert held.wait(5)
    jobs.bind_job()
    try:
        t = target._new_target("Q", [(11, 110, FakeApp(Opener()))], build_new=True)
        assert t.pid == 11
    finally:
        jobs.release_turns()
        th.join(5)


# ---- Task 6b: _spawn_excel (now also the cold route) as _open_path: an Excel that exits, a cold Excel that rejects

def test_spawn_excel_fails_fast_when_the_new_excel_exits(tmp_path, monkeypatch, fast):
    proc = FakeProc(exits_after=2)
    monkeypatch.setattr(target, "_launch_excel", lambda p: proc)
    monkeypatch.setattr(instances, "interactive_excels", lambda: [])
    with pytest.raises(HostError, match=r"Excel \(pid 4242\) closed before it opened the TN template"):
        target._spawn_excel(tmp_path / "tn_template.xltx", set(), target.OPEN_TIMEOUT_S)
    assert proc.polls == 2


def test_spawn_excel_waits_while_a_cold_excel_rejects_its_workbooks(tmp_path, monkeypatch, fast):
    """A cold Excel rejects calls while it loads (Smart View, OneDrive): its workbooks are read again next round, not
    raised out of the scan."""
    class ColdApp:
        def __init__(self):
            self.reads = 0

        @property
        def Workbooks(self):
            self.reads += 1
            if self.reads <= 2:
                raise pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)
            return [FakeWb(str(tmp_path / "tn_template1"))]
    app = ColdApp()
    monkeypatch.setattr(target, "_launch_excel", lambda p: FakeProc(exits_after=None))
    monkeypatch.setattr(instances, "interactive_excels", lambda: [(9, 90, app)])
    (pid, _hwnd, got), wb = target._spawn_excel(tmp_path / "tn_template.xltx", set(), target.OPEN_TIMEOUT_S)
    assert (pid, got, app.reads) == (9, app, 3) and wb.Name == "tn_template1"


def test_spawn_excel_waits_while_a_cold_excel_rejects_a_workbooks_name(tmp_path, monkeypatch, fast):
    class ColdBook:
        reads = 0

        @property
        def Name(self):
            ColdBook.reads += 1
            if ColdBook.reads == 1:
                raise pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)
            return "tn_template1"
    app = FakeApp([ColdBook()])
    monkeypatch.setattr(target, "_launch_excel", lambda p: FakeProc(exits_after=None))
    monkeypatch.setattr(instances, "interactive_excels", lambda: [(9, 90, app)])
    (pid, _hwnd, got), _wb = target._spawn_excel(tmp_path / "tn_template.xltx", set(), target.OPEN_TIMEOUT_S)
    assert (pid, got, ColdBook.reads) == (9, app, 2)


def test_spawn_excel_raises_a_com_error_that_is_not_busy(tmp_path, monkeypatch, fast):
    class BrokenApp:
        @property
        def Workbooks(self):
            raise pywintypes.com_error(-2147352567, "Exception occurred.", None, None)
    monkeypatch.setattr(target, "_launch_excel", lambda p: FakeProc(exits_after=None))
    monkeypatch.setattr(instances, "interactive_excels", lambda: [(9, 90, BrokenApp())])
    with pytest.raises(pywintypes.com_error):
        target._spawn_excel(tmp_path / "tn_template.xltx", set())
