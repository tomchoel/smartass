"""Phase 3c Task 1b: a job's start in an Excel that rejects calls - the reads before its turn wait for a busy Excel
(a save after the last reply runs there for about a second), the start gives up with the busy message after 30 s in
all as before the save after the reply (a cell in edit mode, a dialog), only the add-in's own answer that a save of
ours runs is waited for longer, and an abandoned job stops in every wait. On fakes, with a fake clock: no real wait."""
from __future__ import annotations

import time
from pathlib import Path

import pytest
import pywintypes

from smartass.core import config
from smartass.xl import dialogs, instances, jobs, login, target
from smartass.xl.host import HostError
from smartass.xl.inproc import later, runner
from smartass.xl.workbench import books

REJECTED = pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)


@pytest.fixture
def clock(monkeypatch):
    """time.monotonic and time.sleep made a fake clock (a sleep moves it on); the job's token aborted at the sleep
    numbered abort_at."""
    now = {"t": 1000.0, "sleeps": 0, "abort_at": None}
    token = jobs.JobToken()

    def sleep(s):
        now["t"] += s
        now["sleeps"] += 1
        if now["abort_at"] is not None and now["sleeps"] >= now["abort_at"]:
            token.aborted = True
    monkeypatch.setattr(time, "monotonic", lambda: now["t"])
    monkeypatch.setattr(time, "sleep", sleep)
    monkeypatch.setattr(later, "_queued", {})                 # no save asked for here: the start's 30 s hold
    monkeypatch.setattr(later, "_asked_at", {})
    jobs.bind_job(token)
    yield now
    jobs.release_turns()
    jobs.bind_job()


class Rejects:
    """An attribute read Excel rejects until the fake clock reaches until (None: for ever)."""
    def __init__(self, clock, until=None):
        self.clock, self.until = clock, until

    def __call__(self, value):
        if self.until is None or self.clock["t"] < self.until:
            raise REJECTED
        return value


class Book:
    def __init__(self, name, full, rejects=None):
        self._name, self.FullName, self.rejects = name, full, rejects

    @property
    def Name(self):
        return self.rejects(self._name) if self.rejects else self._name


class App:
    def __init__(self, books=(), active=None, rejects=None, ready=True):
        self._books, self.ActiveWorkbook, self.rejects, self._ready = list(books), active, rejects, ready

    @property
    def Workbooks(self):
        return self.rejects(self._books) if self.rejects else self._books

    @property
    def Ready(self):
        return self.rejects(self._ready) if self.rejects else self._ready


# ------------------------------------------------------------------ reads before the turn wait for a busy Excel
def test_the_active_workbooks_name_is_read_with_a_busy_excel_waited_for(clock, monkeypatch):
    """Review 2 (Task 2): target='active' read the active workbook's name with plain COM before its turn; an 'active'
    question sent within about 1.4 s of a reply failed with "Call was rejected by callee"."""
    from smartass.epm import sso
    wb = Book("answer.xlsx", r"C:\a\answer.xlsx", Rejects(clock, until=clock["t"] + 1))
    app = App(active=wb)
    monkeypatch.setattr(sso, "ensure_global_watcher", lambda: True)
    monkeypatch.setattr(instances, "excel_instances", lambda: [instances.XlInstance(7, 70, ["answer"], app)])
    monkeypatch.setattr(jobs, "hold_turn", lambda *a, **k: True)
    monkeypatch.setattr(target, "_finish_target", lambda h, t, allow: t)
    t = target.resolve_target(None, "active")
    assert t.wb is wb and t.pid == 7


@pytest.mark.parametrize("which", ["answer.xlsx", None])
def test_an_xl_tool_finds_its_workbook_with_a_busy_excel_waited_for(clock, monkeypatch, which):
    """Review 2 (Task 2): find_book (every xl_* tool) scanned the open workbooks' names with plain COM before its turn,
    so a tool sent right after a reply failed while Excel saved."""
    wb = Book("answer.xlsx", r"C:\a\answer.xlsx", Rejects(clock, until=clock["t"] + 1))
    app = App([Book("other.xlsx", r"C:\a\other.xlsx"), wb], active=wb)
    monkeypatch.setattr(books, "excel_instances", lambda: [instances.XlInstance(7, 70, ["answer"], app)])
    monkeypatch.setattr(books, "hold_turn", lambda *a, **k: True)
    assert books.find_book(which).wb is wb


# ------------------------------------------------------------------ the start's waits: 30 s in all
def test_a_jobs_start_in_an_excel_that_rejects_calls_gives_up_after_30_s_in_all(clock):
    """The scan before the turn waited 60 s and the wait for Ready 30 s more (review 2: the waits added up to 240 s
    where the start gave up after 30 s before the save after the reply): now they share one 30 s."""
    t0 = clock["t"]
    path = Path(r"C:\a\answer.xlsx")
    with jobs.starting():
        app = App([Book("answer.xlsx", str(path))], rejects=Rejects(clock, until=t0 + 20))
        assert target._find_open([(7, 70, app)], path) is not None          # answers after 20 s
        app.rejects = Rejects(clock)                                         # then a cell in edit mode, for ever
        with pytest.raises(HostError, match="edit mode"):
            target._wait_ready(app)
    assert clock["t"] - t0 <= jobs.START_WAIT_S + 2


def test_the_time_a_start_spends_otherwise_is_not_counted(clock):
    """Another agent's turn (up to 10 min), a login or Excel opening a file are not waits for a busy Excel: a wait
    after them still has the start's 30 s."""
    app = App(rejects=Rejects(clock, until=clock["t"] + 320))
    t0 = clock["t"]
    with jobs.starting():
        time.sleep(300)                                                      # another agent's turn, say
        target._wait_ready(app)                                              # rejects 20 s more, then ready
    assert clock["t"] - t0 >= 320


def test_a_save_the_first_contact_runs_itself_is_not_counted_against_the_start(clock, monkeypatch):
    """Review (Task 1b): a job sent within the 0.3 s delay has the previous save run inside its first contact (the
    add-in's Clear saves, then answers "clear"). A 35 s save there, then one rejected call of the wait for Ready: the
    job goes on - only Excel's rejections count towards the start's 30 s."""
    t0 = clock["t"]

    def run(macro):
        clock["t"] += 35.0
        return "clear"
    app = _no_add_in_loaded(monkeypatch, run)
    with jobs.starting():
        later.clear(app)
        target._wait_ready(App(rejects=Rejects(clock, until=t0 + 35.4)))
    assert clock["t"] - t0 >= 35


def test_while_a_save_asked_for_here_may_still_run_the_start_waits_longer_for_rejections(clock, monkeypatch):
    """Review (Task 1b): a save this process asked for after its last reply and whose outcome no job has read yet may
    be what Excel is busy with: its rejections are waited for as long as each wait's own limit, not the start's 30 s."""
    path = Path(r"C:\a\answer.xlsx")
    monkeypatch.setattr(later, "_queued", {"t1": (7, 0.0)})
    monkeypatch.setattr(later, "_asked_at", {"t1": time.time()})
    app = App([Book("answer.xlsx", str(path))], rejects=Rejects(clock, until=clock["t"] + 45))
    with jobs.starting():
        assert target._find_open([(7, 70, app)], path) is not None


def _no_add_in_loaded(monkeypatch, run):
    app = type("App", (), {"Run": staticmethod(run)})()
    monkeypatch.setattr(config, "INPROC", True)
    monkeypatch.setattr(runner, "_EXCELS", {})
    monkeypatch.setattr(runner, "_identify", lambda app, check=True: (7, 0.0))
    monkeypatch.setattr(later, "_build_hash", lambda: "abcd1234")
    return app


def test_the_first_contact_with_an_excel_that_rejects_calls_gives_up_with_the_start(clock, monkeypatch):
    def run(macro):
        raise REJECTED
    app = _no_add_in_loaded(monkeypatch, run)
    t0 = clock["t"]
    with jobs.starting(), pytest.raises(HostError, match="Excel is busy"):
        later.clear(app)
    assert clock["t"] - t0 <= jobs.START_WAIT_S + 2


def test_a_save_of_ours_that_runs_is_waited_for_beyond_the_starts_30_s(clock, monkeypatch):
    """The add-in's own answer that a save runs (and Excel's rejections after it: the save) keeps CLEAR_LIMIT_S."""
    t0 = clock["t"]

    def run(macro):
        if clock["t"] < t0 + 40:
            return runner.BUSY + "the save after the reply is running"
        if clock["t"] < t0 + 50:
            raise REJECTED
        return "clear"
    app = _no_add_in_loaded(monkeypatch, run)
    with jobs.starting():
        later.clear(app)
    assert clock["t"] - t0 >= 50


# ------------------------------------------------------------------ an abandoned job stops in every wait
def _rejected_com(clock, monkeypatch, tmp_path):
    jobs.com(lambda: Rejects(clock)(1), limit=600)


def _asked(clock, monkeypatch, tmp_path):
    runner.asked(lambda: Rejects(clock)(1), 600)


def _ready(clock, monkeypatch, tmp_path):
    target._wait_ready(App(ready=False), limit=600)


def _turn(clock, monkeypatch, tmp_path):
    other = jobs.Turn(4711)
    assert other.try_acquire()
    mine = jobs.Turn(4711)
    try:
        with pytest.raises(jobs.JobAborted):
            mine.acquire("this request")
        assert mine.f.closed
    finally:
        other.close()
    raise jobs.JobAborted("checked")


def _busy_holder(clock, monkeypatch, tmp_path):
    monkeypatch.setattr(jobs, "hold_turn", lambda *a, **k: True)
    monkeypatch.setattr(instances, "excel_instances", lambda: [instances.XlInstance(7, 70, ["answer"], None)])
    monkeypatch.setattr(target, "BUSY_WAIT_S", 600)
    target._wait_busy_holder(7, Path(r"C:\a\answer.xlsx"))


def _login(clock, monkeypatch, tmp_path):
    monkeypatch.setattr(login, "login_open", lambda pid: [1])
    monkeypatch.setattr(dialogs, "is_hung", lambda w: False)
    monkeypatch.setattr(dialogs, "scan", lambda pid: [])
    monkeypatch.setattr(login.timing, "say", lambda *a, **k: None)
    login.wait_login_cleared(7, limit=600)


class _Proc:
    pid = 4242

    def poll(self):
        return None


def _spawn(clock, monkeypatch, tmp_path):
    monkeypatch.setattr(target, "_launch_excel", lambda path: _Proc())
    monkeypatch.setattr(instances, "interactive_excels", lambda: [])
    target._spawn_excel(tmp_path / "tn_template.xltx", set())


def _launched(clock, monkeypatch, tmp_path):
    book = tmp_path / "answer.xlsx"
    book.write_bytes(b"x")
    monkeypatch.setattr(target, "_launch_excel", lambda path: _Proc())
    monkeypatch.setattr(instances, "interactive_excels", lambda: [])
    target._open_path(None, book, [], [])


@pytest.mark.parametrize("wait", [_rejected_com, _asked, _ready, _turn, _busy_holder, _login, _spawn, _launched],
                         ids=lambda f: f.__name__.strip("_"))
def test_an_abandoned_job_stops_in_every_wait(clock, monkeypatch, tmp_path, wait):
    """Review 2 (Task 2): the job's abort was checked only between the add-in's busy answers; a call Excel rejects
    was asked again for up to the limit without it."""
    clock["abort_at"] = 3
    t0 = clock["t"]
    with pytest.raises(jobs.JobAborted):
        wait(clock, monkeypatch, tmp_path)
    assert clock["t"] - t0 < 60
