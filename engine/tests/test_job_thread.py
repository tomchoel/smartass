"""Task D (found while looking), Task D2 (the rule): a job's objects - COM proxies among them: its Target's app and workbook, the ranges of
a batch - are released on the job's own STA thread, never where its error is read. An error carried the frames of the
failed job, with their locals, to the event-loop thread (live in a hidden Excel, 21.11.2026: a failed user job's range
was released on MainThread after its thread had ended), and a proxy released on another thread while the apartment
that made it is alive calls that Excel from there and waits for it - forever, when that Excel has stopped
responding. On fakes: no Excel."""
import asyncio
import gc
import threading
import time

import pytest
import pywintypes

from smartass import exceljob
from smartass.core import config
from smartass.xl import host as xlhost
from smartass.xl import hostproc, jobs
from smartass.xl.host import ExcelHost, HostError, HostTimeout
from smartass.xl.hostproc import ComThread
from smartass.xl.inproc import runner


class Proxy:
    """Stands for a COM proxy: records the thread it is released on."""

    def __init__(self, released: list):
        self.released = released

    def __del__(self):
        self.released.append(threading.current_thread().name)


def _fails_in_its_own_frame(released):
    app = Proxy(released)                                      # noqa: F841 - a local of the failing frame
    raise RuntimeError("Excel refused the call")


def _fails_with_a_context(released):
    def inner():
        rng = Proxy(released)                                  # noqa: F841
        raise ValueError("rejected")
    try:
        inner()
    except ValueError:
        raise HostError("Excel was busy") from None            # the context (and its frames) stays on the error


@pytest.mark.parametrize("job", [_fails_in_its_own_frame, _fails_with_a_context], ids=["own frame", "context"])
def test_a_failed_jobs_objects_are_released_on_its_own_thread_not_where_its_error_is_read(job):
    released: list = []
    worker = ComThread(name="user-job-test")
    worker.start()
    fut = worker.submit(lambda: job(released))
    worker.q.put((None, None))
    worker.join(10)
    assert released == ["user-job-test"]
    with pytest.raises((RuntimeError, HostError)):
        fut.result()


# ---------------------------------------------------------------------------------------------------- Task D2: the rule
# No COM object outlives the thread of the job that made it: every job - one that answers, one that fails, one the host
# gave up on - has the objects it made released on its own thread before that thread ends (its error's frames, and
# every reference cycle: the garbage collector run there), and only then are its turns released. Live, 21.11.2026: Excel
# stopped answering between one job's thread ending and the next one's starting, and the next job's thread was found
# releasing a COM object (stack A). The garbage collector is off in these tests, so nothing is collected by chance.


@pytest.fixture
def no_gc():
    gc.collect()
    gc.disable()
    try:
        yield
    finally:
        gc.enable()


def _host() -> ExcelHost:
    h = ExcelHost.__new__(ExcelHost)                 # the user-Excel jobs need no hidden Excel
    h.last_error, h.pid = None, None
    return h


class _Registry:
    def remove(self, pid):
        pass


def _hidden_host() -> ExcelHost:
    """The host with its own long-lived COM thread and a hidden Excel that is only a pid: nothing is started (its
    start is a no-op) and nothing is killed (the tests patch kill_ours)."""
    h = ExcelHost.__new__(ExcelHost)
    h.thread = ComThread()
    h.thread.start()
    h.registry = _Registry()
    h.app = h.wb = None
    h.pid, h.ctime = 4848, 1.0
    h.connected = {}
    h.restarts, h.last_error, h.reset_gen = 0, None, 0
    h._resetting, h._reset_lock = threading.Event(), threading.Lock()
    h._ensure_started = lambda: None
    return h


def _made(released: list) -> None:
    """A job's objects: one only its frame holds, one in a reference cycle (only the garbage collector frees it)."""
    local, cyclic = Proxy(released), Proxy(released)
    cyclic.me = cyclic
    return local


def _answers(released, go):
    def job(hh):
        _made(released)
        return "answered"
    return job


def _fails(released, go):
    def job(hh):
        app = _made(released)                                   # noqa: F841 - a local of the failing frame
        raise RuntimeError("Excel refused the call")
    return job


def _given_up(released, go):
    def job(hh):
        app = _made(released)                                   # noqa: F841
        go.wait(30)                                             # a call into an Excel that does not answer in time
        return "late"
    return job


@pytest.mark.anyio
@pytest.mark.parametrize("make", [_answers, _fails, _given_up], ids=["answers", "fails", "given up on"])
async def test_every_object_a_job_made_is_released_on_its_own_thread(make, no_gc):
    released: list = []
    go = threading.Event()
    try:
        await _host().call(make(released, go), timeout=0.5, what="the job", user_excel=True)
    except HostError:
        pass
    go.set()
    deadline = time.monotonic() + 10
    while len(released) < 2 and time.monotonic() < deadline:      # the job given up on ends when Excel answers
        await asyncio.sleep(0.05)
    gc.collect()                                     # here, on another thread: anything the job left is released here
    assert released == ["user-job-the job", "user-job-the job"]


@pytest.mark.anyio
async def test_a_later_job_on_another_thread_releases_none_of_an_earlier_jobs_objects(no_gc):
    released: list = []
    h = _host()

    def first(hh):
        _made(released)
        return "first"

    def second(hh):
        gc.collect()                                 # the garbage collector, as it runs on any thread by itself
        return "second"
    assert await h.call(first, timeout=10, what="first", user_excel=True) == "first"
    assert await h.call(second, timeout=10, what="second", user_excel=True) == "second"
    assert released == ["user-job-first", "user-job-first"]


class _Stuck:
    """An object whose release does not return until Excel answers (a COM proxy of an Excel that stopped answering:
    stack A)."""

    def __init__(self, excel: threading.Event):
        self.excel = excel

    def __del__(self):
        self.excel.wait(30)


@pytest.mark.anyio
async def test_a_failed_jobs_error_comes_first_when_releasing_its_objects_does_not_return(no_gc):
    """Task D review, Minor 3: the job's outcome is handed over before its objects are released, so a release that
    blocks does not keep the error from the caller (who otherwise got 'did not respond' after the whole timeout)."""
    excel = threading.Event()

    def job(hh):
        app = _Stuck(excel)                          # noqa: F841 - released when the error's frames are
        raise RuntimeError("Excel refused the call")
    t0 = time.monotonic()
    try:
        with pytest.raises(HostError) as e:
            await _host().call(job, timeout=5, what="the job", user_excel=True)
        assert not isinstance(e.value, HostTimeout) and "Excel refused the call" in str(e.value)
        assert time.monotonic() - t0 < 2
    finally:
        excel.set()


@pytest.mark.anyio
async def test_a_release_that_does_not_return_does_not_keep_the_next_job_waiting_silently(monkeypatch, tmp_path,
                                                                                        no_gc):
    """The job's answer comes first; its turn stays held while its objects are released (nothing else works in that
    Excel meanwhile), but a release that does not return within RELEASE_WAIT_S gives the job up: a job waiting for
    that Excel's turn is refused with what holds it, as for a job given up on in Excel. Once the release returns, the
    turn is free."""
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(jobs, "ABANDONED_WAIT_S", 0.3)
    monkeypatch.setattr(xlhost, "RELEASE_WAIT_S", 0.3)
    h = _host()
    excel = threading.Event()

    def job(hh):
        jobs.hold_turn(4343, what="fin_refresh")
        stuck = _Stuck(excel)
        stuck.me = stuck
        return "answered"

    def next_job(hh):
        jobs.hold_turn(4343, what="this request")
        return "ran"
    t0 = time.monotonic()
    try:
        assert await h.call(job, timeout=5, what="fin_refresh", user_excel=True) == "answered"
        assert time.monotonic() - t0 < 2
        with pytest.raises(HostError) as e:
            await h.call(next_job, timeout=10, what="fin_analyze", user_excel=True)
        assert not isinstance(e.value, HostTimeout)
        assert "fin_refresh" in str(e.value) and "this request was not started" in str(e.value)
    finally:
        excel.set()
    deadline = time.monotonic() + 10
    while True:                                      # the release returned: the turn is free again
        try:
            assert await h.call(next_job, timeout=5, what="fin_analyze", user_excel=True) == "ran"
            break
        except HostError:
            if time.monotonic() > deadline:
                raise
            await asyncio.sleep(0.1)


@pytest.mark.anyio
async def test_a_hidden_jobs_release_that_does_not_return_restarts_the_hidden_excel(monkeypatch, no_gc):
    """Task D2 review, Important 1: the hidden Excel's jobs (the EPM pulls) run one after another on the host's one
    long-lived thread. A release after a job's answer that does not return wedged that thread: every later call queued
    for QUEUE_TIMEOUT_S and failed, with no reset. Now a release that has not returned RELEASE_WAIT_S after the answer
    restarts the hidden Excel as a job that does not answer does: the kill makes the blocked release return, and the
    next call runs."""
    monkeypatch.setattr(xlhost, "RELEASE_WAIT_S", 0.3)
    monkeypatch.setattr(config, "QUEUE_TIMEOUT_S", 3)
    monkeypatch.setattr(config, "RESET_WAIT_S", 3)
    excel = threading.Event()                        # set: that Excel answers again (here: it was killed)
    killed: list = []

    def kill(pid, ctime):                            # the reset's kill of the engine's own Excel
        killed.append(pid)
        excel.set()
        return True
    monkeypatch.setattr(xlhost, "kill_ours", kill)
    monkeypatch.setattr(xlhost, "windows_of", lambda pid: [])
    h = _hidden_host()

    def pull(hh):
        stuck = _Stuck(excel)                        # released by the collector at the job's end
        stuck.me = stuck
        return "pulled"
    try:
        t0 = time.monotonic()
        assert await h.call(pull, timeout=5, what="an EPM pull") == "pulled"
        assert time.monotonic() - t0 < 2             # the answer first
        assert await h.call(lambda hh: "next", timeout=5, what="the next pull") == "next"
        assert killed == [4848] and h.pid is None and h.restarts == 1
    finally:
        excel.set()
        h.thread.q.put((None, None))


def test_the_release_after_a_job_handed_to_the_hidden_excel_from_a_user_job_is_watched_too(monkeypatch, no_gc):
    """on_host (a user job's work in the hidden Excel: the bridge add-in's build, a new analysis file) runs on the same
    host thread: its release is watched as ExcelHost.call's."""
    monkeypatch.setattr(xlhost, "RELEASE_WAIT_S", 0.3)
    monkeypatch.setattr(config, "RESET_WAIT_S", 3)
    excel = threading.Event()
    killed: list = []

    def kill(pid, ctime):
        killed.append(pid)
        excel.set()
        return True
    monkeypatch.setattr(xlhost, "kill_ours", kill)
    monkeypatch.setattr(xlhost, "windows_of", lambda pid: [])
    h = _hidden_host()

    def build(hh=None):
        stuck = _Stuck(excel)
        stuck.me = stuck
        return "built"
    try:
        assert h.on_host(build, timeout=5) == "built"
        assert h.on_host(lambda: "next", timeout=5) == "next"
        assert killed == [4848]
    finally:
        excel.set()
        h.thread.q.put((None, None))


@pytest.mark.anyio
async def test_a_job_that_returns_a_com_object_fails_on_its_own_thread():
    """A job's result is plain data: a COM object in it would be released on the thread that reads the result."""
    import pythoncom

    def job(hh):
        d = pythoncom.CoCreateInstance("Scripting.Dictionary", None, pythoncom.CLSCTX_INPROC_SERVER,
                                       pythoncom.IID_IDispatch)
        return {"workbook": "x.xlsx", "checks": [("ok", d)]}
    with pytest.raises(HostError) as e:
        await _host().call(job, timeout=10, what="the job", user_excel=True)
    assert "COM object" in str(e.value)



# ------------------------------------------------------------------------------------------- Task D3: what is enforced
# gc.collect() is process-wide: at the end of job X it released, on X's thread, the cycles a job Y still running on
# another thread had left (a user job and the hidden Excel's EPM pulls overlap), and Python's own collector runs on
# whatever thread allocates. So a job's end collects only when no other COM job is running, and the cycles that held
# a job's COM objects are broken where they are made: those objects go by reference counting, on the job's own thread.
# The automatic collector stays ON in these tests.


class _GcCalls:
    """hostproc's gc with each collect recorded by the thread that ran it."""

    def __init__(self):
        self.threads: list = []

    def collect(self, *a):
        self.threads.append(threading.current_thread().name)
        return gc.collect(*a)


async def _ended(name: str) -> None:
    deadline = time.monotonic() + 10
    while any(t.name == name for t in threading.enumerate()) and time.monotonic() < deadline:
        await asyncio.sleep(0.02)


@pytest.mark.anyio
async def test_a_jobs_end_collects_only_when_no_other_com_job_is_running(monkeypatch):
    calls = _GcCalls()
    monkeypatch.setattr(hostproc, "gc", calls)
    assert gc.isenabled()
    h = _host()
    a_runs, a_go = threading.Event(), threading.Event()

    def a(hh):
        a_runs.set()
        a_go.wait(10)
        return "a"
    first = asyncio.ensure_future(h.call(a, timeout=20, what="A", user_excel=True))
    try:
        assert await asyncio.to_thread(a_runs.wait, 10)
        assert await h.call(lambda hh: "b", timeout=10, what="B", user_excel=True) == "b"
        await _ended("user-job-B")                   # B's end has run, while A runs
        assert calls.threads == []
    finally:
        a_go.set()
    assert await first == "a"
    await _ended("user-job-A")
    assert calls.threads == ["user-job-A"]           # A's end, alone


class _Fails:
    """A job whose fail hands the error back as it is (exceljob._Job.fail with nothing to add)."""

    @staticmethod
    def fail(e):
        return e


def _busy() -> pywintypes.com_error:
    return pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)


def _read_rejected(released) -> None:
    """A verb's read of a sheet that Excel rejected (busy): runner._readable sends it on up; the caller asks again,
    and the error is dropped inside the job."""
    def read():
        rng = Proxy(released)                                   # noqa: F841 - a local of the frame the error holds
        try:
            raise _busy()
        except pywintypes.com_error as e:
            runner._readable(e)
    try:
        read()
    except pywintypes.com_error:
        pass


def _batch_failed(released) -> None:
    """A batch's error raised through exceljob._raise and handled inside the job."""
    def draw():
        t = Proxy(released)                                     # noqa: F841
        try:
            raise RuntimeError("Excel operation 3 failed")
        except RuntimeError as e:
            exceljob._raise(_Fails(), e)
    try:
        draw()
    except RuntimeError:
        pass


@pytest.mark.anyio
async def test_two_overlapping_jobs_objects_are_released_on_their_own_threads_with_the_collector_on():
    """Each job's error is dropped inside it while the other runs; nothing either made waits for a collector."""
    assert gc.isenabled()
    released: list = []
    h = _host()
    a_done, a_go = threading.Event(), threading.Event()

    def a(hh):
        _read_rejected(released)
        a_done.set()
        a_go.wait(10)
        return "a"

    def b(hh):
        _batch_failed(released)
        return "b"
    first = asyncio.ensure_future(h.call(a, timeout=20, what="A", user_excel=True))
    try:
        assert await asyncio.to_thread(a_done.wait, 10)
        assert await h.call(b, timeout=10, what="B", user_excel=True) == "b"
        await _ended("user-job-B")
    finally:
        a_go.set()
    assert await first == "a"
    await _ended("user-job-A")
    assert sorted(released) == ["user-job-A", "user-job-B"]


def _aborted_batch(released) -> None:
    """The COM executor's batch of a job given up on: execute raises JobAborted before its first operation."""
    token = jobs.JobToken("the job")
    token.aborted = True
    jobs.bind_job(token)
    try:
        runner.execute(Proxy(released), {"ops": [{"k": "mark", "name": "x"}]}, 5)
    finally:
        jobs.unbind_job()


class _RefusingExcel(Proxy):
    def RegisterXLL(self, path):
        raise RuntimeError("Excel refused the add-in")


def _register_refused(released) -> None:
    runner._register_xll(_RefusingExcel(released), "x.xll", 5)


@pytest.mark.parametrize("site", [_read_rejected, _batch_failed, _aborted_batch, _register_refused],
                         ids=["runner._readable", "exceljob._raise", "runner.execute (aborted)", "runner._register_xll"])
def test_an_error_raised_where_it_was_kept_leaves_no_reference_cycle(site, no_gc):
    """An exception kept in a local (or a parameter, or a dict) of a frame its own traceback holds makes a cycle with
    that frame and every object its frames hold: only a collector frees them, on whatever thread it runs."""
    released: list = []
    try:
        site(released)
    except Exception:   # noqa: BLE001
        pass
    assert released == ["MainThread"]                # freed as the error was dropped, with no collector
