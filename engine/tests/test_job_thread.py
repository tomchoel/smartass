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

from smartass.core import config
from smartass.xl import host as xlhost
from smartass.xl import jobs
from smartass.xl.host import ExcelHost, HostError, HostTimeout
from smartass.xl.hostproc import ComThread


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
