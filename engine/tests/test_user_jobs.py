"""What a job in the USER's Excel goes through (ExcelHost.call(user_excel=True) in 0.5.0; xl/userjobs.UserJobs since
the split of light-fast phase 2 Task 5), step by step, pinned on fakes before the split and unchanged after it: the
job's own STA thread, the deadline and what happens at it (the job given up on, nothing closed, no release watch),
which errors become which texts. Since Task 6 deleted the hidden Excel these run through UserJobs, the runner
serverkit._call_user uses (only the entry point changed). The turn, the given-up mark and the release on the job's
own thread are pinned in test_excel_turn.py and test_job_thread.py."""
import asyncio
import threading
import time

import pytest
import pywintypes

from smartass.core import config
from smartass.xl import jobs
from smartass.core.errors import HostError, HostTimeout, InputError
from smartass.xl.userjobs import UserJobs


def _host() -> UserJobs:
    """The runner of jobs in the user's Excel (serverkit.user_jobs makes one per server)."""
    return UserJobs()


async def _ended(prefix: str) -> None:
    deadline = time.monotonic() + 10
    while any(t.name.startswith(prefix) for t in threading.enumerate()) and time.monotonic() < deadline:
        await asyncio.sleep(0.02)


@pytest.mark.anyio
async def test_a_job_runs_on_an_sta_thread_of_its_own_named_for_it_that_ends_after_it():
    seen = {}

    def job(hh):
        import pythoncom
        t = threading.current_thread()
        seen["name"], seen["daemon"] = t.name, t.daemon
        seen["token"] = jobs._job.token.what
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)        # already: S_FALSE, no error
        pythoncom.CoUninitialize()
        seen["handed"] = hh
        return {"ok": 1}
    h = _host()
    assert await h.call(job, timeout=10, what="fin_one") == {"ok": 1}
    assert seen["name"] == "user-job-fin_one" and seen["daemon"] and seen["token"] == "fin_one"
    assert seen["name"] != threading.current_thread().name and seen["handed"] is h
    await _ended("user-job-fin_one")
    assert not any(t.name == "user-job-fin_one" for t in threading.enumerate())


@pytest.mark.anyio
async def test_parallel_jobs_run_at_the_same_time_on_their_own_threads():
    both = threading.Barrier(2, timeout=10)

    def job(hh):
        both.wait()                                  # each waits for the other: only if they run at once
        return threading.current_thread().name
    h = _host()
    got = await asyncio.gather(h.call(job, timeout=10, what="A"),
                               h.call(job, timeout=10, what="B"))
    assert got == ["user-job-A", "user-job-B"]


@pytest.mark.anyio
async def test_at_the_deadline_the_job_is_given_up_on_nothing_is_closed_and_no_release_is_watched(monkeypatch):
    excel = threading.Event()
    after = {}

    def stuck(hh):
        excel.wait(30)                               # a COM call into an Excel that does not answer
        try:
            jobs._check_alive()                      # its next COM call
        except jobs.JobAborted as e:
            after["aborted"] = str(e)
        return "late"
    h = _host()
    t0 = time.monotonic()
    try:
        with pytest.raises(HostTimeout) as e:
            await h.call(stuck, timeout=0.4, what="fin_refresh")
        assert time.monotonic() - t0 < 3
        assert str(e.value) == ("Your Excel did not respond within 0s during fin_refresh (a dialog, edit mode, or a "
                                "large refresh). Nothing was closed. Check Excel for an open dialog, press Esc/Enter, "
                                "and ask again.")
        assert h.last_error == "fin_refresh exceeded 0s in the user's Excel"
        assert not any(t.name.startswith("release-watch") for t in threading.enumerate())
    finally:
        excel.set()
    await _ended("user-job-fin_refresh")
    assert after["aborted"].startswith("This request was abandoned after a timeout")


@pytest.mark.anyio
async def test_without_a_timeout_a_job_has_the_user_excel_deadline(monkeypatch):
    monkeypatch.setattr(config, "USER_CALL_TIMEOUT_S", 0.3)
    excel = threading.Event()
    try:
        with pytest.raises(HostTimeout, match="within 0s during xl_tidy"):
            await _host().call(lambda hh: excel.wait(30), what="xl_tidy")
    finally:
        excel.set()


def _com(hresult, excepinfo=None):
    return pywintypes.com_error(hresult, "x", excepinfo, None)


@pytest.mark.anyio
@pytest.mark.parametrize("hresult", [-2147023174, -2147023170, -2147417848])
async def test_an_excel_that_disconnected_is_said_so_and_counted(hresult):
    def job(hh):
        raise _com(hresult)
    h = _host()
    with pytest.raises(HostError) as e:
        await h.call(job, timeout=10, what="fin_x")
    assert type(e.value) is HostError
    assert str(e.value) == ("Your Excel stopped answering or closed during fin_x; smartass closed nothing. If it "
                            "closed, open it again; then retry.")
    assert h.restarts == 1 and h.last_error.startswith("Excel disconnected: ")


@pytest.mark.anyio
@pytest.mark.parametrize("error", [_com(-2147418111), _com(-2147417846),
                                   _com(-2147352567, (0, None, None, None, 0, -2146777998))],
                         ids=["call rejected", "retry later", "busy recalculating"])
async def test_a_busy_excel_is_said_in_words_with_no_dialog_scan(error):
    def job(hh):
        raise error
    with pytest.raises(HostError) as e:
        await _host().call(job, timeout=10, what="fin_x")
    assert str(e.value) == ("Excel was busy during fin_x - probably a cell is being edited or a dialog is open. "
                            "Press Enter or Esc in Excel and ask again.")
    assert not isinstance(e.value, AttributeError)


@pytest.mark.anyio
async def test_another_com_error_is_named_with_the_request():
    def job(hh):
        raise _com(-2146827284)
    with pytest.raises(HostError) as e:
        await _host().call(job, timeout=10, what="fin_x")
    assert type(e.value) is HostError and str(e.value).startswith("COM error during fin_x: (-2146827284, 'x'")


@pytest.mark.anyio
async def test_an_unexpected_error_is_readable_with_its_type_and_the_restart_advice():
    def job(hh):
        raise KeyError("Q9")
    with pytest.raises(HostError) as e:
        await _host().call(job, timeout=10, what="fin_x")
    assert str(e.value) == ("Unexpected KeyError during fin_x: 'Q9'. Retry once; if it repeats, run "
                            "fin_status(restart=True).")


@pytest.mark.anyio
@pytest.mark.parametrize("error", [HostError("Excel said no"), InputError("bad target"),
                                   jobs.JobAborted("This request was abandoned")])
async def test_the_engines_own_errors_pass_as_they_are(error):
    def job(hh):
        raise error
    with pytest.raises(type(error)) as e:
        await _host().call(job, timeout=10, what="fin_x")
    assert str(e.value) == str(error) and not isinstance(e.value, HostTimeout)


@pytest.mark.anyio
async def test_a_job_that_answers_has_its_release_watched_and_its_turns_freed(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "SVKIT", tmp_path)

    def job(hh):
        jobs.hold_turn(5151, what="fin_x")
        return "ok"
    h = _host()
    assert await h.call(job, timeout=10, what="fin_x") == "ok"
    await _ended("user-job-fin_x")
    jobs.bind_job()
    try:
        assert jobs.hold_turn(5151, wait=False) is True                    # freed on the job's own thread
    finally:
        jobs.release_turns()
