"""xl/userjobs.UserJobs, the runner serverkit._call_user uses (light-fast phase 2 Task 5; since Task 6 the only one):
the same job as the 0.5.0 host gave (test_user_jobs.py, test_job_thread.py and test_excel_turn.py pin it through
UserJobs), handed the runner itself, with its own release wait. On fakes: no Excel."""
import asyncio
import threading
import time

import pytest

from smartass.core import config
from smartass.core.errors import HostError, HostTimeout
from smartass.xl import jobs, userjobs
from smartass.xl.userjobs import UserJobs


@pytest.mark.anyio
async def test_a_job_is_handed_the_runner_on_its_own_thread():
    seen = {}

    def job(h):
        seen["h"], seen["thread"] = h, threading.current_thread().name
        return "ok"
    users = UserJobs()
    assert await users.call(job, timeout=10, what="fin_one") == "ok"
    assert seen == {"h": users, "thread": "user-job-fin_one"}


@pytest.mark.anyio
async def test_at_the_deadline_the_runner_records_it_and_gives_the_job_up():
    excel = threading.Event()
    users = UserJobs()
    try:
        with pytest.raises(HostTimeout, match="Your Excel did not respond within 0s during fin_refresh"):
            await users.call(lambda h: excel.wait(30), timeout=0.3, what="fin_refresh")
        assert users.last_error == "fin_refresh exceeded 0s in the user's Excel"
    finally:
        excel.set()


class _Stuck:
    def __init__(self, excel):
        self.excel = excel

    def __del__(self):
        self.excel.wait(30)


@pytest.mark.anyio
async def test_a_release_that_does_not_return_gives_the_job_up_after_the_runners_release_wait(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(jobs, "ABANDONED_WAIT_S", 0.3)
    monkeypatch.setattr(userjobs, "RELEASE_WAIT_S", 0.3)
    users = UserJobs()
    excel = threading.Event()

    def job(h):
        jobs.hold_turn(4344, what="fin_refresh")
        stuck = _Stuck(excel)
        stuck.me = stuck
        return "answered"
    try:
        assert await users.call(job, timeout=5, what="fin_refresh") == "answered"
        with pytest.raises(HostError, match="fin_refresh.*this request was not started"):
            await users.call(lambda h: jobs.hold_turn(4344, what="this request"), timeout=10, what="fin_analyze")
    finally:
        excel.set()
    deadline = time.monotonic() + 10
    while any(t.name == "user-job-fin_refresh" for t in threading.enumerate()) and time.monotonic() < deadline:
        await asyncio.sleep(0.05)


def test_the_server_has_one_runner(monkeypatch):
    from smartass.core import serverkit
    monkeypatch.setattr(serverkit, "_users", None)
    assert isinstance(serverkit.user_jobs(), UserJobs) and serverkit.user_jobs() is serverkit.user_jobs()
    assert not hasattr(serverkit, "host") and not hasattr(serverkit, "_call")
