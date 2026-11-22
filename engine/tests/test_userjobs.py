"""xl/userjobs.UserJobs, the runner serverkit._call_user uses (light-fast phase 2 Task 5): the same job as
ExcelHost.call(user_excel=True) gave (test_user_jobs.py, test_job_thread.py and test_excel_turn.py pin it through
UserJobs since Task 6), handed the runner itself, with its own release wait, and the hidden Excel asked for only when
a job uses it (app, on_host). On fakes: no Excel."""
import asyncio
import threading
import time

import pytest

from smartass.core import config
from smartass.core.errors import HostError, HostTimeout
from smartass.xl import jobs, userjobs
from smartass.xl.userjobs import UserJobs


def _no_hidden():
    raise AssertionError("the hidden Excel was asked for")


@pytest.mark.anyio
async def test_a_job_is_handed_the_runner_on_its_own_thread_and_never_asks_for_the_hidden_excel():
    seen = {}

    def job(h):
        seen["h"], seen["thread"] = h, threading.current_thread().name
        return "ok"
    users = UserJobs(hidden=_no_hidden)
    assert await users.call(job, timeout=10, what="fin_one") == "ok"
    assert seen == {"h": users, "thread": "user-job-fin_one"}


@pytest.mark.anyio
async def test_at_the_deadline_the_runner_records_it_and_gives_the_job_up():
    excel = threading.Event()
    users = UserJobs(hidden=_no_hidden)
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
    users = UserJobs(hidden=_no_hidden)
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


def test_the_hidden_excel_is_asked_for_only_when_a_job_uses_it():
    asked = []

    class Hidden:
        app = "the hidden Excel's Application"

        def on_host(self, fn, timeout=None):
            return ("on the host thread", fn(), timeout)

    def hidden():
        asked.append(1)
        return Hidden()
    users = UserJobs(hidden=hidden)
    assert asked == [] and users.com is userjobs.retry_busy
    assert users.app == "the hidden Excel's Application" and asked == [1]
    assert users.on_host(lambda: "built", 7) == ("on the host thread", "built", 7) and asked == [1, 1]


def test_two_jobs_asking_for_the_hidden_excel_at_once_make_one(monkeypatch):
    """UserJobs.on_host / app reach serverkit.host() from the jobs' own threads: two parallel first jobs that both need
    the bridge built must not each start a hidden Excel (one of them would never be shut down)."""
    from smartass.core import serverkit
    from smartass.xl import host
    made = []

    class SlowHost:
        def __init__(self, visible=False):
            time.sleep(0.2)                          # ExcelHost's construction takes a moment
            made.append(self)
    monkeypatch.setattr(host, "ExcelHost", SlowHost)
    monkeypatch.setattr(serverkit, "_host", None)
    got = []
    threads = [threading.Thread(target=lambda: got.append(serverkit.host())) for _ in range(4)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(10)
    assert len(made) == 1 and got == made * 4


def test_the_servers_runner_is_made_without_the_hidden_excel(monkeypatch):
    from smartass.core import serverkit
    monkeypatch.setattr(serverkit, "_users", None)
    monkeypatch.setattr(serverkit, "_host", None)
    monkeypatch.setattr(serverkit, "host", _no_hidden)
    assert isinstance(serverkit.user_jobs(), UserJobs) and serverkit.user_jobs() is serverkit.user_jobs()
    assert serverkit._host is None
