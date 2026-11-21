"""Turn: the machine-wide queue for the user's Excel works ACROSS processes (two agents/servers)."""
import subprocess
import sys
import threading
import time

import pytest

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from smartass.core import config  # noqa: E402
from smartass.xl import jobs  # noqa: E402
from smartass.xl.host import ExcelHost, HostError, HostTimeout  # noqa: E402

HOLDER = r"""
import sys, time
sys.path.insert(0, sys.argv[1])
from smartass.xl import jobs
tt = jobs.Turn(int(sys.argv[3]))
tt.acquire("holder")
print("held", flush=True)
time.sleep(float(sys.argv[2]))
tt.close()
"""


def _holder(seconds: str, pid: int = 0):
    root = __file__.rsplit("\\tests", 1)[0]
    p = subprocess.Popen([sys.executable, "-c", HOLDER, root, seconds, str(pid)], stdout=subprocess.PIPE, text=True)
    assert p.stdout.readline().strip() == "held"
    return p


def test_second_process_waits_for_the_turn():
    p = _holder("4")
    t0 = time.monotonic()
    tt = jobs.Turn(0)
    waited = tt.acquire("waiter")
    tt.close()
    elapsed = time.monotonic() - t0
    p.wait(10)
    assert elapsed > 2.5 and waited > 2.5, elapsed          # really waited for the other process


def test_turns_are_per_excel_and_released_at_job_end():
    p = _holder("3", pid=111)
    jobs.bind_job()
    assert jobs.hold_turn(111, wait=False) is False     # Excel 111 busy (another agent)
    assert jobs.hold_turn(222, wait=False) is True      # a different Excel is free: parallel
    assert jobs.hold_turn(222, wait=False) is True      # re-entrant within the same job
    jobs.release_turns()
    p.wait(10)
    jobs.bind_job()
    assert jobs.hold_turn(111, wait=False) is True      # free again after the other agent finished
    jobs.release_turns()


def test_turn_is_free_again_and_reentrant_sequentially():
    for _ in range(3):
        t0 = time.monotonic()
        tt = jobs.Turn(0)
        tt.acquire("again")
        tt.close()
        assert time.monotonic() - t0 < 1.0


@pytest.mark.anyio
async def test_a_job_given_up_on_does_not_keep_the_next_job_waiting_for_that_excels_turn(monkeypatch, tmp_path):
    """Task D (live, 21.11.2026: stack B): the host gave up on a user-Excel job whose Excel did not answer, but that
    job's thread still held the Excel's turn, so the next job in that Excel waited silently for the turn until it was
    given up on too. Now the next job is refused at once with what holds the turn; once the job given up on has
    ended, the turn is free again."""
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(jobs, "TURN_WAIT_S", 60)
    h = ExcelHost.__new__(ExcelHost)                 # the user-Excel jobs need no hidden Excel
    h.last_error, h.pid = None, None
    excel_answers = threading.Event()

    def stuck(hh):
        jobs.hold_turn(4242, what="fin_refresh")
        excel_answers.wait(30)                       # a COM call into an Excel that does not answer
        return "late"

    def next_job(hh):
        jobs.hold_turn(4242, what="this request")
        return "ran"

    with pytest.raises(HostTimeout):
        await h.call(stuck, timeout=0.3, what="fin_refresh", user_excel=True)
    t0 = time.monotonic()
    try:
        with pytest.raises(HostError) as e:
            await h.call(next_job, timeout=20, what="fin_analyze", user_excel=True)
        assert not isinstance(e.value, HostTimeout)
        assert "fin_refresh" in str(e.value) and "this request was not started" in str(e.value)
        assert time.monotonic() - t0 < jobs.ABANDONED_WAIT_S + 2
    finally:
        excel_answers.set()
    deadline = time.monotonic() + 10
    while True:                                      # the job given up on ends and frees the turn
        try:
            assert await h.call(next_job, timeout=5, what="fin_analyze", user_excel=True) == "ran"
            break
        except HostError:
            if time.monotonic() > deadline:
                raise
            time.sleep(0.1)
