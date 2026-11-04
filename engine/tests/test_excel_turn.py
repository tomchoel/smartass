"""Turn: the machine-wide queue for the user's Excel works ACROSS processes (two agents/servers)."""
import subprocess
import sys
import time

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from smartass.xl import jobs  # noqa: E402

HOLDER = r"""
import sys, time
sys.path.insert(0, sys.argv[1])
from smartview_mcp.excel import jobs
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
