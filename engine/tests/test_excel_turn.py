"""excel_turn: the machine-wide queue for the user's Excel works ACROSS processes (two agents/servers)."""
import subprocess
import sys
import time

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from smartview_mcp import userexcel  # noqa: E402

HOLDER = r"""
import sys, time
sys.path.insert(0, sys.argv[1])
from smartview_mcp import userexcel
with userexcel.excel_turn("holder"):
    print("held", flush=True)
    time.sleep(float(sys.argv[2]))
"""


def test_second_process_waits_for_the_turn():
    root = __file__.rsplit("\\tests", 1)[0]
    p = subprocess.Popen([sys.executable, "-c", HOLDER, root, "4"], stdout=subprocess.PIPE, text=True)
    assert p.stdout.readline().strip() == "held"
    t0 = time.monotonic()
    with userexcel.excel_turn("waiter") as waited:
        elapsed = time.monotonic() - t0
    p.wait(10)
    assert elapsed > 2.5 and waited > 2.5, elapsed          # really waited for the other process


def test_turns_are_per_excel_and_released_at_job_end():
    root = __file__.rsplit("\\tests", 1)[0]
    p = subprocess.Popen([sys.executable, "-c", HOLDER.replace('excel_turn("holder")', 'excel_turn("holder", pid=111)'),
                          root, "3"], stdout=subprocess.PIPE, text=True)
    assert p.stdout.readline().strip() == "held"
    userexcel.bind_job()
    assert userexcel.hold_turn(111, wait=False) is False     # Excel 111 busy (another agent)
    assert userexcel.hold_turn(222, wait=False) is True      # a different Excel is free: parallel
    assert userexcel.hold_turn(222, wait=False) is True      # re-entrant within the same job
    userexcel.release_turns()
    p.wait(10)
    userexcel.bind_job()
    assert userexcel.hold_turn(111, wait=False) is True      # free again after the other agent finished
    userexcel.release_turns()


def test_turn_is_free_again_and_reentrant_sequentially():
    for _ in range(3):
        t0 = time.monotonic()
        with userexcel.excel_turn("again"):
            pass
        assert time.monotonic() - t0 < 1.0
