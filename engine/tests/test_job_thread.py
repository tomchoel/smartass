"""Task D (found while looking): a job's objects - COM proxies among them: its Target's app and workbook, the ranges of
a batch - are released on the job's own STA thread, never where its error is read. An error carried the frames of the
failed job, with their locals, to the event-loop thread (live in a hidden Excel, 21.11.2026: a failed user job's range
was released on MainThread after its thread had ended), and a proxy released on another thread while the apartment
that made it is alive calls that Excel from there and waits for it - forever, when that Excel has stopped
responding. On fakes: no Excel."""
import threading

import pytest

from smartass.xl.host import HostError
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
