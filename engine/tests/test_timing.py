import pytest

from smartview_mcp import timing


@pytest.fixture(autouse=True)
def _no_timer_left():
    yield
    timing._local.timer = timing._local.pending_notify = None


def test_say_reaches_the_client_of_the_running_job():
    notes = []
    timing.Timer.note_notify(notes.append)
    timing.Timer.start()
    timing.say("EPM gave an empty response; retrying in 3 s")
    assert notes == ["EPM gave an empty response; retrying in 3 s"]


def test_say_without_a_job_or_with_a_broken_callback_is_silent():
    timing._local.timer = None
    timing.say("nobody listens")
    timing.Timer.note_notify(lambda t: 1 / 0)
    timing.Timer.start()
    timing.say("still fine")
