import pytest

from smartass.core import timing


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


def test_com_calls_are_counted_per_stage_nested_counted_in_parents(monkeypatch):
    from smartass.xl import jobs
    monkeypatch.setattr(jobs, "_check_alive", lambda: None)
    tm = timing.Timer.start()
    with timing.stage("write"):
        for _ in range(3):
            jobs.com(lambda: 1)
        with timing.stage("inner"):
            jobs.com(lambda: 2)
    jobs.com(lambda: 3)                                   # outside any stage: counted nowhere
    s = tm.summary()
    assert s["com calls.write"] == 4 and s["com calls.write.inner"] == 1
    assert not any(k.startswith("com calls.") and k not in ("com calls.write", "com calls.write.inner") for k in s)


def test_com_without_a_timer_or_in_another_thread_is_not_counted(monkeypatch):
    import threading
    from smartass.xl import jobs
    monkeypatch.setattr(jobs, "_check_alive", lambda: None)
    assert jobs.com(lambda: 5) == 5                       # no timer: no crash
    tm = timing.Timer.start()
    with timing.stage("write"):
        t = threading.Thread(target=lambda: jobs.com(lambda: 1))
        t.start()
        t.join()
    assert not any(k.startswith("com calls.") for k in tm.summary())
