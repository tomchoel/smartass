"""Always warm: no idle close, keep-alive, yield to another smartass process (real _loop thread, fake session)."""
import os
import queue
import threading
import time

import pytest

from smartview_mcp import config
from smartview_mcp import drill_browser as db


class FakeCtx:
    def close(self):
        pass


def _session(tmp_path, lock=None):
    s = db._Session()
    s.ctx, s.page = FakeCtx(), object()
    s.plock = lock
    s.profile = tmp_path / "prof"
    s.profile.mkdir(exist_ok=True)
    s.started_at = time.time() - 100
    s.last_alive = time.monotonic()
    s.last_used = time.monotonic() - 1000                # unused for a long time
    return s


def _run(s, keepalive=None, idle=0.02):
    q = queue.Queue()
    t = threading.Thread(target=db.DrillBrowser._loop, args=(q, s, idle, keepalive), daemon=True)
    t.start()
    return q, t


def _stop(q, t):
    q.put((None, None))
    t.join(5)


@pytest.fixture
def warm(monkeypatch):
    monkeypatch.setattr(config, "DRILL_KEEP_WARM", True)
    monkeypatch.setattr(config, "DRILL_KEEPALIVE_S", 0.05)
    monkeypatch.setattr(db, "_kill_driver", lambda d: [])


def test_keep_warm_never_closes_and_runs_keepalive(warm, tmp_path):
    s = _session(tmp_path)
    calls = []

    def bad(sess):
        calls.append(1)
        raise RuntimeError("boom")                      # a failing keep-alive must not kill the worker
    q, t = _run(s, bad)
    time.sleep(0.4)
    assert s.ctx is not None and len(calls) >= 2 and t.is_alive()
    done = []
    fut_q = db.cf.Future()
    q.put((lambda sess: done.append(1) or 7, fut_q))    # the worker still takes jobs
    assert fut_q.result(5) == 7
    _stop(q, t)


def test_keepalive_not_run_while_requests_in_flight(warm, tmp_path):
    s = _session(tmp_path)
    s.inflight[1] = (None, "t", db.cf.Future(), time.monotonic() + 60)
    calls = []
    monkey = db.DrillBrowser._collect
    db.DrillBrowser._collect = staticmethod(lambda sess: None)
    try:
        q, t = _run(s, lambda sess: calls.append(1))
        time.sleep(0.3)
        assert not calls
        s.inflight.clear()
        _stop(q, t)
    finally:
        db.DrillBrowser._collect = monkey


def test_wanted_flag_makes_idle_worker_yield(warm, tmp_path):
    lock = db.ProfileLock(tmp_path / "lockprof")
    assert lock.acquire()
    s = _session(tmp_path, lock)
    flag = s.profile / ".wanted"
    flag.write_text("1")
    os.utime(flag, (time.time() + 5, time.time() + 5))
    q, t = _run(s, lambda sess: None)
    for _ in range(200):
        if s.ctx is None:
            break
        time.sleep(0.02)
    assert s.ctx is None and not lock.held and not flag.exists()
    _stop(q, t)


def test_recently_used_owner_does_not_yield(warm, tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DRILL_YIELD_IDLE_S", 120)
    s = _session(tmp_path)
    s.last_used = time.monotonic() - 5
    flag = s.profile / ".wanted"
    flag.write_text("1")
    os.utime(flag, (time.time() + 5, time.time() + 5))
    q, t = _run(s, lambda sess: None)
    time.sleep(0.2)
    assert s.ctx is not None and flag.exists()
    s.last_used = time.monotonic() - 121                 # now unused for long enough
    for _ in range(200):
        if s.ctx is None:
            break
        time.sleep(0.02)
    assert s.ctx is None
    _stop(q, t)


def test_a_job_counts_as_use(warm, tmp_path):
    s = _session(tmp_path)
    q, t = _run(s, lambda sess: None)
    fut = db.cf.Future()
    q.put((lambda sess: 1, fut))
    fut.result(5)
    assert time.monotonic() - s.last_used < 5
    _stop(q, t)


def test_old_wanted_flag_is_ignored(warm, tmp_path):
    s = _session(tmp_path)
    flag = s.profile / ".wanted"
    flag.write_text("1")
    os.utime(flag, (time.time() - 1000, time.time() - 1000))
    q, t = _run(s, lambda sess: None)
    time.sleep(0.2)
    assert s.ctx is not None
    _stop(q, t)


def test_keep_warm_off_closes_when_idle(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "DRILL_KEEP_WARM", False)
    monkeypatch.setattr(db, "_kill_driver", lambda d: [])
    s = _session(tmp_path)
    q, t = _run(s, lambda sess: None)
    for _ in range(100):
        if s.ctx is None:
            break
        time.sleep(0.02)
    assert s.ctx is None
    _stop(q, t)


def test_drillbusy_writes_wanted_flag(tmp_path):
    owner = db.ProfileLock(tmp_path / "prof")
    assert owner.acquire()
    b = db.DrillBrowser(profile=tmp_path / "prof")
    try:
        fut = b.submit(lambda s: b._start(s))
        with pytest.raises(db.DrillBusy):
            fut.result(5)
        assert (tmp_path / "prof" / ".wanted").exists()
    finally:
        with b._lock:
            b.q.put((None, None))
        b.thread.join(5)
        owner.release()


def test_warmup_drillbusy_writes_no_flag_and_says_when(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DRILL_KEEP_WARM", True)
    assert db.KEEPALIVE_WAIT_S == 10
    owner = db.ProfileLock(tmp_path / "prof")
    assert owner.acquire()
    b = db.DrillBrowser(profile=tmp_path / "prof")
    try:
        def warm(s):
            s.warmup = True
            try:
                b._ensure(s)
            finally:
                s.warmup = False
        with pytest.raises(db.DrillBusy, match="idle minutes"):
            b.submit(warm).result(5)
        assert not (tmp_path / "prof" / ".wanted").exists()
        with pytest.raises(db.DrillBusy, match="lets go after 2 idle minutes"):
            b.submit(lambda s: b._start(s)).result(5)
    finally:
        with b._lock:
            b.q.put((None, None))
        b.thread.join(5)
        owner.release()


def test_keepalive_resigns_in_when_session_gone(tmp_path):
    b = db.DrillBrowser(profile=tmp_path / "prof2")
    try:
        s = b.s
        ensured = []
        b._ensure = lambda sess: ensured.append(sess.ready)

        class P:
            pass
        s.page = P()
        s.ready = True
        monkey = db.DrillBrowser._fire
        # EPM answers not-200 (signed out): ready flips off and _ensure runs again
        def fire(sess, page, url, opts, t):
            f = db.cf.Future()
            f.set_result((401, "<html>login.microsoftonline</html>"))
            return f
        db.DrillBrowser._fire = staticmethod(fire)
        try:
            P.url = "https://login.microsoftonline.com/x"
            b._keepalive(s)
        finally:
            db.DrillBrowser._fire = monkey
        assert len(ensured) == 2 and s.ready is False
    finally:
        with b._lock:
            b.q.put((None, None))
        b.thread.join(5)


def test_a_second_server_takes_the_next_profile_slot(tmp_path, monkeypatch):
    from smartview_mcp import config, drill_browser as db
    monkeypatch.setattr(config, "DRILL_PROFILE_SLOTS", 3)
    base = tmp_path / "drill_edge"
    other = db.ProfileLock(base)                       # another smartass process holds slot 1
    assert other.acquire()
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    b.profile = b.base_profile = base
    b.plock = db.ProfileLock(base)
    assert b._take_slot() and b.profile == tmp_path / "drill_edge_2" and b.plock.held
    third = db.DrillBrowser.__new__(db.DrillBrowser)
    third.profile = third.base_profile = base
    third.plock = db.ProfileLock(base)
    assert third._take_slot() and third.profile == tmp_path / "drill_edge_3"
    fourth = db.DrillBrowser.__new__(db.DrillBrowser)
    fourth.profile = fourth.base_profile = base
    fourth.plock = db.ProfileLock(base)
    assert not fourth._take_slot() and fourth.profile == base   # all slots taken: DrillBusy as before
    for x in (other, b.plock, third.plock):
        x.release()
