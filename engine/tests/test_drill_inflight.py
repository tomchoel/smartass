"""DrillBrowser.request: fetches that run in the tab while the worker stays free (the real worker thread, fake tabs)."""
import threading

import pytest

from smartview_mcp import drill_browser as db


class BoxPage:
    """Plays window.__smartass: _FIRE_JS records the fetch, _TAKE_JS hands over what answer() put in the box."""

    def __init__(self):
        self.tok, self.fired, self.done, self.lock = "t1", {}, [], threading.Lock()
        self.broken = None

    def evaluate(self, js, arg=None):
        if self.broken:
            raise self.broken
        with self.lock:
            if js is db._FIRE_JS:
                rid, url, opts, ms = arg
                self.fired[rid] = (url, opts, ms)
                return self.tok
            if js is db._TAKE_JS:
                d, self.done = self.done, []
                return [self.tok, d]
        raise AssertionError(js)

    def answer(self, url, status=200, text="ok"):
        with self.lock:
            rid = next(r for r, (u, _, _) in self.fired.items() if u == url)
            del self.fired[rid]
            self.done.append([rid, status, text])

    def navigate(self):
        with self.lock:
            self.tok, self.done = "t2", []


@pytest.fixture
def browser(tmp_path):
    b = db.DrillBrowser(profile=tmp_path / "prof")
    yield b
    with b._lock:
        b.q.put((None, None))
    b.thread.join(5)


def _wait_fired(page, n):
    for _ in range(500):
        if len(page.fired) >= n:
            return
        threading.Event().wait(0.01)
    raise AssertionError(f"only {len(page.fired)} fired")


def test_many_requests_are_in_flight_at_once_and_answer_out_of_order(browser):
    page = BoxPage()
    futs = [browser.request(lambda s, i=i: (page, f"/u{i}", {"method": "GET"}), 30) for i in range(5)]
    _wait_fired(page, 5)                                  # all five started before any answer: the worker is free
    assert all(not f.done() for f in futs)
    for i in (3, 0, 4, 1, 2):
        page.answer(f"/u{i}", text=f"a{i}")
    assert [f.result(5) for f in futs] == [(200, f"a{i}") for i in range(5)]
    assert browser.s.inflight == {}


def test_options_and_the_abort_time_reach_the_fetch(browser):
    page = BoxPage()
    f = browser.request(lambda s: (page, "/x", {"method": "POST", "body": "b"}), 2.5)
    _wait_fired(page, 1)
    url, opts, ms = page.fired[next(iter(page.fired))]
    assert (url, opts, ms) == ("/x", {"method": "POST", "body": "b"}, 2500)
    assert "AbortSignal.timeout(ms)" in db._FIRE_JS and "credentials: 'include'" in db._FIRE_JS
    page.answer("/x")
    assert f.result(5) == (200, "ok")


def test_a_navigated_tab_fails_its_requests(browser):
    page = BoxPage()
    f = browser.request(lambda s: (page, "/x", {}), 30)
    _wait_fired(page, 1)
    page.navigate()
    with pytest.raises(db.DrillError, match="navigated away"):
        f.result(5)


def test_a_closed_tab_fails_its_requests_with_a_redacted_reason(browser):
    page = BoxPage()
    f = browser.request(lambda s: (page, "/x", {}), 30)
    _wait_fired(page, 1)
    page.broken = RuntimeError("Target closed https://x/p?sso_token=SECRET")
    with pytest.raises(db.DrillError, match=r"tab is gone") as e:
        f.result(5)
    assert "SECRET" not in str(e.value)


def test_no_answer_in_time_fails(browser, monkeypatch):
    page = BoxPage()
    f = browser.request(lambda s: (page, "/x", {}), -4.9)          # deadline = now + timeout + 5 s: just past
    with pytest.raises(db.DrillError, match="did not answer in time"):
        f.result(5)


def test_a_failing_prep_fails_the_request(browser):
    def prep(s):
        raise db.DrillError("sign-in failed")
    with pytest.raises(db.DrillError, match="sign-in failed"):
        browser.request(prep, 30).result(5)


def test_a_request_cancelled_before_it_starts_is_never_fired(browser):
    page, gate = BoxPage(), threading.Event()
    browser.submit(lambda s: gate.wait(5))                # the worker is busy
    f = browser.request(lambda s: (page, "/x", {}), 30)
    assert f.cancel()
    gate.set()
    f.job.result(5)
    assert page.fired == {}


def test_closing_the_browser_fails_the_requests_in_flight(browser, monkeypatch):
    page = BoxPage()
    f = browser.request(lambda s: (page, "/x", {}), 30)
    _wait_fired(page, 1)
    monkeypatch.setattr(db, "_kill_driver", lambda d: [])
    browser.submit(db.DrillBrowser._close).result(5)
    with pytest.raises(db.DrillError, match="closed"):
        f.result(5)


def test_an_answer_for_a_cancelled_caller_is_dropped_quietly(browser):
    page = BoxPage()
    f = browser.request(lambda s: (page, "/x", {}), 30)
    _wait_fired(page, 1)
    f.cancel()
    page.answer("/x")
    browser.submit(lambda s: None).result(5)              # a job, then a collect: must not raise on the worker
    ok = browser.request(lambda s: (page, "/y", {}), 30)
    _wait_fired(page, 1)
    page.answer("/y")
    assert ok.result(5) == (200, "ok")


def test_pager_steps_like_page_through():
    import json
    pg = db.Pager(10, can_recover=True)
    page = lambda items, more, total=None: json.dumps({"responsePayload": {"items": items, "hasMore": more,  # noqa: E731
                                                                          "totalResults": total}})
    assert pg.take(200, page([1, 2], True, 3)) == "more" and pg.off == 2
    assert pg.take(200, "<html>exit</html>") == "recover" and pg.off == 2
    assert pg.take(200, page([3], False)) == "done"
    assert pg.result() == ([1, 2, 3], 3)
    with pytest.raises(db.DrillError, match="HTTP 200 without data at offset 0"):
        db.Pager(10, can_recover=False).take(200, "<html/>")
    with pytest.raises(db.TooManyLines):
        db.Pager(2).take(200, page([1], True, 3))
