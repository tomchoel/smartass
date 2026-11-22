import pytest

from smartass.core import timing
from smartass.xl import login
from smartass.xl.host import HostError


@pytest.fixture(autouse=True)
def _no_timer_left():
    yield
    timing._local.timer = timing._local.pending_notify = None


class Clock:
    def __init__(self):
        self.t = 0.0

    def now(self):
        return self.t

    def sleep(self, s):
        self.t += s


def _fake(monkeypatch, open_for: float, hung: bool = False):
    c = Clock()
    monkeypatch.setattr(login, "login_open", lambda pid: [7] if c.t < open_for else [])
    monkeypatch.setattr(login.dialogs, "scan", lambda pid: [])
    monkeypatch.setattr(login.dialogs, "is_hung", lambda hwnd: hung)
    return c


def test_a_slow_sso_click_tells_the_user_once(monkeypatch):
    c = _fake(monkeypatch, open_for=12)
    notes = []
    timing.Timer.note_notify(notes.append)
    timing.Timer.start()
    login.wait_login_cleared(1, clock=c.now, sleep=c.sleep)
    assert len(notes) == 1 and "Login" in notes[0]


def test_a_hung_login_in_the_users_excel_fails_fast_and_closes_nothing(monkeypatch):
    c = _fake(monkeypatch, open_for=999, hung=True)
    closed = []
    monkeypatch.setattr(login.dialogs, "close", closed.append)
    with pytest.raises(HostError) as e:
        login.wait_login_cleared(1, clock=c.now, sleep=c.sleep)
    assert "not responding" in str(e.value) and "nothing was closed" in str(e.value)
    assert c.t <= login.LOGIN_HUNG_S + 1 and closed == []
