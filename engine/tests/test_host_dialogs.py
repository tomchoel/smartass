"""ExcelHost.call: Smart View sign-in / WebView2 dialogs in the hidden Excel lead to one restart and retry.
Fake Excel (pids only), fake dialog detector and fake kill - no real Excel or Win32 windows are touched."""
import threading
import time

import pytest
import pywintypes

from smartass.core import config
from smartview_mcp import host
from smartview_mcp.dialogs import Dialog
from smartview_mcp.host import ComThread, ExcelHost, HostError, HostTimeout, SignInBlocked

LOGIN_IE = Dialog(11, "#32770", "Login", "", "login", True)
MESSAGE = Dialog(12, "#32770", "Oracle Smart View for Office", "Empty response from provider", "message")


class _Registry:
    def remove(self, pid):
        pass


class FakeExcel:
    """Each start gets a new pid. A 'hanging' call blocks until that pid is killed, then fails like a dead RPC."""

    def __init__(self, monkeypatch, dialogs_by_start):
        self.dialogs_by_start = dialogs_by_start
        self.starts, self.killed = [], []
        self.released: dict[int, threading.Event] = {}
        h = ExcelHost.__new__(ExcelHost)
        h.thread = ComThread()
        h.thread.start()
        h.registry = _Registry()
        h.app = h.wb = h.pid = h.ctime = None
        h.connected, h.restarts, h.last_error = {}, 0, None
        h._known_before, h._start_wall = set(), time.time()
        h._resetting, h._reset_lock, h.reset_gen = threading.Event(), threading.Lock(), 0
        h._ensure_started = self._ensure
        h.detect = self._detect
        self.h = h
        self.start_s = 0.0
        monkeypatch.setattr(host, "kill_ours", self._kill)
        monkeypatch.setattr(host, "windows_of", lambda pid: [])
        monkeypatch.setattr(host.dialogs, "webview2_holders", lambda exclude_parent: [])
        monkeypatch.setattr(config, "DIALOG_CHECK_AFTER_S", 0.05)
        monkeypatch.setattr(config, "DIALOG_POLL_S", 0.05)

    def _ensure(self):
        if self.h.pid is None or self.h.pid in self.killed:     # a killed Excel is gone: start a new one
            self.starts.append(1)
            pid = 1000 + len(self.starts)
            self.released[pid] = threading.Event()
            time.sleep(self.start_s)
            self.h.pid, self.h.ctime = pid, 1.0

    def _detect(self, pid):
        i = pid - 1001
        return list(self.dialogs_by_start[i]) if i < len(self.dialogs_by_start) else []

    def _kill(self, pid, ctime):
        self.killed.append(pid)
        self.released[pid].set()
        return True

    def hang(self, h):
        if not self.released[h.pid].wait(5):
            raise AssertionError("never killed")
        raise pywintypes.com_error(-2147023170, "The remote procedure call failed.", None, None)


@pytest.mark.anyio
async def test_login_dialog_hang_restarts_excel_and_retries_once(monkeypatch):
    xl = FakeExcel(monkeypatch, [[LOGIN_IE], []])

    def fn(h):
        return "ok" if h.pid == 1002 else xl.hang(h)

    assert await xl.h.call(fn, timeout=5, what="sv_answer") == "ok"
    assert xl.killed == [1001] and len(xl.starts) == 2


@pytest.mark.anyio
async def test_login_dialog_twice_raises_the_cause_and_next_step(monkeypatch):
    xl = FakeExcel(monkeypatch, [[LOGIN_IE], [LOGIN_IE]])
    with pytest.raises(HostError) as e:
        await xl.h.call(xl.hang, timeout=5, what="sv_answer")
    msg = str(e.value)
    assert "sign-in window" in msg and "Internet Explorer" in msg
    assert "restarted Excel once; it happened again" in msg
    assert "smartass doctor --bundle" in msg
    assert not isinstance(e.value, HostTimeout)
    assert xl.killed == [1001, 1002]


@pytest.mark.anyio
async def test_message_dialog_is_named_in_the_timeout_error(monkeypatch):
    xl = FakeExcel(monkeypatch, [[MESSAGE]])
    with pytest.raises(HostTimeout) as e:
        await xl.h.call(xl.hang, timeout=0.5, what="sv_answer")
    assert "a dialog is open in the hidden Excel: 'Oracle Smart View for Office: Empty response from provider'" \
        in str(e.value)
    assert xl.killed == [1001] and len(xl.starts) == 1


@pytest.mark.anyio
async def test_young_webview2_login_is_left_to_the_sso_watcher(monkeypatch):
    xl = FakeExcel(monkeypatch, [[Dialog(11, "#32770", "Login", "", "login", False)]])

    def fn(h):
        time.sleep(0.4)                       # the SSO watcher clicks 'Telenor SSO' meanwhile
        return "ok"

    assert await xl.h.call(fn, timeout=5, what="sv_answer") == "ok"
    assert xl.killed == []


@pytest.mark.anyio
async def test_sign_in_blocked_raised_by_the_job_restarts_and_retries(monkeypatch):
    xl = FakeExcel(monkeypatch, [[LOGIN_IE], []])

    def fn(h):
        if h.pid == 1001:
            h.check_signin()                  # what ops does after a transient rc=752 / rc=57704
            raise AssertionError("check_signin should have raised")
        return "ok"

    assert await xl.h.call(fn, timeout=5, what="sv_answer") == "ok"
    assert xl.killed == [1001]


@pytest.mark.anyio
async def test_a_call_queued_behind_a_sign_in_reset_does_not_orphan_an_excel(monkeypatch):
    import asyncio
    xl = FakeExcel(monkeypatch, [[LOGIN_IE], [], []])
    xl.start_s = 0.5                          # a start outlasts the reset's wait for the COM thread
    monkeypatch.setattr(config, "RESET_WAIT_S", 0.2)

    def a(h):
        return "a" if h.pid != 1001 else xl.hang(h)

    async def b():
        await asyncio.sleep(0.02)             # queued while A is blocked by the Login window
        return await xl.h.call(lambda h: "b", timeout=5, what="sv_members")

    assert await asyncio.gather(xl.h.call(a, timeout=5, what="sv_answer"), b()) == ["a", "b"]
    started = list(range(1001, 1001 + len(xl.starts)))
    assert all(p in xl.killed or p == xl.h.pid for p in started), (started, xl.killed, xl.h.pid)
    assert xl.killed == [1001] and xl.h.reset_gen == 1


@pytest.mark.anyio
async def test_a_job_starting_as_the_reset_begins_is_not_waited_for(monkeypatch):
    """check_signin path: the COM thread is idle, so a queued job B starts the moment the reset raises its flag.
    The reset must not wait on B (B waits for the reset) nor swap the thread under it."""
    xl = FakeExcel(monkeypatch, [[], []])
    h = xl.h
    assert await h.call(lambda h: h.pid, timeout=5) == 1001
    monkeypatch.setattr(config, "RESET_WAIT_S", 1.0)
    thread0, b = h.thread, {}

    class Hook(threading.Event):
        def set(self):
            super().set()
            n = h.thread.job_no

            def job_b():
                h._wait_reset()
                h._ensure_started()
                return h.pid
            b["fut"] = h.thread.submit(job_b)
            while h.thread.job_no == n:          # B is running (and waiting for the reset)
                time.sleep(0.01)
    h._resetting = Hook()
    t0 = time.monotonic()
    h._hard_reset("sign-in stuck", 1001)
    assert time.monotonic() - t0 < 0.8 and h.thread is thread0
    assert b["fut"].result(5) == 1002 == h.pid and xl.killed == [1001]


@pytest.mark.anyio
async def test_a_job_waiting_for_the_reset_is_not_counted_as_blocked(monkeypatch):
    xl = FakeExcel(monkeypatch, [[], []])
    h = xl.h
    await h.call(lambda h: h.pid, timeout=5)
    monkeypatch.setattr(config, "RESET_WAIT_S", 1.0)
    thread0 = h.thread
    h._resetting.set()
    fut = h.thread.submit(h._wait_reset)
    while not h.thread.waiting_reset:
        time.sleep(0.01)
    t0 = time.monotonic()
    h._do_hard_reset("x", h.thread, h.thread.job_no)   # even a snapshot taken after that job started
    assert time.monotonic() - t0 < 0.8 and h.thread is thread0
    h._resetting.clear()
    fut.result(5)


def test_check_signin_ignores_plain_messages():
    h = ExcelHost.__new__(ExcelHost)
    h.pid = 5
    h.detect = lambda pid: [MESSAGE]
    h.check_signin()
    h.detect = lambda pid: [LOGIN_IE]
    with pytest.raises(SignInBlocked):
        h.check_signin()


@pytest.mark.anyio
async def test_hung_login_restarts_excel_and_names_the_cause(monkeypatch):
    hung = Dialog(13, "#32770", "Login", "", "login", False, hung=True)
    xl = FakeExcel(monkeypatch, [[hung], [hung]])
    with pytest.raises(SignInBlocked) as e:
        await xl.h.call(xl.hang, timeout=5, what="sv_answer")
    assert "stopped responding" in str(e.value)
    assert xl.killed == [1001, 1002]


@pytest.mark.anyio
async def test_provider_empty_restarts_the_hidden_excel_once_and_says_so(monkeypatch):
    xl = FakeExcel(monkeypatch, [[], []])
    notes = []

    def fn(h):
        if h.pid == 1001:
            raise host.ProviderEmpty("rc=752 Empty response from provider", pid=1001, notify=notes.append)
        return "ok"

    assert await xl.h.call(fn, timeout=5, what="sv_answer") == "ok"
    assert xl.killed == [1001]
    assert any("yours is untouched" in n for n in notes)


@pytest.mark.anyio
async def test_provider_empty_twice_gives_the_next_step(monkeypatch):
    xl = FakeExcel(monkeypatch, [[], []])

    def fn(h):
        raise host.ProviderEmpty("rc=752 Empty response from provider", pid=h.pid, notify=None)

    with pytest.raises(HostError) as e:
        await xl.h.call(fn, timeout=5, what="sv_answer")
    assert "Empty response from provider" in str(e.value) and "smartass doctor --bundle" in str(e.value)
    assert xl.killed == [1001, 1002]


@pytest.mark.anyio
async def test_provider_empty_in_a_user_excel_job_restarts_only_ours_and_does_not_retry(monkeypatch):
    xl = FakeExcel(monkeypatch, [[]])
    xl.h._ensure_started()                    # the hidden Excel exists (pid 1001)
    runs = []

    def fn(h):
        runs.append(1)
        raise host.ProviderEmpty("rc=752 Empty response from provider", pid=1001, notify=None)

    with pytest.raises(HostError) as e:
        await xl.h.call(fn, timeout=5, what="sv_answer", user_excel=True)
    assert len(runs) == 1 and xl.killed == [1001]
    assert "ask again" in str(e.value)
