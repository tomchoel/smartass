"""The hidden Excel's Smart View message boxes are closed - its own pid only, that caption only."""
import threading

from smartass.xl import dialogs, host
from smartass.xl.dialogs import Dialog

SV = Dialog(1, "#32770", "Oracle Smart View for Office", "Empty response from provider", "message")
OTHER = Dialog(2, "#32770", "Microsoft Excel", "Something", "message")
LOGIN = Dialog(3, "#32770", "Login", "", "login")


def test_closes_only_smart_view_boxes_of_the_pid():
    scanned, pressed, closed = [], [], []

    def scan(pid):
        scanned.append(pid)
        return [SV, OTHER, LOGIN]
    texts = dialogs.close_sv_boxes(42, scan=scan, press=lambda h: pressed.append(h) or True,
                                   close=lambda hs: closed.extend(hs))
    assert scanned == [42] and pressed == [1] and closed == []
    assert texts == ["Empty response from provider"]


def test_falls_back_to_wm_close_when_not_ok_only():
    closed = []
    dialogs.close_sv_boxes(42, scan=lambda pid: [SV], press=lambda h: False, close=lambda hs: closed.extend(hs))
    assert closed == [1]


def test_scan_error_never_raises():
    def boom(pid):
        raise RuntimeError("x")
    assert dialogs.close_sv_boxes(42, scan=boom) == []


def test_host_watcher_acts_on_host_pid_only(monkeypatch):
    h = host.ExcelHost.__new__(host.ExcelHost)
    h.pid, h.ctime = 77, 5.0
    h._box_stop = threading.Event()
    seen = []
    monkeypatch.setattr(host, "_is_ours", lambda pid, ctime: True)
    monkeypatch.setattr(host.dialogs, "close_sv_boxes", lambda pid: seen.append(pid) or ["t"])
    h._watch_boxes_once()
    h.pid = None
    h._watch_boxes_once()
    assert seen == [77]


def test_host_watcher_skips_a_pid_that_is_no_longer_the_hosts_excel(monkeypatch):
    h = host.ExcelHost.__new__(host.ExcelHost)
    h.pid, h.ctime = 77, 5.0
    seen = []
    monkeypatch.setattr(host.dialogs, "close_sv_boxes", lambda pid: seen.append(pid) or [])
    monkeypatch.setattr(host, "_is_ours", lambda pid, ctime: False)
    h._watch_boxes_once()
    assert seen == []
