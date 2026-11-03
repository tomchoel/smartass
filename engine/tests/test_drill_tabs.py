"""drill_tabs: several drill pages at once on the browser's one worker thread (fake browser and tabs)."""
import concurrent.futures as cf
import json
import threading
from types import SimpleNamespace

import pytest

from smartview_mcp import drill_tabs
from smartview_mcp.drill_browser import CellTooWide


class _Req:
    def __init__(self, url):
        self.url, self.post_data, self.headers = url, "{}", {"x": "1"}


class _Resp:
    url = "https://epm/aif/ui/model/view?page=DRILL_THROUGH"

    def json(self):
        return {"v": 1}


class _Tab:
    def __init__(self, world):
        self.world, self.handlers, self.url, self.closed, self.epm, self.target = world, {}, "about:blank", False, None, None

    def on(self, ev, fn):
        self.handlers[ev] = fn

    def goto(self, url, wait_until=None, timeout=None):
        self.url = self.target = url
        self.world.open.add(self)
        self.world.peak = max(self.world.peak, len(self.world.open))

    def is_closed(self):
        return self.closed

    def close(self):
        self.closed = True
        self.world.open.discard(self)

    def wait_for_timeout(self, ms):
        for tab in list(self.world.open):
            how = self.world.how.get(tab.target.rsplit("/", 1)[-1], "ok")
            if how == "ok":
                tab.handlers["request"](_Req("https://epm/aif/ui/model/drilldown/getData?limit=500&offset=0"))
                tab.handlers["response"](_Resp())
            elif how == "signin":
                tab.url = "https://epm/ui/v1/signin"
            elif how == "wide":
                tab.epm = "EPMFDM-140373: more than 3 dimensions"

    def evaluate(self, js, args):
        n = self.world.lines[self.target.rsplit("/", 1)[-1]]
        return {"s": 200, "t": json.dumps({"responsePayload": {"items": [{"AMOUNTX": "-1"}] * n,
                                                               "totalResults": n, "hasMore": False}})}


class _Browser:
    base = "https://epm"

    def __init__(self, world):
        self.world, self.fallbacks = world, []

    def submit(self, fn):
        fut = cf.Future()

        def go():
            self.world.gate.wait(5)
            try:
                fut.set_result(fn(self.world.s))
            except BaseException as e:   # noqa: BLE001
                fut.set_exception(e)
        threading.Thread(target=go, daemon=True).start()
        return fut

    def _ensure(self, s):
        pass

    @staticmethod
    def _epm_error(tab):
        return tab.epm

    def _fetch(self, s, url, max_lines, progress):
        self.fallbacks.append(url)
        return {"items": [{"AMOUNTX": "-2"}], "total": 1, "view": {"v": 1}, "secs": 0}


def _world(how=None, lines=None):
    w = SimpleNamespace(open=set(), peak=0, how=how or {}, lines=lines or {}, gate=threading.Event())
    w.s = SimpleNamespace(ctx=SimpleNamespace(new_page=lambda: _Tab(w)), page=_Tab(w), deadline=0)
    return w


def test_drills_load_in_several_tabs_at_once():
    w = _world(lines={"a": 3, "b": 1, "c": 2})
    pump = drill_tabs.TabPump(lambda: b, tabs=2)
    b = _Browser(w)
    futs = [pump.fetch(u, 100) for u in ("a", "b", "c")]
    w.gate.set()                                                # the worker starts with all three queued
    assert [len(f.result(timeout=5)["items"]) for f in futs] == [3, 1, 2]
    assert w.peak == 2 and not w.open and b.fallbacks == []     # two tabs at a time, all closed after


def test_a_signed_out_tab_falls_back_and_a_too_wide_cell_fails():
    w = _world(how={"x": "signin", "w": "wide"}, lines={"a": 1})
    b = _Browser(w)
    pump = drill_tabs.TabPump(lambda: b, tabs=3)
    fa, fx, fw = (pump.fetch(u, 100) for u in ("a", "x", "w"))
    w.gate.set()
    assert len(fa.result(timeout=5)["items"]) == 1 and fx.result(timeout=5)["items"][0]["AMOUNTX"] == "-2"
    with pytest.raises(CellTooWide):
        fw.result(timeout=5)
    assert b.fallbacks == ["x"] and not w.open


def test_a_pump_job_that_never_ran_fails_its_drills():
    class Dead(_Browser):
        def submit(self, fn):
            fut = cf.Future()
            fut.cancel()                                        # e.g. kill() dropped the queued job
            return fut
    pump = drill_tabs.TabPump(lambda: Dead(_world()), tabs=2)
    with pytest.raises(drill_tabs.DrillError, match="restarted"):
        pump.fetch("a", 10).result(timeout=5)
    assert pump._job is None
