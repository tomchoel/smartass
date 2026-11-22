"""Unit tests for the analysis logic that needs no Excel."""
from types import SimpleNamespace

from smartass.epm.grid import Axis, Layout
from smartass.epm.grid import verify_labels as _verify_labels
from smartass.xl.instances import is_archive, local_path


def test_verify_labels_catches_shifted_rows():
    lay = Layout([], [Axis("All Accounts", ["ac.C30", "ac.C40S"])], [Axis("All Time Periods", ["Sep 26"])])
    ok = SimpleNamespace(row_labels=[["C30 REVENUES (ac)"], ["C40S Total Cost (ac)"]], col_labels=[["Sep 26"]])
    shifted = SimpleNamespace(row_labels=[["C40S Total Cost (ac)"], ["C48 Contribution (ac)"]], col_labels=[["Sep 26"]])
    assert _verify_labels(lay, ok) == []
    assert len(_verify_labels(lay, shifted)) == 2
    from smartass.epm.grid import verify_labels
    assert verify_labels is _verify_labels           # one implementation, shared by the COM and HTTP routes


def test_archive_guard_local_url_and_unsaved(monkeypatch, tmp_path):
    from smartass.core import config
    root = tmp_path / "Analyse (Telenor Norge AS) - Produktregnskapet"
    monkeypatch.setattr(config, "ARCHIVE_ROOT", root)
    assert is_archive(str(root / "2025" / "x.xlsx"))
    assert is_archive("https://telenorgroup.sharepoint.com/sites/AnalyseTelenorNorgeAS/Shared Documents/2025/x.xlsx")
    assert not is_archive(str(tmp_path / "analyser" / "a.xlsx"))
    assert not is_archive("Book1")


def test_local_path_maps_personal_onedrive(monkeypatch):
    from smartass.xl import instances
    monkeypatch.setattr(instances, "_onedrive_mounts",
                        lambda: [("https://contoso-my.sharepoint.com/personal/a_b/Documents/", r"C:\OD")])
    url = "https://contoso-my.sharepoint.com/personal/a_b/Documents/smartass/analyser/a%20b.xlsx"
    assert local_path(url) == r"C:\OD\smartass\analyser\a b.xlsx"


def test_slow_stage_sends_progress_notes_and_stops_after(monkeypatch):
    import time as _t

    from smartass.core import timing
    monkeypatch.setattr(timing, "PROGRESS_AFTER_S", 0.05)
    notes = []
    timing.Timer.note_notify(notes.append)
    timing.Timer.start()
    with timing.stage("open target"):
        with timing.stage("slow"):
            _t.sleep(0.13)
    n = len(notes)
    _t.sleep(0.12)
    assert n >= 1 and "open target.slow" in notes[0] and notes[0].startswith("still working: ")
    assert len(notes) == n                       # no notes once the stages are done
    timing.Timer.start()                         # the next timer on this thread has no notify of its own
    assert timing.Timer.current().notify is None


def test_progress_job_clears_its_notify_after_the_job():
    import asyncio

    from smartass.core import serverkit as server
    from smartass.core import timing

    async def run():
        server._ctx.set(object())                 # any ctx: notify is only called by a ticker
        job = server._with_progress(lambda h: timing.Timer.start())
        t = job(None)
        return t
    t = asyncio.run(run())
    assert t.notify is None and getattr(timing._local, "pending_notify", None) is None


def test_com_retries_a_busy_excel_then_says_so_in_words(monkeypatch):
    import pytest
    import pywintypes

    from smartass.xl import hostproc
    from smartass.xl import jobs
    from smartass.xl.host import HostError
    monkeypatch.setattr(hostproc.time, "sleep", lambda s: None)
    rejected = pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)
    recalc = pywintypes.com_error(-2147352567, "Exception occurred.", (0, None, None, None, 0, -2146777998), None)
    answers = [rejected, recalc, "ok"]

    def call():
        a = answers.pop(0)
        if isinstance(a, Exception):
            raise a
        return a
    assert jobs.com(call, check=False) == "ok"
    with pytest.raises(HostError, match="Excel is busy"):
        jobs.com(lambda: (_ for _ in ()).throw(rejected), limit=0, check=False)
    other = pywintypes.com_error(-2147352567, "Exception occurred.", (0, None, "Bad", None, 0, -2146827284), None)
    with pytest.raises(pywintypes.com_error):
        jobs.com(lambda: (_ for _ in ()).throw(other), check=False)
