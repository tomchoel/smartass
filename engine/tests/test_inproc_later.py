"""Phase 3c Task 2: xl/inproc/later.py, the save after the reply - asked of the add-in in one Run, run by it once Excel
has painted (here: when the test lets the fake add-in's timer fire), its outcome read once by the next job in that
Excel, waited for by the scripts, shown by fin_status. On the fake Excel with the fake add-in (fakeaddin.py)."""
from __future__ import annotations

import json
import logging

import pytest
from fakeaddin import install

import fakexl
from smartass.core import config
from smartass.core.timing import BATCH_PREFIX, Timer, stage
from smartass.xl.inproc import later, runner
from smartass.xl.inproc.batch import Batch


@pytest.fixture
def excel(monkeypatch, tmp_path):
    app = fakexl.Application()
    wb = fakexl.Workbook(["Sources"], app=app, styles_ver=False)
    wb.Name = "answer.xlsx"
    wb.AutoSaveOn = False
    addin = install(monkeypatch, app, tmp_path)
    monkeypatch.setattr(later, "_queued", {})
    monkeypatch.setattr(later, "_done", {})
    return app, wb, addin


def test_the_save_is_asked_for_in_one_run_and_done_by_excel_after_the_reply(excel):
    app, wb, addin = excel
    Timer.start()
    with stage("save"):
        token = later.queue(app, "answer.xlsx", True)
    assert token and addin.asked == [("answer.xlsx", True, token, config.SAVE_DELAY_S * 1000)]
    assert wb.saved == 0 and wb.AutoSaveOn is False                 # nothing yet: Excel paints first
    assert Timer.current().summary()[BATCH_PREFIX + "save"] == 1
    out = addin.release()
    assert out["ok"] and wb.saved == 1 and wb.AutoSaveOn is True


def test_the_next_job_reads_the_outcome_once(excel):
    app, wb, addin = excel
    token = later.queue(app, "answer.xlsx", False)
    addin.release()
    out = later.take(app)
    assert out["token"] == token and out["ok"] and out["seconds"] == 0.25 and out["how"] == "after the reply"
    assert later.take(app) is None                                  # read once
    assert json.loads((config.INPROC_DIR / "saved" / "last.json").read_text())["token"] == token


def test_a_batch_that_comes_before_the_save_ran_runs_it_first_and_only_once(excel):
    app, wb, addin = excel
    later.queue(app, "answer.xlsx", True)
    b = Batch(quiet=False)
    b.ret(b.workbook("answer.xlsx").Name, "name")
    runner.run(app, b)
    assert addin.log == ["save answer.xlsx", "autosave answer.xlsx True", "batch"]
    assert later.take(app)["how"] == "before a batch"
    assert addin.release() is None and wb.saved == 1                # the timer finds nothing left to save


def test_without_the_add_in_or_with_sv_save_sync_nothing_is_asked(excel, monkeypatch):
    app, wb, addin = excel
    monkeypatch.setattr(config, "SAVE", "sync")
    assert later.queue(app, "answer.xlsx", True) is None
    monkeypatch.setattr(config, "SAVE", "later")
    monkeypatch.setattr(config, "INPROC", False)
    assert later.queue(app, "answer.xlsx", True) is None and later.take(app) is None
    assert addin.asked == [] and not later.available(app)


def test_a_refused_request_is_none_with_a_warning(excel, caplog):
    app, wb, addin = excel
    addin.answer = "not run: refused: not a token"
    with caplog.at_level(logging.WARNING):
        assert later.queue(app, "answer.xlsx", True) is None
    assert "refused" in caplog.text and later._queued == {}


def test_an_outcome_of_an_earlier_excel_with_the_same_pid_is_not_taken(excel, monkeypatch):
    app, wb, addin = excel
    later.queue(app, "answer.xlsx", False)
    addin.release()
    monkeypatch.setattr(runner, "_identify", lambda app, check=True: (id(app), 4e9))   # started after the outcome
    assert later.take(app) is None


def test_the_scripts_wait_for_every_save_asked_for_here(excel):
    app, wb, addin = excel
    first = later.queue(app, "answer.xlsx", False)
    second = later.queue(app, "answer.xlsx", False)          # runs the first before it is queued
    assert later.take(app)["token"] == first                 # a job read the first one
    addin.release()
    got = later.wait(1)
    assert [o["token"] for o in got] == [first, second] and all(o["ok"] for o in got)
    assert later.wait(1) == []                               # each once


def test_a_save_that_never_reports_is_a_failure_after_the_wait(excel):
    app, wb, addin = excel
    token = later.queue(app, "answer.xlsx", False)
    (out,) = later.wait(0.05)
    assert out["token"] == token and not out["ok"] and "no outcome" in out["error"]


def test_fin_status_shows_the_last_outcome_without_figures(excel):
    app, wb, addin = excel
    later.queue(app, "answer.xlsx", True)
    assert later.status()["waiting"] == 1 and later.status()["last"] is None
    addin.save_error = "the disk is full"
    addin.release()
    st = later.status()
    assert st["waiting"] == 0
    assert set(st["last"]) == {"workbook", "ok", "error", "closed", "seconds", "at", "how"}
    assert st["last"]["workbook"] == "answer.xlsx" and st["last"]["ok"] is False and "disk" in st["last"]["error"]
    assert wb.AutoSaveOn is True                             # back on after the failed save too
