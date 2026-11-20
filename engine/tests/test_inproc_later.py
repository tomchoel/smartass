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
from smartass.xl import jobs
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
    monkeypatch.setattr(later, "_asked_at", {})
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


def test_only_an_open_workbook_is_asked_for_and_another_file_of_its_name_is_not_saved(excel):
    app, wb, addin = excel
    assert later.queue(app, "not open.xlsx", False) is None and addin.pending is None
    later.queue(app, "answer.xlsx", False)
    wb.FullName = r"C:\elsewhere\answer.xlsx"                    # closed, and another file of that name opened
    out = addin.release()
    assert out["closed"] and not out["ok"] and wb.saved == 0


def test_a_failed_outcome_put_back_is_read_by_the_next_job_and_a_newer_one_is_not_overwritten(excel):
    app, wb, addin = excel
    later.queue(app, "answer.xlsx", False)
    addin.save_error = "locked"
    addin.release()
    out = later.take(app)
    later.give_back(app, out)
    assert later.take(app)["token"] == out["token"]
    later.queue(app, "answer.xlsx", False)
    addin.save_error = None
    newer = addin.release()
    later.give_back(app, out)
    assert later.take(app)["token"] == newer["token"]


def test_a_save_whose_excel_quit_first_is_not_counted_as_to_come_for_ever(excel, monkeypatch):
    app, wb, addin = excel
    monkeypatch.setattr(later, "_asked_at", {})
    later.queue(app, "answer.xlsx", False)
    assert later.status()["waiting"] == 1
    monkeypatch.setattr(later, "FRESH_S", -1)
    assert later.status()["waiting"] == 0



# ------------------------------------------------------------------ review fixes: the job's first contact (clear)
def test_the_jobs_first_contact_runs_a_waiting_save_first(excel):
    app, wb, addin = excel
    later.queue(app, "answer.xlsx", True)
    later.clear(app)
    assert wb.saved == 1 and wb.AutoSaveOn is True and addin.log[:2] == ["clear", "save answer.xlsx"]
    assert later.take(app)["how"] == "before a job" and addin.release() is None


def test_the_first_contact_while_a_save_runs_is_asked_again_until_it_is_done(excel, monkeypatch):
    app, wb, addin = excel
    monkeypatch.setattr(runner, "BUSY_POLL_S", 0.001)
    addin.loaded, addin.busy = True, 3                      # loaded by another server process: called by name
    later.clear(app)
    assert addin.busy == 0 and addin.log == ["clear"]
    addin.busy = 10 ** 6
    with pytest.raises(later.HostError, match="busy"):
        later.clear(app, limit=0.01)


def test_the_first_contact_with_an_excel_without_the_add_in_is_nothing(monkeypatch, tmp_path):
    import pywintypes
    app = fakexl.Application()
    monkeypatch.setattr(config, "INPROC", True)
    monkeypatch.setattr(config, "INPROC_DIR", tmp_path / "inproc")
    monkeypatch.setattr(runner, "_EXCELS", {})
    monkeypatch.setattr(runner, "_identify", lambda app, check=True: (id(app), 0.0))
    asked = []

    def run(macro, *a):
        asked.append(macro)
        raise pywintypes.com_error(-2147352567, "Exception occurred.", (0, "Microsoft Excel", "Cannot run the macro",
                                                                        None, 0, -2146827284), None)
    app.Run = run
    later.clear(app)                                        # no add-in loaded by anyone: no save can be waiting
    assert len(asked) == 1 and asked[0].startswith("SmartassClear_")


def test_every_turn_taken_in_an_excel_starts_with_the_first_contact(excel, monkeypatch):
    app, wb, addin = excel
    seen = []
    monkeypatch.setattr(later, "clear", lambda a, *k, **kw: seen.append(a))
    try:
        jobs.hold_turn(424242, app=app)
        jobs.hold_turn(424242, app=app)                     # again in the same job: asked again (a few ms)
        jobs.hold_turn(424243)
    finally:
        jobs.release_turns()
    assert seen == [app, app]


TURNS_WITHOUT_AN_EXCEL = {("_spawn_excel", "none"),       # an Excel this job started itself: no add-in there yet
                          ("_wait_busy_holder", "None")}   # it does not answer COM yet; its next turn names it


def _turns_without_their_excel(source: str, rel: str) -> list[str]:
    """'<rel>:<line> in <function>' of each hold_turn call in source that does not name its Excel (app=, not None),
    but the two TURNS_WITHOUT_AN_EXCEL: an Excel the job started itself, and a busy holder before it answers COM."""
    import ast
    found = []

    def visit(node, where):
        for child in ast.iter_child_nodes(node):
            inner = child.name if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)) else where
            if isinstance(child, ast.Call):
                f = child.func
                if (f.attr if isinstance(f, ast.Attribute) else getattr(f, "id", None)) == "hold_turn":
                    app = next((k.value for k in child.keywords if k.arg == "app"), None)
                    kind = ("none" if app is None
                            else "None" if isinstance(app, ast.Constant) and app.value is None else "app")
                    if kind != "app" and (where, kind) not in TURNS_WITHOUT_AN_EXCEL:
                        found.append(f"{rel}:{child.lineno} in {where}")
            visit(child, inner)
    visit(ast.parse(source), "")
    return found


def test_every_turn_the_target_code_takes_names_its_excel():
    """The structural guarantee: each hold_turn in the code that resolves a target or finds a workbook passes the
    Excel's app, so the first contact runs before any plain COM writes there (an Excel this job starts itself has no
    add-in yet, and a busy holder that does not answer COM yet: the exceptions)."""
    from pathlib import Path
    import smartass
    root = Path(smartass.__file__).parent
    texts = {rel: (root / rel).read_text("utf-8") for rel in ("xl/target.py", "xl/workbench/books.py")}
    assert all("hold_turn(" in t for t in texts.values())
    assert [m for rel, text in texts.items() for m in _turns_without_their_excel(text, rel)] == []


def test_the_turn_check_finds_a_turn_without_its_excel():
    """Review 2 (Task 2): the check let any hold_turn(..., wait=False) without an app pass."""
    rogue = ("def find(pid, app):\n"
             "    jobs.hold_turn(pid, wait=False)\n"
             "    jobs.hold_turn(pid, app=None)\n"
             "    hold_turn(pid, what='x', app=app)\n"
             "def _spawn_excel(tpl, known):\n"
             "    jobs.hold_turn(pid, wait=False)\n")
    assert _turns_without_their_excel(rogue, "x.py") == ["x.py:2 in find", "x.py:3 in find"]


def test_an_outcome_the_add_in_could_not_write_comes_back_with_the_first_contact(excel):
    app, wb, addin = excel
    later.queue(app, "answer.xlsx", False)
    addin.write_fails, addin.save_error = True, "locked"
    addin.release()
    assert later.take(app) is None                          # nothing on disk
    later.clear(app)
    out = later.take(app)
    assert out["ok"] is False and out["error"] == "locked" and out["unwritten"] == "the disk is full"


@pytest.mark.parametrize("text", ["{not json", '{"token": "a", "ok": false, "started": 1}', ""])
def test_a_malformed_outcome_is_an_unknown_one_and_read_once(excel, text):
    app, wb, addin = excel
    folder = config.INPROC_DIR / "saved"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / f"{id(app)}.json").write_text(text, encoding="utf-8")
    out = later.take(app)
    assert out["unknown"] is True and out["ok"] is False and "could not be read" in out["error"]
    assert later.take(app) is None


def test_old_outcomes_of_excels_that_are_gone_are_pruned(excel, monkeypatch):
    import os
    import time
    app, wb, addin = excel
    folder = config.INPROC_DIR / "saved"
    folder.mkdir(parents=True, exist_ok=True)
    gone, old, alive = folder / "999999991.json", folder / "4.json", folder / f"{os.getpid()}.json"
    for p in (gone, old, alive):
        p.write_text("{}", encoding="utf-8")
    week = time.time() - 8 * 24 * 3600
    os.utime(old, (week, week))
    monkeypatch.setattr(later, "_alive", lambda pid: pid in (4, os.getpid()))
    later.queue(app, "answer.xlsx", False)
    assert not gone.exists() and not old.exists() and alive.exists()
