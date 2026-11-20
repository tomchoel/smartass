# engine/tests/v2/test_file_job.py
"""Phase 3 Tasks 5 and 9, phase 3b Task 2b: pipeline.excel_job, a new answer through a built file (render/xlsx.build,
then xl/movein) brought in by batches (xl/inproc, here through the COM executor on the fake Excel): the steps in
order, the rollback and the temporary file on a failure, the batches of a whole job, the reply, one timestamp per
draw, where the job lands and when Excel may paint, and server.py drawing every new question with it. All on fakes;
every figure is invented."""
from __future__ import annotations

import asyncio
import logging
from types import SimpleNamespace

import pytest
from test_look import fakexl_sheet_with, lines_ranking, pl_table
from test_v2_pipeline import LINE_CELL, FakeProvider, lines_only_spec
from v2util import facts, resolver, spec_dict

import fakexl
from smartass import followup, pipeline, server
from smartass.core import config
from smartass.core.timing import BATCH_PREFIX, COM_PREFIX
from smartass.model.errors import GuardError
from smartass.model.lines import LineSet
from smartass.render import look, xlsx
from smartass.xl import movein, sheets
from smartass.xl.host import HostError
from smartass.xl.inproc import runner
from smartass.xl.inproc.batch import Batch
from smartass.xl.inproc.runner import BatchError
from smartass.xl.target import Target

H = None                                        # the host: every seam that would use it is replaced
STAGES = ["open target", "draw", "draw.look", "draw.build", "draw.move in", "draw.finish", "draw.close out", "save"]
BATCHES = {"draw.look": 1, "draw.close out": 1, "save": 1}      # + the bring-in batch, booked in 'draw' itself
BASE = xlsx.Base("Telenor", 11, row_px=19, col_px=72)
SUPPLIERS = [{"AP_CUSTOMER_NAME": "Bouvet", "AMOUNT": -30.0}, {"AP_CUSTOMER_NAME": "Accenture", "AMOUNT": -70.0}]


def _pre(spec=None, lines=None):
    return pipeline.prepare(spec or spec_dict(), resolver(), FakeProvider(facts()), lines=lines)


@pytest.fixture
def tmp_home(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "HOME", tmp_path)
    return tmp_path


def _temp_files(home) -> list:
    return list((home / "tmp").glob("*.xlsx")) if (home / "tmp").exists() else []


# ------------------------------------------------------------------ the steps, every seam replaced
def _look(book=None, **kw) -> pipeline.Look:
    return pipeline.Look(book or movein.Book([movein.Sheet("Sources")]), kw.pop("screen", movein.Screen()), BASE, [],
                         None, sheets.SpecRows(), (",", "\xa0"))


class Wired:
    """Every batch of excel_job replaced by a recorder (the build is the real one: the temporary file is real)."""

    def __init__(self, monkeypatch):
        self.calls: list[str] = []
        self.failures: dict[str, Exception] = {}
        self.rolled_back: list[str] | None = None
        wb = fakexl.Workbook(["Sources"], styles_ver=False)
        self.t = Target(app=wb.Application, wb=wb, pid=0, hwnd=0, created=True)
        self.built: list[xlsx.BuiltFile] = []
        real_build = xlsx.build

        def build(path, *a, **k):
            self.path, self.base = path, k.get("base")
            self.built.append(real_build(path, *a, **k))
            return self.built[-1]

        def bring_in(job, built, plan, pre, header, book):
            job.moved, job.built = True, built
            job.t.sheets_added.extend(built.sheets)
            job.inserted = True
            return (), ({}, {}), 1, None

        def rollback(t):
            self.rolled_back = list(t.sheets_added)
        put = monkeypatch.setattr
        put(pipeline, "_target", self.rec("target", ret=self.t))
        put(pipeline, "_look", self.rec("look", ret=_look()))
        put(xlsx, "build", self.rec("build", build))
        put(pipeline, "_bring_in", self.rec("bring_in", bring_in))
        put(pipeline, "compare", lambda *a: [])
        put(pipeline, "_close_out", self.rec("close_out", ret=(None, None, [])))
        put(pipeline, "_save", self.rec("save"))
        put(pipeline, "_rollback", self.rec("rollback", rollback))
        put(pipeline, "remember", self.rec("remember"))

    def rec(self, name, fn=None, ret=None):
        def f(*a, **k):
            self.calls.append(name)
            if name in self.failures:
                raise self.failures[name]
            return fn(*a, **k) if fn is not None else ret
        return f

    def fail(self, name: str, error: Exception) -> None:
        self.failures[name] = error


@pytest.fixture
def wired(monkeypatch, tmp_home):
    return Wired(monkeypatch)


def test_the_file_job_runs_its_steps_in_order(wired, tmp_home):
    out = pipeline.excel_job(H, _pre())
    assert wired.calls == ["target", "look", "build", "bring_in", "close_out", "save", "remember"]
    assert out["sheets"] == ["Q1 Answer", "Q1 Data"] and out["checks"]["sheet_equals_model"] == "OK"
    assert wired.path.parent == tmp_home / "tmp" and wired.path.suffix == ".xlsx" and not wired.path.exists()


def test_the_file_is_built_in_the_destinations_base(wired):
    pipeline.excel_job(H, _pre())
    assert wired.base == BASE


def test_the_temporary_file_has_a_new_name_every_time(wired):
    pipeline.excel_job(H, _pre())
    first = wired.path
    pipeline.excel_job(H, _pre())
    assert wired.path != first and len(wired.path.stem) == 32                  # a uuid4 hex


def test_a_failure_after_the_move_in_removes_the_new_sheets_and_the_temp_file(wired, tmp_home):   # Review Focus 3
    wired.fail("close_out", HostError("Excel is busy"))
    with pytest.raises(HostError, match="busy"):
        pipeline.excel_job(H, _pre())
    assert wired.rolled_back == ["Q1 Answer", "Q1 Data"] and not _temp_files(tmp_home)
    assert wired.calls.count("rollback") == 1
    assert "save" not in wired.calls and "remember" not in wired.calls


def test_a_build_that_fails_leaves_no_partial_file_and_inserts_nothing(wired, tmp_home, monkeypatch):
    def half(path, *a, **k):
        path.write_bytes(b"PK half a file")
        raise ValueError("the build broke")
    monkeypatch.setattr(xlsx, "build", wired.rec("build", half))
    with pytest.raises(ValueError, match="broke"):
        pipeline.excel_job(H, _pre())
    assert "move_in" not in wired.calls and wired.calls[-1] == "rollback" and not _temp_files(tmp_home)


def test_more_lines_than_a_sheet_has_rows_is_a_refusal_with_its_fix(wired, tmp_home, monkeypatch):  # Focus 5
    def too_many(path, *a, **k):
        raise xlsx.TooManyRows("sheet 'Q1 Lines b0' would need 1 100 000 rows; an Excel sheet has 1 048 576")
    monkeypatch.setattr(xlsx, "build", wired.rec("build", too_many))
    with pytest.raises(GuardError, match="1 100 000 rows") as e:
        pipeline.excel_job(H, _pre())
    assert "narrow" in e.value.fix and wired.calls[-1] == "rollback" and not _temp_files(tmp_home)


def test_the_spec_is_stored_without_a_drawn_range(fake_target, calculated, monkeypatch):
    seen, real = {}, sheets.save_spec

    def save_spec(t, q, spec, ids=None, owned=None, create=False, b=None, rows=None):
        seen.update(q=q, spec=spec, ids=ids, owned=owned)
        return real(t, q, spec, ids, owned, create, b, rows)
    monkeypatch.setattr(sheets, "save_spec", save_spec)
    pre = _pre()
    _job(pre, calculated)
    assert seen == {"q": 1, "spec": {"v2": pre.raw}, "ids": {}, "owned": None}


def test_the_chart_lint_and_the_drawn_charts_reach_the_reply_once(wired, monkeypatch):
    real = pipeline.layout

    def layout(*a, **k):
        plan = real(*a, **k)
        plan.chart_lint.append("chart 'P&L': a lint line")
        return plan
    monkeypatch.setattr(pipeline, "layout", layout)
    out = pipeline.excel_job(H, _pre())
    assert out["warnings"].count("chart 'P&L': a lint line") == 1
    assert out["charts"] == wired.built[0].charts and out["charts"]


def test_the_reply_has_the_keys_of_excel_jobs_reply(wired):
    pre = _pre()
    out = pipeline.excel_job(H, pre)
    assert set(out) == set(pipeline.reply(pre, 1, "x", [], [], [])) and out["q"] == 1


def test_an_arrange_warning_goes_to_the_reply_without_a_rollback(fake_target, calculated, monkeypatch):
    """Arranging the sheets is cosmetic: a move Excel refuses is a warning, the answer stays."""
    fake_target.wb.order.insert(0, fakexl.Worksheet(fake_target.wb, "Q1 Notes"))   # Q1 Answer must move before it
    for ws in fake_target.wb.order:
        ws.wb = fake_target.wb

    def refuse(ws, *args, **kw):
        raise fakexl.ComError("the move was refused")
    monkeypatch.setattr(fakexl.Worksheet, "Move", refuse)
    out, _ = _job(_pre(), calculated)
    assert any(w.startswith("could not arrange the sheets (") and "refused" in w for w in out["warnings"])
    assert "Q1 Answer" in [ws.Name for ws in fake_target.wb.order] and fake_target.wb.saved == 1


def test_the_answer_has_live_key_points_and_the_agents_free_points(wired, monkeypatch):
    """Moved from the COM job's tests: the plan the file is built from."""
    from smartass.render.layout import TEXT_COL
    plans, real = [], xlsx.build

    def build(path, name, plan, *a, **k):
        plans.append(plan)
        return real(path, name, plan, *a, **k)
    monkeypatch.setattr(xlsx, "build", build)
    d = spec_dict()
    d["blocks"].append({"type": "commentary", "key_points": ["Agent point"]})
    pipeline.excel_job(H, _pre(d))
    p = plans[0]
    r = dict(p.zones)["key_points"]
    pts = [p.cells.get((row, TEXT_COL)) for row in range(r + 1, dict(p.zones)["charts"])]
    assert pts[0].startswith('="•  ') and "FIXED(" in pts[0]
    assert "•  Agent point" in pts


def test_the_reply_carries_the_layout_findings_and_the_check_knows_when_sources_failed(fake_target, calculated,
                                                                                       monkeypatch):
    seen = {}

    def check(plan, block, charts, status, sources_ok):
        seen.update(ok=sources_ok, status=status)
        return ["layout: column E too narrow; fix: widen column E"]
    monkeypatch.setattr(pipeline, "_check", check)
    out, _ = _job(_pre(), calculated)
    assert out["layout"] == ["layout: column E too narrow; fix: widen column E"] and seen["ok"] is True
    assert list(seen["status"]) == ["fm"]                       # the Data header's status cell, read in the finish
    monkeypatch.setattr(pipeline, "_write_sources", lambda *a: "could not log the pull on the Sources sheet (x)")
    _job(_pre(), calculated)
    assert seen["ok"] is False


def test_the_real_check_seam_never_raises_on_odd_reads():
    from smartass.render.layout import SheetPlan
    out = pipeline._check(SheetPlan(), None, None, {"fm": ("D13", None)}, True)
    assert out and all(f.startswith("layout:") for f in out)


# ------------------------------------------------------------------ the whole job on the fake Excel
@pytest.fixture
def fake_target(monkeypatch, tmp_home):
    """A new analysis workbook (its first sheet is the empty 'Sources'), our own file: saved by the job."""
    wb = fakexl.Workbook(["Sources"])
    wb.Name, wb.FullName = "2026-11-19_0930 answer.xlsx", str(tmp_home / "analyser" / "answer.xlsx")
    wb.saved = 0

    def save():
        wb.saved += 1
    wb.Save = save
    t = Target(app=wb.Application, wb=wb, pid=0, hwnd=0, created=True)
    monkeypatch.setattr(pipeline, "_target", lambda h, target, title, build_new=False: t)
    monkeypatch.setattr(pipeline, "remember", lambda pre, out: None)
    return t


@pytest.fixture
def calculated(monkeypatch):
    """The fake Excel does not calculate: the real block is read (in the finish batch), then the model's values are
    laid over it where the plan put them, as Excel would show them."""
    real, seen = pipeline.read_block, {}

    def read_block(ws, plan, b):
        seen["plan"] = plan
        get = real(ws, plan, b)

        def block(res):
            got = [list(row) for row in get(res)]
            for key in plan.addr:
                r, c = plan.at(key)
                got[r - 1][c - 1] = seen["pre"].result.cells.get(key)
            return tuple(tuple(row) for row in got)
        return block
    monkeypatch.setattr(pipeline, "read_block", read_block)
    return seen


def _job(pre, calculated):
    calculated["pre"] = pre
    return pipeline.with_timings(lambda h: pipeline.excel_job(h, pre))(H)


def _batches(timings: dict) -> dict:
    """{stage: batches} of a job's Timer summary, the parent stage 'draw' left out."""
    return {k[len(BATCH_PREFIX):]: v for k, v in timings.items() if k.startswith(BATCH_PREFIX)
            and k[len(BATCH_PREFIX):] != "draw"}


def test_the_whole_job_on_the_fake_excel_is_four_batches(fake_target, calculated, monkeypatch, tmp_home):
    """A new answer in a new workbook (its Sources log made and furnished): look, bring in (move in + finish), close
    out, save - one Excel run each (one jobs.com through the COM executor), and the workbook's name read once; the
    bring-in batch's time booked as the stages 'move in' and 'finish'."""
    calls = fakexl.count_com(monkeypatch)
    pre = _pre()
    assert len(pre.spec.queries) == 1
    out, timings = _job(pre, calculated)
    assert _batches(timings) == BATCHES and timings[BATCH_PREFIX + "draw"] == 3
    assert len(calls) == 5, {k: v for k, v in timings.items() if k.startswith(COM_PREFIX)}
    assert timings["draw.move in"] > 0 and timings["draw.finish"] >= 0 and "draw.move in.excel.insert" in timings
    assert out["checks"]["sheet_equals_model"] == "OK" and out["mismatches"] == []
    assert set(STAGES) <= set(timings)
    wb = fake_target.wb
    assert [ws.Name for ws in wb.order] == ["Q1 Answer", "Q1 Data", "Sources", "_svmeta"]
    assert wb.Worksheets("Q1 Answer").Range(calculated["plan"].status_cell).Value == "OK"
    assert [lo.Name for ws in wb.order for lo in ws.ListObjects] == ["Data_Q1"]
    assert wb.saved == 1 and not _temp_files(tmp_home) and len(wb.Application.Workbooks) == 1
    app = wb.Application
    assert (app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation) == (True, True, True, -4105)


@pytest.mark.parametrize("names", [["Sources"], ["Notes", "Sources"],
                                   ["Q1 Answer", "Notes", "Q1 Data", "Sources", "_svmeta"]])
def test_the_answer_lands_where_add_sheet_would_put_it(fake_target, calculated, monkeypatch, names):
    def book():
        wb = fakexl.Workbook(names)
        for ws in wb.order:
            ws.Visible = 0 if ws.Name == "_svmeta" else -1
        return wb
    ref = SimpleNamespace(wb=book(), app=None, sheets_added=[])
    q = sheets.next_question(ref)
    sheets.add_sheet(ref, f"Q{q} Answer", after_q=q)
    want = [ws.Name for ws in ref.wb.order]
    fake_target.wb.order = book().order
    for ws in fake_target.wb.order:
        ws.wb = fake_target.wb
    real, got = fakexl._insert, []

    def insert(dest, path, before, after):
        first = real(dest, path, before, after)
        got.append([ws.Name for ws in dest.order if ws.Name != f"Q{q} Data"])
        return first
    monkeypatch.setattr(fakexl, "_insert", insert)
    _job(_pre(), calculated)
    assert got == [want]


def test_a_stored_answer_without_an_owned_range_is_found_by_a_follow_up(fake_target, calculated):
    pre = _pre()
    _job(pre, calculated)
    stored = sheets._stored_specs(fake_target)
    assert stored[1]["owned"] is None
    assert followup.find_question(stored, None, "Q1 Answer") == (1, pre.raw, None)


def test_a_failed_finish_leaves_the_workbook_and_excel_as_they_were(fake_target, calculated, monkeypatch, tmp_home):
    wb, app = fake_target.wb, fake_target.app
    before = fakexl.snapshot(wb)

    def refuse(self, *a):
        raise fakexl.ComError("a table was refused")
    monkeypatch.setattr(fakexl._ListObjects, "Add", refuse)
    with pytest.raises(BatchError, match="ListObjects.Add.*a table was refused"):
        _job(_pre(), calculated)
    assert fakexl.snapshot(wb) == before and fake_target.sheets_added == []
    assert (app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation) == (True, True, True, -4105)
    assert app.ActiveWorkbook is wb and wb.active is wb.Worksheets("Sources")
    assert not _temp_files(tmp_home) and len(app.Workbooks) == 1 and wb.saved == 0


def test_an_insert_that_fails_half_way_is_rolled_back(fake_target, calculated, monkeypatch, tmp_home):
    """Excel put the Answer in, then the insert failed: the rollback removes it (its name was in t.sheets_added
    before the call), the workbook and Excel are as they were and the temporary file is gone."""
    wb, app = fake_target.wb, fake_target.app
    before = fakexl.snapshot(wb)
    real = fakexl._Sheets.Add

    def half(self, *args):
        real(self, *args)
        for name in ("Q1 Data",):
            self.wb.order.remove(self(name))
        raise fakexl.ComError("Add method of Sheets class failed")
    monkeypatch.setattr(fakexl._Sheets, "Add", half)
    with pytest.raises(BatchError, match="Add method"):
        _job(_pre(), calculated)
    assert fakexl.snapshot(wb) == before and fake_target.sheets_added == []
    assert (app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation) == (True, True, True, -4105)
    assert not _temp_files(tmp_home) and len(app.Workbooks) == 1 and wb.saved == 0


def test_a_lines_block_with_no_lines_draws_an_empty_table(fake_target, calculated):      # Review Focus 5
    ls = LineSet("b0", list(SUPPLIERS), ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=-100.0, cell=LINE_CELL)
    pre = _pre(lines_only_spec(), lines={"b0": ls})
    ls.rows.clear()                                     # prepare refuses an empty set; the job must still draw one
    out, _ = _job(pre, calculated)
    assert out["sheets"] == ["Q1 Answer", "Q1 Data", "Q1 Lines b0"]
    lines = fake_target.wb.Worksheets("Q1 Lines b0")
    (lo,) = lines.ListObjects
    assert lo.Name == "Lines_Q1_b0" and fakexl.rect_text(*lo.rect) == "A4:B5"
    assert [lines.cells[(4, c)]["Value"] for c in (1, 2)] == ["AP_CUSTOMER_NAME", "AMOUNT"]
    assert (5, 1) not in lines.cells or "Value" not in lines.cells[(5, 1)]


def test_the_lines_sheet_carries_todays_title_source_and_provenance(fake_target, calculated, monkeypatch):
    from smartass.render import datasheet
    monkeypatch.setattr(datasheet, "windows_user", lambda: "tom")
    ls = LineSet("b0", list(SUPPLIERS), ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=-100.0, cell=LINE_CELL)
    _job(_pre(lines_only_spec(), lines={"b0": ls}), calculated)
    ws = fake_target.wb.Worksheets("Q1 Lines b0")
    v = lambda a: ws.Range(a).Value                     # noqa: E731
    assert v("A1") == "Ledger lines behind ac.C60 / un.SY.2TEC / Yr 26" and v("A2") == pipeline.LINES_SOURCE
    assert (v("C3"), v("D3")) == ("Cube value", -100.0)
    assert v("F1") == "Drill cell" and (1, 6) in ws.prefixed and v("G2").endswith(" · tom")
    assert v("G3") == "=SUM(Lines_Q1_b0[AMOUNT])" and (3, 7) not in ws.prefixed


# ------------------------------------------------------------------ Task 8B / 2b: one look, few batches, the answer shown
def test_a_second_answer_into_the_workbook_is_four_batches(fake_target, calculated, monkeypatch):
    _job(_pre(), calculated)                            # the first answer furnishes the Sources log
    calls = fakexl.count_com(monkeypatch)
    out, timings = _job(_pre(), calculated)
    assert _batches(timings) == BATCHES and out["checks"]["sheet_equals_model"] == "OK"
    assert len(calls) == 4                              # one per batch: this Target knows its workbook's name already
    assert [ws.Name for ws in fake_target.wb.order] == ["Q1 Answer", "Q2 Answer", "Q1 Data", "Q2 Data", "Sources",
                                                        "_svmeta"]


def test_the_sheets_are_read_once_per_job(fake_target, calculated, monkeypatch):
    """One walk over the sheets (the look's Book), one over the chart sheets; the defined names are not walked
    (re-review M5: 18 s over COM for 2 000 names) but looked up by the answer's table names, once, in the bring-in."""
    _job(_pre(), calculated)
    reads, real = [], runner.VERBS["each"]
    monkeypatch.setitem(runner.VERBS, "each", lambda c, paths: reads.append(list(paths)) or real(c, paths))
    asked, real_exist = [], runner.VERBS["exist"]
    monkeypatch.setitem(runner.VERBS, "exist", lambda c, keys, refuse: asked.append((list(keys), refuse))
                        or real_exist(c, keys, refuse))
    _job(_pre(), calculated)
    assert reads == [movein.BOOK_PATHS, ["Name"]] and asked == [(["Data_Q2"], True)]


def _paints(monkeypatch, app) -> list:
    """Every paint pause of the job, as an event of the fake Excel (in order with the insert, the formulas and the
    calculation), with the screen updating Excel had then."""
    def sleep(s):
        assert s == movein.PAINT_PAUSE > 0
        app.events.append(("", "paint", app.ScreenUpdating))
    import time
    monkeypatch.setattr(pipeline, "time", SimpleNamespace(sleep=sleep, perf_counter=time.perf_counter))
    return app.events


def test_the_answer_is_painted_once_with_its_numbers_before_the_close_out(fake_target, calculated, monkeypatch):
    """Controller ruling R1/R2: the move-in and the finish are one batch (a boundary between them painted a half-drawn
    sheet in the owner's Excel), then ONE pause (movein.PAINT_PAUSE, 0.15 s) for Excel to paint the answer with its
    numbers and chart, screen updating already back on, before the close-out batch."""
    app = fake_target.app
    events = _paints(monkeypatch, app)
    _job(_pre(), calculated)
    kinds = [k for _, k, *_ in events if k in ("Sheets.Add", "Formula2", "Calculate", "paint", "Value2")]
    assert kinds.count("paint") == 1
    paint = kinds.index("paint")
    assert kinds.index("Sheets.Add") < paint and max(i for i, k in enumerate(kinds) if k == "Calculate") < paint
    assert "Value2" in kinds[paint:]                    # the close-out (the status cell) comes after the paint
    assert [e[2] for e in events if e[1] == "paint"] == [True]
    assert app.ActiveWorkbook is fake_target.wb and fake_target.wb.active.Name == "Q1 Answer"
    assert app.ScreenUpdating is True and movein.PAINT_PAUSE == 0.15


def test_sv_paint_pause_0_means_no_pause(fake_target, calculated, monkeypatch):
    monkeypatch.setattr(movein, "PAINT_PAUSE", 0.0)
    events = _paints(monkeypatch, fake_target.app)
    out, timings = _job(_pre(), calculated)
    assert not [e for e in events if e[1] == "paint"] and "draw.paint" not in timings


def test_a_workbook_that_is_not_on_screen_is_never_painted_and_the_users_workbook_stays_active(fake_target,
                                                                                              calculated,
                                                                                              monkeypatch):
    app = fake_target.app
    mine = fakexl.Workbook(["Mine"], app=app)               # the user works in another workbook meanwhile
    mine.Name, mine.FullName = "mine.xlsx", r"C:\Users\me\mine.xlsx"
    events = _paints(monkeypatch, app)
    _job(_pre(), calculated)
    assert not [e for e in events if e[1] == "paint"] and app.ActiveWorkbook is mine and mine.active.Name == "Mine"


def test_the_active_workbook_of_an_excel_that_is_not_visible_ends_on_the_new_answer(fake_target, calculated,
                                                                                   monkeypatch):
    """Task 8B review 1: the close-out lands on the new Answer whenever the target is the active workbook; a hidden
    Excel is not painted (no pause) but its file opens on the answer."""
    app = fake_target.app
    app.Visible = False
    events = _paints(monkeypatch, app)
    out, _ = _job(_pre(), calculated)
    assert out["checks"]["sheet_equals_model"] == "OK" and not [e for e in events if e[1] == "paint"]
    assert app.ActiveWorkbook is fake_target.wb and fake_target.wb.active.Name == "Q1 Answer"


def test_the_draw_stamps_one_time_for_the_data_header_and_the_answer_footer(wired, monkeypatch):
    """Moved from the COM writer's tests: the Data header's Fetched and the Answer's print footer never disagree."""
    import datetime

    from smartass.render import datasheet

    class Fixed(datetime.datetime):
        @classmethod
        def now(cls, tz=None):
            return cls(2026, 11, 4, 23, 59, 30, tzinfo=datetime.timezone(datetime.timedelta(hours=1)))

        def astimezone(self, tz=None):  # the machine's zone (UTC on CI) must not move the pinned Oslo time
            return self if tz is None else super().astimezone(tz)
    seen = {}
    real = xlsx.build

    def build(path, name, plan, data, lines, **k):
        seen.update(plan=plan, header=data.header)
        return real(path, name, plan, data, lines, **k)
    monkeypatch.setattr(xlsx, "build", build)
    monkeypatch.setattr(pipeline, "datetime", Fixed)
    monkeypatch.setattr(datasheet, "windows_user", lambda: "tom")
    pipeline.excel_job(H, _pre())
    assert dict(seen["header"])["Fetched"] == "2026-11-04 23:59 +0100 · tom"
    assert seen["plan"].footer == "Source: Oracle EPM Cloud · fetched 2026-11-04 · page &P of &N"


def test_the_job_opens_no_workbook_the_sheets_come_in_with_one_insert(fake_target, calculated, tmp_home):
    """Task 2a: no temporary workbook to open, close or unlink (the 'would not close' path is gone with it)."""
    app = fake_target.app
    out, _ = _job(_pre(), calculated)
    kinds = [k for _, k, *_ in app.events]
    assert kinds.count("Sheets.Add") == 1 and not {"Open", "Copy", "Close", "BreakLink"} & set(kinds)
    assert not out.get("warnings") and len(app.Workbooks) == 1 and not _temp_files(tmp_home)


# ------------------------------------------------------------------ the Key point heights and the page breaks
def _point_plan(text: str):
    """A narrow answer (A 2, B 9, C 46, D 16) with two Key point formulas, and the block Excel shows for it."""
    from smartass.render.layout import SheetPlan
    p = SheetPlan(last_col="D", last_row=3)
    p.cells[(1, 2)] = "Title"
    p.rows.append({"row": 1, "kind": "title"})
    for r in (2, 3):
        p.cells[(r, 2)] = '="•  Opex "&FIXED(D9,0)&" NOK"'
        p.rows.append({"row": r, "kind": "text"})
    p.footer = "Source: Oracle EPM Cloud · fetched 2026-11-05 · page &P of &N"
    return p, ((None, "Title", None, None), (None, text, None, None), (None, text, None, None))


def fit(ws, plan, block) -> list:
    """pipeline._fit_text_rows recorded on the fake sheet ws and run as a batch, then the file's breaks again when
    Excel refused one (_refused_breaks); the batch's operations."""
    t = SimpleNamespace(app=ws.wb.Application, wb=ws.wb)
    b = Batch()
    breaks = pipeline._fit_text_rows(b.workbook(ws.wb.Name).Worksheets(ws.Name), plan, block, b)
    res = runner.run(t.app, b)
    pipeline._refused_breaks(t, ws.Name, breaks, res)
    return b.to_dict()["ops"]


def test_key_point_rows_get_the_height_of_their_calculated_text_in_one_operation_per_run(monkeypatch):
    long = "•  Opex 1 234 567 NOK vs PY; " + "consultants, licences and the new data centre " * 3
    plan, block = _point_plan(long)
    ws = fakexl.Workbook(["Q1 Answer"]).order[0]
    want = look.text_heights(plan, look.answer_look(plan), block)
    assert set(want) == {2, 3} and len(set(want.values())) == 1
    ops = fit(ws, plan, block)
    assert {r: ws.rowh[r] for r in (2, 3)} == want and [x for x in ws.log if x[0] == "RowHeight"] == [
        ("RowHeight", "2:3")]
    assert [op["name"] for op in ops if op["k"] == "set"] == ["RowHeight"]   # the breaks did not move: none set


def _calculated_sheet(case):
    """A laid-out case of test_look on the fake sheet with the values Excel would show, and its block."""
    plan, values = case()
    ws = fakexl_sheet_with(plan, values)
    return plan, values, ws, ws.Range("A1:" + plan.last_col + str(plan.last_row)).Value2


def test_page_breaks_are_reset_when_the_calculated_key_points_move_them():
    plan, values, ws, block = _calculated_sheet(lines_ranking)
    full = look.answer_look(plan, texts={rc: v for rc, v in values.items() if isinstance(v, str)})
    fit(ws, plan, block)
    assert ws.page_breaks == full.breaks != look.answer_look(plan).breaks
    assert {r: ws.rowh[r] for r in full.text_rows} == {r: full.heights[r] for r in full.text_rows}


def test_page_breaks_are_left_alone_when_they_do_not_move():
    plan, _, ws, block = _calculated_sheet(pl_table)
    fit(ws, plan, block)
    assert ws.page_breaks is None


def test_a_printer_that_refuses_the_breaks_does_not_fail_the_answer(caplog):
    plan, _, ws, block = _calculated_sheet(lines_ranking)

    def refuse():
        raise fakexl.ComError("no printer")
    ws.ResetAllPageBreaks = refuse
    with caplog.at_level(logging.WARNING):
        fit(ws, plan, block)
    assert ws.page_breaks is None and any("page breaks" in r.getMessage() for r in caplog.records)


# ------------------------------------------------------------------ server.py: one renderer
def test_a_new_question_is_drawn_by_excel_job(monkeypatch):
    seen = []

    async def user(fn, what):
        return fn(None)

    async def prepare(raw, provider, spec=None, target="session", complete=False):
        return SimpleNamespace(spec=spec), []
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "_provider", lambda: "provider")
    monkeypatch.setattr(server, "_prepare", prepare)
    monkeypatch.setattr(server, "excel_job", lambda h, pre, target="session": seen.append("excel_job") or {"q": 1})
    from smartass.model.spec import AnalysisSpec
    asyncio.run(server._new_question(AnalysisSpec.model_validate(spec_dict()), "session"))
    assert seen == ["excel_job"]


def test_there_is_one_renderer_and_no_switch_for_it():
    assert not hasattr(config, "RENDER") and not hasattr(server, "_answer_job")
    assert not hasattr(pipeline, "file_job") and not hasattr(pipeline, "file_redraw_job")


# ------------------------------------------------------------------ phase 3b Task 2b: the batches
def _runs(monkeypatch) -> list:
    runs, real = [], runner.run
    monkeypatch.setattr(runner, "run", lambda app, b, *a, **k: runs.append(b) or real(app, b, *a, **k))
    return runs


def test_the_look_and_the_save_are_not_quiet_and_only_the_bring_in_sets_manual_calculation(fake_target, calculated,
                                                                                          monkeypatch):
    """Each quiet batch turns the settings off and puts the user's back itself; the look only reads and the save runs
    with the user's settings. Review M6: only the batch that writes the formulas sets Calculation to manual (putting
    automatic back recalculates what is dirty in every open workbook)."""
    runs = _runs(monkeypatch)
    _job(_pre(), calculated)
    assert [(b.quiet, b.calculation) for b in runs] == [(False, True), (True, True), (True, False), (False, True)]
    app = fake_target.app
    assert app.Calculation == -4105


def test_a_defined_name_an_answer_table_wants_stops_the_bring_in_before_anything(fake_target, calculated,
                                                                                 monkeypatch):
    """Re-review M5: the answer's table names are looked up as the bring-in batch's first operation (a new answer's
    number comes from the look); one that is taken fails the batch there - before anything is inserted - and is the
    NameTaken it was, with its text; no clean-up batch is needed (nothing came in)."""
    fake_target.wb.Names.Add("data_q1", "=Sources!$A$1", False)
    before = fakexl.snapshot(fake_target.wb)
    runs = _runs(monkeypatch)
    with pytest.raises(movein.NameTaken) as e:
        _job(_pre(), calculated)
    assert str(e.value) == (f"the name 'data_q1' already exists in {fake_target.wb.Name}: rename or delete it, then "
                            "ask again")                       # the workbook's own spelling, as before
    assert len(runs) == 2 and fakexl.snapshot(fake_target.wb) == before and fake_target.sheets_added == []
    assert not [e for e in fake_target.app.events if e[1] == "Sheets.Add"]
    ops = runs[1].to_dict()["ops"]
    first = next(i for i, op in enumerate(ops) if op["k"] in ("set", "verb") or op["name"] == "Add")
    assert ops[first]["k"] == "verb" and ops[first]["name"] == "exist"


def test_a_frozen_header_excel_refuses_is_undone_and_no_warning(fake_target, calculated, monkeypatch):
    """A window without a client area (a hidden Excel, live 20.11.2026) refuses FreezePanes: as styling._freeze,
    the split is taken back and the log is complete (no Sources warning); the Answer stays the active sheet."""
    real = fakexl._Window.__setattr__

    def setattr_(win, k, v):
        if k == "FreezePanes" and v is True:
            raise fakexl.ComError("Unable to set the FreezePanes property of the Window class")
        real(win, k, v)
    monkeypatch.setattr(fakexl._Window, "__setattr__", setattr_)
    out, _ = _job(_pre(), calculated)
    wb = fake_target.wb
    assert not [w for w in out["warnings"] if "Sources" in w] and not [x for x in out["layout"] if "Sources" in x]
    assert wb.window_state["Sources"]["SplitRow"] == 0 and wb.window_state["Sources"]["FreezePanes"] is False
    assert wb.active.Name == "Q1 Answer"


def test_the_labels_are_positioned_before_the_calculation_and_lifted_right_after_it(fake_target, calculated,
                                                                                   monkeypatch):
    """Live (hidden Excel, 20.11.2026, scripts/check_insert.py --parts labels): a point label positioned after the
    calculation lands 12-24 pt elsewhere than before it, and a lift set before the calculation is lost (the label
    ends where Excel puts it); so the labels are positioned and coloured before Calculate and lifted right after it,
    inside the same batch (review M3: the verb nudge; no label jumps after a paint)."""
    seen = []
    for name in ("fix_labels", "calculate", "lift_labels"):
        real = getattr(movein, name)
        monkeypatch.setattr(movein, name, lambda *a, real=real, name=name, **k: seen.append(name) or real(*a, **k))
    runs = _runs(monkeypatch)
    _job(_pre(), calculated)
    assert seen == ["fix_labels", "calculate", "lift_labels"] and len(runs) == 4


def test_fix_labels_positions_and_colours_and_lift_labels_nudges_inside_excel():
    from smartass.render.xlsx_charts import LabelFix
    built = SimpleNamespace(sheets=["Q1 Answer"], label_fixes=[LabelFix(1, 2, 3, position="right"),
                                                             LabelFix(1, 1, 2, font_color="#1F3864", lift_pt=10.0),
                                                             LabelFix(1, 1, 3, lift_pt=10.0)])
    b = Batch()
    wbn = b.workbook("book.xlsx")
    spans = movein.fix_labels(b, wbn, built)
    sets = [(op["name"], op["value"]) for op in b.to_dict()["ops"] if op["k"] == "set"]
    assert sets == [("Position", -4152), ("Color", 0x64381F)] and len(spans) == 2
    b = Batch()
    spans = movein.lift_labels(b, b.workbook("book.xlsx"), built)
    verbs = [op for op in b.to_dict()["ops"] if op["k"] == "verb"]
    assert [(v["name"], v["args"][1:]) for v in verbs] == [("nudge", ["Top", -10.0, 10.0])] * 2 and len(spans) == 2
    assert all(v.get("try") for v in verbs)


def test_in_the_users_workbook_there_is_no_save_batch(fake_target, calculated, monkeypatch):
    """Review I2: the user's workbook is never saved; with no new Sources log to name, no save batch runs at all
    (a batch that does nothing must not put a finished answer at risk): look, bring in, close out."""
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: False)
    _job(_pre(), calculated)                            # the first answer's new log: its print setup named in a batch
    runs = _runs(monkeypatch)
    out, timings = _job(_pre(), calculated)
    assert len(runs) == 3 and BATCH_PREFIX + "save" not in timings and fake_target.wb.saved == 0
    assert out["checks"]["sheet_equals_model"] == "OK"


def test_a_users_sheet_named_sources_in_another_case_is_found_and_left_alone(fake_target, calculated):
    """Re-review: Excel's sheet names are not case-sensitive, so the user's own 'sources' sheet is the 'Sources'
    sheet the look reads the kind of: not ours, so the pull goes to a new 'Sources SV' log and no second
    'Sources' is made."""
    src = fake_target.wb.Worksheets("Sources")
    src.Name = "sources"
    src.put(1, 1, "My own list")
    before = fakexl.sheet_state(src)
    out, _ = _job(_pre(), calculated)
    names = [ws.Name for ws in fake_target.wb.order]
    assert fakexl.sheet_state(src) == before and [n for n in names if n.lower() == "sources"] == ["sources"]
    assert "Sources SV" in names and not [w for w in out["warnings"] if "Sources" in w]
    assert out["checks"]["sheet_equals_model"] == "OK"


def test_a_sources_sheet_whose_kind_cannot_be_read_is_not_taken_for_the_log(fake_target, calculated, monkeypatch):
    """Review M4: when the kind of an existing 'Sources' sheet cannot be read, the pull is not logged (a warning, as
    before) and the sheet is not furnished as if it were ours."""
    src = fake_target.wb.Worksheets("Sources")
    src.put(1, 1, "My own list")
    before = fakexl.sheet_state(src)
    real, failed = src.Range, []

    def range_(a, b=None):
        if a == "A1:B1" and not failed:                 # the look's read only: later reads would work
            failed.append(a)
            raise fakexl.ComError("Excel would not read it")
        return real(a, b)
    src.Range = range_
    out, _ = _job(_pre(), calculated)
    del src.Range
    assert fakexl.sheet_state(src) == before and "Sources SV" not in [ws.Name for ws in fake_target.wb.order]
    assert any(w.startswith("could not log the pull on the Sources sheet") for w in out["warnings"])
    assert out["checks"]["sheet_equals_model"] == "OK"
