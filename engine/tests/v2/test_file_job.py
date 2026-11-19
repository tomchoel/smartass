# engine/tests/v2/test_file_job.py
"""Phase 3 Task 5: pipeline.file_job, a new answer through a built file (render/xlsx.build, then xl/movein): the
steps in order, the rollback and the temporary file on a failure, the COM calls of a whole job on the fake Excel,
the reply as excel_job gives it, and config.RENDER picking the job in server.py. All on fakes; every figure is
invented."""
from __future__ import annotations

import asyncio
import contextlib
import logging
from types import SimpleNamespace

import pytest
from test_look import fakexl_sheet_with, lines_ranking, pl_table
from test_v2_pipeline import LINE_CELL, FakeProvider, lines_only_spec
from v2util import facts, resolver, spec_dict

import fakexl
from smartass import followup, pipeline, server
from smartass.core import config
from smartass.core.timing import COM_PREFIX
from smartass.model.errors import GuardError
from smartass.model.lines import LineSet
from smartass.render import look, xlsx
from smartass.xl import movein, sheets
from smartass.xl.host import HostError
from smartass.xl.target import Target

H = None                                        # the host: every seam that would use it is replaced
STAGES = ["open target", "draw", "draw.build", "draw.move in", "draw.finish", "draw.read back", "draw.sources",
          "draw.arrange", "save"]
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
class Wired:
    """Every step of file_job replaced by a recorder (the build is the real one: the temporary file is real)."""

    def __init__(self, monkeypatch):
        self.calls: list[str] = []
        self.failures: dict[str, Exception] = {}
        self.rolled_back: list[str] | None = None
        wb = fakexl.Workbook(["Sources"], styles_ver=False)
        self.t = Target(app=wb.Application, wb=wb, pid=0, hwnd=0, created=True)
        self.answer, self.data = fakexl.Workbook(["Q1 Answer", "Q1 Data"]).order
        self.built: list[xlsx.BuiltFile] = []
        real_build = xlsx.build

        def build(path, *a):
            self.path = path
            self.built.append(real_build(path, *a))
            return self.built[-1]

        def move_in(t, built, after, first=False):
            t.sheets_added.extend(built.sheets)

        def rollback(t):
            self.rolled_back = list(t.sheets_added)
        put = monkeypatch.setattr
        put(pipeline, "_target", self.rec("target", ret=self.t))
        put(pipeline, "_next_q", self.rec("next_q", ret=1))
        put(xlsx, "build", self.rec("build", build))
        put(movein, "quiet", lambda app: contextlib.nullcontext())
        put(pipeline, "_after", lambda t, q: None)
        put(movein, "move_in", self.rec("move_in", move_in))
        put(movein, "finish", self.rec("finish"))
        put(movein, "calculate", self.rec("calculate"))
        put(pipeline, "_answer_sheets", lambda t, q: (self.answer, self.data))
        put(pipeline, "read_block", self.rec("read_block", ret=()))
        put(pipeline, "compare", lambda *a: [])
        put(pipeline, "write_status", self.rec("status", ret="OK"))
        put(pipeline, "_fit_text_rows", lambda ws, plan, block: None)
        put(pipeline, "_write_sources", self.rec("sources"))
        put(pipeline, "_save_spec", self.rec("save_spec"))
        put(pipeline, "_arrange", self.rec("arrange"))
        put(pipeline, "_check", self.rec("check", ret=[]))
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
    out = pipeline.file_job(H, _pre())
    assert wired.calls == ["target", "next_q", "build", "move_in", "finish", "calculate", "read_block", "status",
                           "sources", "save_spec", "arrange", "check", "save", "remember"]
    assert out["sheets"] == ["Q1 Answer", "Q1 Data"] and out["checks"]["sheet_equals_model"] == "OK"
    assert wired.path.parent == tmp_home / "tmp" and wired.path.suffix == ".xlsx" and not wired.path.exists()


def test_the_temporary_file_has_a_new_name_every_time(wired):
    pipeline.file_job(H, _pre())
    first = wired.path
    pipeline.file_job(H, _pre())
    assert wired.path != first and len(wired.path.stem) == 32                  # a uuid4 hex


def test_a_failure_after_the_move_in_removes_the_new_sheets_and_the_temp_file(wired, tmp_home):   # Review Focus 3
    wired.fail("finish", HostError("Excel is busy"))
    with pytest.raises(HostError, match="busy"):
        pipeline.file_job(H, _pre())
    assert wired.rolled_back == ["Q1 Answer", "Q1 Data"] and not _temp_files(tmp_home)
    assert wired.calls.count("rollback") == 1
    assert "save" not in wired.calls and "remember" not in wired.calls


def test_a_build_that_fails_leaves_no_partial_file_and_copies_nothing(wired, tmp_home, monkeypatch):
    def half(path, *a):
        path.write_bytes(b"PK half a file")
        raise ValueError("the build broke")
    monkeypatch.setattr(xlsx, "build", wired.rec("build", half))
    with pytest.raises(ValueError, match="broke"):
        pipeline.file_job(H, _pre())
    assert "move_in" not in wired.calls and wired.calls[-1] == "rollback" and not _temp_files(tmp_home)


def test_more_lines_than_a_sheet_has_rows_is_a_refusal_with_its_fix(wired, tmp_home, monkeypatch):  # Focus 5
    def too_many(path, *a):
        raise xlsx.TooManyRows("sheet 'Q1 Lines b0' would need 1 100 000 rows; an Excel sheet has 1 048 576")
    monkeypatch.setattr(xlsx, "build", wired.rec("build", too_many))
    with pytest.raises(GuardError, match="1 100 000 rows") as e:
        pipeline.file_job(H, _pre())
    assert "narrow" in e.value.fix and wired.calls[-1] == "rollback" and not _temp_files(tmp_home)


def test_the_spec_is_stored_without_an_owned_range(wired, monkeypatch):
    seen = {}
    monkeypatch.setattr(pipeline, "_save_spec", lambda t, q, raw, owned: seen.update(q=q, raw=raw, owned=owned))
    pre = _pre()
    pipeline.file_job(H, pre)
    assert seen == {"q": 1, "raw": pre.raw, "owned": None}


def test_the_chart_lint_and_the_drawn_charts_reach_the_reply_once(wired, monkeypatch):
    real = pipeline.layout

    def layout(*a, **k):
        plan = real(*a, **k)
        plan.chart_lint.append("chart 'P&L': a lint line")
        return plan
    monkeypatch.setattr(pipeline, "layout", layout)
    out = pipeline.file_job(H, _pre())
    assert out["warnings"].count("chart 'P&L': a lint line") == 1
    assert out["charts"] == wired.built[0].charts and out["charts"]


def test_the_reply_has_the_keys_of_excel_jobs_reply(wired):
    pre = _pre()
    out = pipeline.file_job(H, pre)
    assert set(out) == set(pipeline.reply(pre, 1, "x", [], [], [])) and out["q"] == 1


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
    monkeypatch.setattr(pipeline, "_target", lambda h, target, title: t)
    monkeypatch.setattr(pipeline, "remember", lambda pre, out: None)
    return t


@pytest.fixture
def calculated(monkeypatch):
    """The fake Excel does not calculate: the real block is read (one COM call), then the model's values are laid
    over it where the plan put them, as Excel would show them."""
    real, seen = pipeline.read_block, {}

    def read_block(ws, plan):
        seen["plan"] = plan
        block = [list(row) for row in real(ws, plan)]
        for key in plan.addr:
            r, c = plan.at(key)
            block[r - 1][c - 1] = seen["pre"].result.cells.get(key)
        return tuple(tuple(row) for row in block)
    monkeypatch.setattr(pipeline, "read_block", read_block)
    return seen


def _job(pre, calculated):
    calculated["pre"] = pre
    return pipeline.with_timings(lambda h: pipeline.file_job(h, pre))(H)


def test_the_whole_job_on_the_fake_excel_makes_at_most_45_com_calls(fake_target, calculated, monkeypatch, tmp_home):
    calls = fakexl.count_com(monkeypatch)
    pre = _pre()
    assert len(pre.spec.queries) == 1
    out, timings = _job(pre, calculated)
    sources = timings.get(COM_PREFIX + "draw.sources", 0)
    assert 0 < sources and len(calls) - sources <= 45, {k: v for k, v in timings.items() if k.startswith(COM_PREFIX)}
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
    real, got = movein.move_in, []

    def move_in(t, built, after, first=False):
        real(t, built, after, first)
        got.append([ws.Name for ws in t.wb.order if ws.Name != f"Q{q} Data"])
    monkeypatch.setattr(movein, "move_in", move_in)
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

    def finish(t, built):
        raise HostError("a formula was refused")
    monkeypatch.setattr(movein, "finish", finish)
    with pytest.raises(HostError, match="refused"):
        _job(_pre(), calculated)
    assert fakexl.snapshot(wb) == before and fake_target.sheets_added == []
    assert (app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation) == (True, True, True, -4105)
    assert app.ActiveWorkbook is wb and wb.active is wb.Worksheets("Sources")
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


def test_key_point_rows_get_the_height_of_their_calculated_text_in_one_call_per_run(monkeypatch):
    calls = fakexl.count_com(monkeypatch)
    long = "•  Opex 1 234 567 NOK vs PY; " + "consultants, licences and the new data centre " * 3
    plan, block = _point_plan(long)
    ws = fakexl.Workbook(["Q1 Answer"]).order[0]
    want = look.text_heights(plan, look.answer_look(plan), block)
    assert set(want) == {2, 3} and len(set(want.values())) == 1
    pipeline._fit_text_rows(ws, plan, block)
    assert {r: ws.rowh[r] for r in (2, 3)} == want and [x for x in ws.log if x[0] == "RowHeight"] == [
        ("RowHeight", "2:3")]
    assert len(calls) == 1                              # the breaks did not move: no page-break call


def _calculated_sheet(case):
    """A laid-out case of test_look on the fake sheet with the values Excel would show, and its block."""
    plan, values = case()
    ws = fakexl_sheet_with(plan, values)
    return plan, values, ws, ws.Range("A1:" + plan.last_col + str(plan.last_row)).Value2


def test_page_breaks_are_reset_when_the_calculated_key_points_move_them():
    plan, values, ws, block = _calculated_sheet(lines_ranking)
    full = look.answer_look(plan, texts={rc: v for rc, v in values.items() if isinstance(v, str)})
    pipeline._fit_text_rows(ws, plan, block)
    assert ws.page_breaks == full.breaks != look.answer_look(plan).breaks
    assert {r: ws.rowh[r] for r in full.text_rows} == {r: full.heights[r] for r in full.text_rows}


def test_page_breaks_are_left_alone_when_they_do_not_move():
    plan, _, ws, block = _calculated_sheet(pl_table)
    pipeline._fit_text_rows(ws, plan, block)
    assert ws.page_breaks is None


def test_a_printer_that_refuses_the_breaks_does_not_fail_the_answer(caplog):
    plan, _, ws, block = _calculated_sheet(lines_ranking)

    def refuse():
        raise fakexl.ComError("no printer")
    ws.ResetAllPageBreaks = refuse
    with caplog.at_level(logging.WARNING):
        pipeline._fit_text_rows(ws, plan, block)
    assert ws.page_breaks is None and any("page breaks" in r.getMessage() for r in caplog.records)


# ------------------------------------------------------------------ server.py picks the job
@pytest.mark.parametrize("render, job", [("com", "excel_job"), ("file", "file_job")])
def test_config_render_picks_the_job(monkeypatch, render, job):
    seen = []

    async def user(fn, what):
        return fn(None)

    async def prepare(raw, provider, spec=None, target="session", complete=False):
        return SimpleNamespace(spec=spec), []
    monkeypatch.setattr(config, "RENDER", render)
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "_provider", lambda: "provider")
    monkeypatch.setattr(server, "_prepare", prepare)
    monkeypatch.setattr(server, "excel_job", lambda h, pre, target="session": seen.append("excel_job") or {"q": 1})
    monkeypatch.setattr(server, "file_job", lambda h, pre, target="session": seen.append("file_job") or {"q": 1})
    from smartass.model.spec import AnalysisSpec
    asyncio.run(server._new_question(AnalysisSpec.model_validate(spec_dict()), "session"))
    assert seen == [job] and server._answer_job() is getattr(server, job)


def test_the_default_renderer_is_com(monkeypatch):
    import importlib
    monkeypatch.delenv("SV_RENDER", raising=False)
    fresh = importlib.util.spec_from_file_location("config_fresh", config.__file__)
    mod = importlib.util.module_from_spec(fresh)
    fresh.loader.exec_module(mod)
    assert mod.RENDER == "com"
    monkeypatch.setenv("SV_RENDER", "file")
    fresh.loader.exec_module(mod)
    assert mod.RENDER == "file"
