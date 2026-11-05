# engine/tests/v2/test_v2_pipeline_job.py
import pytest
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import pipeline


class T:
    path = "C:/a/v2/x.xlsx"
    created = True
    wb = object()
    app = object()


def wire(monkeypatch, log, fail_at=None):
    def rec(name, ret=None):
        def f(*a, **k):
            log.append(name)
            if name == fail_at:
                raise RuntimeError("boom")
            return ret
        return f
    monkeypatch.setattr(pipeline, "_target", rec("target", T()))
    monkeypatch.setattr(pipeline, "_next_q", rec("next_q", 4))
    monkeypatch.setattr(pipeline, "_add_sheet", rec("add_sheet", object()))
    monkeypatch.setattr(pipeline, "write_data", rec("write_data"))
    monkeypatch.setattr(pipeline, "write_answer", rec("write_answer", []))
    monkeypatch.setattr(pipeline, "write_charts", rec("write_charts", []))
    monkeypatch.setattr(pipeline, "_calculate", rec("calculate"))
    monkeypatch.setattr(pipeline, "read_block", rec("read_block", ()))
    monkeypatch.setattr(pipeline, "compare", rec("compare", []))
    monkeypatch.setattr(pipeline, "_save_spec", rec("save_spec"))
    monkeypatch.setattr(pipeline, "_arrange", rec("arrange"))
    monkeypatch.setattr(pipeline, "_save", rec("save"))
    monkeypatch.setattr(pipeline, "_rollback", rec("rollback"))


def test_job_order(monkeypatch):
    log = []
    wire(monkeypatch, log)
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    out = pipeline.excel_job(None, pre, target="session")
    assert log == ["target", "next_q", "add_sheet", "add_sheet", "write_data", "write_answer", "write_charts", "calculate",
                   "read_block", "compare", "save_spec", "arrange", "save"]
    assert out["q"] == 4 and out["sheets"] == ["Q4 Answer", "Q4 Data"]


def test_job_rolls_back_on_failure(monkeypatch):
    log = []
    wire(monkeypatch, log, fail_at="write_answer")
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    with pytest.raises(RuntimeError):
        pipeline.excel_job(None, pre, target="session")
    assert log[-1] == "rollback"


def test_arrange_warning_goes_to_the_reply_without_rollback(monkeypatch):
    log = []
    wire(monkeypatch, log)
    monkeypatch.setattr(pipeline, "_arrange", lambda t, sheet: f"arrange failed for {sheet}")
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    out = pipeline.excel_job(None, pre, target="session")
    assert "arrange failed for Q4 Answer" in out["warnings"] and "rollback" not in log


def test_data_tab_colour_is_v1s():
    from smartass.xl import sheetorder
    assert pipeline.DATA_TAB == sheetorder.DATA_TAB == 0xA6A6A6


def test_job_draws_charts_and_saves_owned_range(monkeypatch):
    log, saved = [], {}
    wire(monkeypatch, log)
    monkeypatch.setattr(pipeline, "_save_spec", lambda t, q, raw, owned: saved.update(q=q, owned=owned))
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    pipeline.excel_job(None, pre, target="session")
    assert log.index("write_answer") < log.index("write_charts") < log.index("calculate")
    assert saved["q"] == 4 and saved["owned"]["rows"] > 10 and saved["owned"]["cols"] >= 3


def test_job_writes_each_line_set_before_the_answer(monkeypatch):
    from test_v2_pipeline import LINE_CELL, lines_only_spec

    from smartass.model.lines import LineSet
    log = []
    wire(monkeypatch, log)
    monkeypatch.setattr(pipeline, "_write_lines", lambda t, q, ls: log.append(f"write_lines {q} {ls.alias}"))
    rows = [{"AP_CUSTOMER_NAME": "Bouvet", "AMOUNT": -30.0}]
    ls = LineSet("b0", rows, ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=-30.0, cell=LINE_CELL)
    pre = pipeline.prepare(lines_only_spec(), resolver(), FakeProvider(facts()), lines={"b0": ls})
    out = pipeline.excel_job(None, pre, target="session")
    assert log.index("write_data") < log.index("write_lines 4 b0") < log.index("write_answer")
    assert log.count("write_lines 4 b0") == 1
    assert out["sheets"] == ["Q4 Answer", "Q4 Data", "Q4 Lines b0"]


def test_job_writes_live_key_points_and_free_points(monkeypatch):
    from smartass.render.layout import TEXT_COL
    log, plans = [], []
    wire(monkeypatch, log)
    monkeypatch.setattr(pipeline, "write_answer", lambda ws, plan, warnings: plans.append(plan) or [])
    d = spec_dict()
    d["blocks"].append({"type": "commentary", "key_points": ["Agent point"]})
    pre = pipeline.prepare(d, resolver(), FakeProvider(facts()))
    pipeline.excel_job(None, pre, target="session")
    p = plans[0]
    r = dict(p.zones)["key_points"]
    pts = [p.cells.get((row, TEXT_COL)) for row in range(r + 1, dict(p.zones)["charts"])]
    assert pts[0].startswith('="•  ') and "FIXED(" in pts[0]
    assert "•  Agent point" in pts
