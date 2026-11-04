# engine/tests/v2/test_v2_redraw.py
import pytest
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import pipeline

ROOM = {}
OLD = (("Cube", "Value"), ("FinMax", 1.0))


class T:
    path = "C:/a/v2/x.xlsx"
    created = True


def wire(monkeypatch, log, fail_at=None, restore_fail=()):
    def rec(name, ret=None):
        def f(*a, **k):
            log.append(name)
            if name == fail_at:
                raise RuntimeError("boom")
            return ret
        return f

    def restore(what):
        def f(*a, **k):
            log.append(what)
            if what in restore_fail:
                raise RuntimeError(f"{what} failed")
        return f
    monkeypatch.setattr(pipeline, "_target", rec("target", T()))
    monkeypatch.setattr(pipeline, "_backup", lambda t, name: log.append(f"backup {name}") or f"~{name}")
    monkeypatch.setattr(pipeline, "_sheet", rec("sheet", object()))
    monkeypatch.setattr(pipeline, "_read_data", rec("read_old_data", OLD))
    monkeypatch.setattr(pipeline, "_clear", rec("clear"))
    monkeypatch.setattr(pipeline, "_add_sheet", rec("add_sheet", object()))
    monkeypatch.setattr(pipeline, "write_data", rec("write_data"))
    monkeypatch.setattr(pipeline, "rewrite_data", rec("rewrite_data"))
    monkeypatch.setattr(pipeline, "rewrite_grid", lambda ws, grid, q: log.append(f"rewrite_old {grid == OLD}"))
    monkeypatch.setattr(pipeline, "_make_room",
                        lambda ws, old, new: log.append("make_room") or ROOM.update(old=old, new=new))
    monkeypatch.setattr(pipeline, "write_answer", rec("write_answer", []))
    monkeypatch.setattr(pipeline, "write_charts", rec("write_charts", []))
    monkeypatch.setattr(pipeline, "_calculate", rec("calculate"))
    monkeypatch.setattr(pipeline, "read_block", rec("read_block", ()))
    monkeypatch.setattr(pipeline, "compare", rec("compare", []))
    monkeypatch.setattr(pipeline, "_save_spec", rec("save_spec"))
    monkeypatch.setattr(pipeline, "_arrange", rec("arrange", []))
    monkeypatch.setattr(pipeline, "_drop_backup", lambda t, b: log.append(f"drop {b}"))
    monkeypatch.setattr(pipeline, "_restore", lambda t, name, b: restore("restore " + name)())
    monkeypatch.setattr(pipeline, "_rollback", restore("rollback"))
    monkeypatch.setattr(pipeline, "_save", rec("save"))


def test_redraw_order(monkeypatch):
    log = []
    wire(monkeypatch, log)
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    out = pipeline.redraw_job(None, pre, q=2, owned={"rows": 30, "cols": 6}, target="session")
    assert log[:5] == ["target", "sheet", "read_old_data", "backup Q2 Answer", "sheet"]
    assert log.index("clear") < log.index("rewrite_data") < log.index("write_answer")
    assert log.index("rewrite_data") < log.index("make_room") < log.index("write_answer")
    assert log.index("clear") < log.index("make_room")
    assert ROOM["old"] == {"rows": 30, "cols": 6} and set(ROOM["new"]) == {"rows", "cols"}
    assert ROOM["new"]["rows"] > 5 and ROOM["new"]["cols"] >= 3
    assert "write_data" not in log and "add_sheet" not in log                   # the Data sheet is never rebuilt
    assert not any("Data" in x for x in log if x.startswith(("backup", "drop", "restore")))
    assert log[-2:] == ["drop ~Q2 Answer", "save"]
    assert out["q"] == 2 and out["followup"] is True


def test_redraw_failure_rolls_back_restores_answer_and_rewrites_old_data(monkeypatch):
    log = []
    wire(monkeypatch, log, fail_at="write_answer")
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    with pytest.raises(RuntimeError, match="boom"):
        pipeline.redraw_job(None, pre, q=2, owned=None, target="session")
    assert [x for x in log if x != "sheet"][-3:] == ["rollback", "restore Q2 Answer", "rewrite_old True"]


def test_a_failing_restore_step_neither_hides_the_error_nor_skips_the_others(monkeypatch):
    log = []
    wire(monkeypatch, log, fail_at="write_answer", restore_fail=("restore Q2 Answer",))
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    with pytest.raises(RuntimeError) as ei:
        pipeline.redraw_job(None, pre, q=2, owned=None, target="session")
    msg = str(ei.value)
    assert "boom" in msg and "restore Q2 Answer failed" in msg
    assert [x for x in log if x != "sheet"][-3:] == ["rollback", "restore Q2 Answer", "rewrite_old True"]


def test_restore_when_the_original_sheet_is_gone(monkeypatch):
    from smartview_mcp.core import owned
    from smartview_mcp.excel import jobs, sheets
    monkeypatch.setattr(jobs, "com", lambda call, *a, **k: call())
    calls = []

    class WS:
        Visible = 0
        Name = "~Q2 Data"
    b = WS()

    class WSs:
        def __call__(self, n):
            return b

    class WB:
        Worksheets = WSs()

    class Tg:
        wb = WB()
    monkeypatch.setattr(sheets, "_sheet_exists", lambda t, n: n != "Q2 Data")
    monkeypatch.setattr(owned, "restore", lambda *a, **k: calls.append("v1 restore"))
    pipeline._restore(Tg(), "Q2 Data", "~Q2 Data")
    assert calls == [] and b.Name == "Q2 Data" and b.Visible == -1
    monkeypatch.setattr(sheets, "_sheet_exists", lambda t, n: True)
    pipeline._restore(Tg(), "Q2 Data", "~Q2 Data")
    assert calls == ["v1 restore"]
