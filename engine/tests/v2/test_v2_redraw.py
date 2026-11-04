# engine/tests/v2/test_v2_redraw.py
import pytest
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import pipeline


class T:
    path = "C:/a/v2/x.xlsx"
    created = True


def wire(monkeypatch, log, fail_at=None):
    def rec(name, ret=None):
        def f(*a, **k):
            log.append(name)
            if name == fail_at:
                raise RuntimeError("boom")
            return ret
        return f
    monkeypatch.setattr(pipeline, "_target", rec("target", T()))
    monkeypatch.setattr(pipeline, "_backup", lambda t, name: log.append(f"backup {name}") or f"~{name}")
    monkeypatch.setattr(pipeline, "_sheet", rec("sheet", object()))
    monkeypatch.setattr(pipeline, "_clear", rec("clear"))
    monkeypatch.setattr(pipeline, "_delete_sheet", rec("delete_data"))
    monkeypatch.setattr(pipeline, "_add_sheet", rec("add_sheet", object()))
    monkeypatch.setattr(pipeline, "write_data", rec("write_data"))
    monkeypatch.setattr(pipeline, "write_answer", rec("write_answer", []))
    monkeypatch.setattr(pipeline, "write_charts", rec("write_charts", []))
    monkeypatch.setattr(pipeline, "_calculate", rec("calculate"))
    monkeypatch.setattr(pipeline, "read_block", rec("read_block", ()))
    monkeypatch.setattr(pipeline, "compare", rec("compare", []))
    monkeypatch.setattr(pipeline, "_save_spec", rec("save_spec"))
    monkeypatch.setattr(pipeline, "_arrange", rec("arrange", []))
    monkeypatch.setattr(pipeline, "_drop_backup", lambda t, b: log.append(f"drop {b}"))
    monkeypatch.setattr(pipeline, "_restore", lambda t, name, b: log.append(f"restore {name}"))
    monkeypatch.setattr(pipeline, "_rollback", rec("rollback"))
    monkeypatch.setattr(pipeline, "_save", rec("save"))


def test_redraw_order(monkeypatch):
    log = []
    wire(monkeypatch, log)
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    out = pipeline.redraw_job(None, pre, q=2, owned={"rows": 30, "cols": 6}, target="session")
    assert log[:4] == ["target", "backup Q2 Answer", "backup Q2 Data", "sheet"]
    assert log.index("clear") < log.index("delete_data") < log.index("add_sheet") < log.index("write_data")
    assert log[-3:] == ["drop ~Q2 Answer", "drop ~Q2 Data", "save"]
    assert out["q"] == 2 and out["followup"] is True


def test_redraw_restores_both_sheets_on_failure(monkeypatch):
    log = []
    wire(monkeypatch, log, fail_at="write_answer")
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    with pytest.raises(RuntimeError):
        pipeline.redraw_job(None, pre, q=2, owned=None, target="session")
    assert log[-3:] == ["rollback", "restore Q2 Answer", "restore Q2 Data"]


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
