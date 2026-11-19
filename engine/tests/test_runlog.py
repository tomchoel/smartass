"""The run log (core/runlog.py): one JSON line per fin_analyze / fin_drill / fin_refresh call, with its spans."""
import asyncio
import json

from smartass.core import config, runlog


def test_a_run_is_one_json_line_with_its_spans(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "logs" / "runs.jsonl")
    runlog.begin("fin_analyze")
    with runlog.span("prepare"):
        with runlog.span("prepare.fetch"):
            pass
    with runlog.span("prepare"):
        pass
    runlog.add({"excel job.answer": 1.25, "com calls.excel job": 40})
    runlog.note(q=3, workbook="C:\\a\\b.xlsx", question="P&L Wholesale")
    line = runlog.finish(ok=True)
    stored = [json.loads(x) for x in (tmp_path / "logs" / "runs.jsonl").read_text(encoding="utf-8").splitlines()]
    assert stored == [line] and line["tool"] == "fin_analyze" and line["ok"] is True and line["q"] == 3
    assert set(line["timings"]) == {"prepare", "prepare.fetch", "excel job.answer", "com calls.excel job", "_total"}
    assert line["timings"]["excel job.answer"] == 1.25 and line["timings"]["_total"] >= line["timings"]["prepare"]


def test_outside_a_run_nothing_is_recorded_and_nothing_raises(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    with runlog.span("x"):
        pass
    runlog.add({"y": 1.0})
    runlog.note(q=1)
    assert not (tmp_path / "runs.jsonl").exists()


def test_a_failed_run_keeps_its_timings_and_a_redacted_error(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    runlog.begin("fin_drill")
    line = runlog.finish(ok=False, error="boom https://x/y?sso_token=abcdefghijklmnopqrstuvwxyz")  # allow-secret
    assert line["ok"] is False and "abcdefghij" not in line["error"] and "_total" in line["timings"]


def test_a_log_that_cannot_be_written_never_breaks_the_tool(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path)          # a directory: open() fails
    runlog.begin("fin_analyze")
    assert runlog.finish(ok=True)["ok"] is True


def test_spans_follow_the_call_into_a_worker_thread(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")

    async def tool():
        runlog.begin("fin_analyze")

        def work():
            with runlog.span("prepare.fetch"):
                pass
        await asyncio.to_thread(work)
        return runlog.finish(ok=True)
    assert "prepare.fetch" in asyncio.run(tool())["timings"]


def test_guard_logs_a_run_for_fin_analyze_and_not_for_fin_lookup(monkeypatch, tmp_path):
    from smartass.core import serverkit
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")

    async def fin_analyze():
        with runlog.span("prepare"):
            return {"q": 1}

    async def fin_lookup():
        return {}

    async def fin_drill():
        raise RuntimeError("no")
    asyncio.run(serverkit.guard(fin_analyze)())
    asyncio.run(serverkit.guard(fin_lookup)())
    try:
        asyncio.run(serverkit.guard(fin_drill)())
    except Exception:
        pass
    lines = [json.loads(x) for x in (tmp_path / "runs.jsonl").read_text(encoding="utf-8").splitlines()]
    assert [(x["tool"], x["ok"]) for x in lines] == [("fin_analyze", True), ("fin_drill", False)]
    assert "prepare" in lines[0]["timings"]


def test_a_drawn_answer_logs_the_excel_job_stages_and_its_com_call_total(monkeypatch, tmp_path):
    from types import SimpleNamespace

    from smartass import server
    from smartass.core import serverkit
    from smartass.core.timing import count_com, stage
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")

    async def drill(cell):
        return SimpleNamespace(rows=[{"AMOUNT": -10.0, "AP_CUSTOMER_NAME": "A"}],
                               columns=["AP_CUSTOMER_NAME", "AMOUNT"], cell_value=-10.0,
                               reconciliation={"gap": 0.0}, warnings=[])

    async def user(fn, what):                        # the host would run fn on its COM thread
        return fn(None)

    def job(h, pre, target="session"):
        with stage("draw"):
            count_com()
            with stage("answer"):
                count_com()
        with stage("save"):
            count_com()
        return {"q": 7, "workbook": "W.xlsx"}
    monkeypatch.setattr(server, "_drill", drill)
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "_provider", lambda: "provider")
    monkeypatch.setattr(server, "run_prepare", lambda raw, provider, lines=None, **kw: "PRE")
    monkeypatch.setattr(server, "excel_job", job)
    cell = {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"}
    asyncio.run(serverkit.guard(server.fin_drill)(cell=cell))
    line = json.loads((tmp_path / "runs.jsonl").read_text(encoding="utf-8"))
    assert (line["tool"], line["q"], line["workbook"]) == ("fin_drill", 7, "W.xlsx")
    assert line["question"].startswith("Ledger lines behind")
    assert {"prepare", "prepare.drill", "prepare.fetch", "excel job.draw", "excel job.draw.answer",
            "excel job.save"} <= set(line["timings"])
    assert line["timings"]["com calls.excel job"] == 3 and line["timings"]["excel job.com calls.draw"] == 2
