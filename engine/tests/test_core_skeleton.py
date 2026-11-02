"""The new core's walking skeleton (docs/plans/2026-11-02-walking-skeleton.md)."""
import datetime as dt
import json
from types import SimpleNamespace

import pytest
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import config, svhttp, tools_core, userexcel
from smartview_mcp.core import check, contract, pipeline, progress, runlog, source
from smartview_mcp.grid import Axis, Layout
from smartview_mcp.host import HostError, InputError


def test_fte_rows_are_the_sheet_rows_of_man_year_lines():
    lay = Layout([], [Axis("All Accounts", ["ac.C30", "ac.STA20Y"])], [Axis("All Time Periods", ["Sep 26"])])
    parsed = SimpleNamespace(row_labels=[["C30 REVENUES (ac)"], ["STA20Y Man-years (ac)"]])
    assert userexcel.fte_rows(lay, parsed) == [4]

PIN = contract.Pull.model_validate({"id": "m", "connection": "FinMax",
                                    "rows": [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
                                    "columns": [{"dimension": "All Time Periods", "members": ["Sep 26"]}]})


def test_cells_keep_text_as_text_and_blanks_empty():
    assert source.cells([[None, "Jan 26"], ["EBITDA (ac)", 1.5, "x"]]) == (("", "'Jan 26", ""), ("'EBITDA (ac)", 1.5, "'x"))


def test_cells_write_missing_data_as_zero_like_smart_views_numericzero():
    # a #Missing text cell would make var/sum formulas on that line #VALUE!; labels and #NoAccess stay text
    vals = [[None, "Sep 26", "Sep 25"], ["EBITDA (ac)", "#Missing", 2.0], ["#Missing", "#NoAccess", None]]
    assert source.cells(vals, data_from=(1, 1)) == (("", "'Sep 26", "'Sep 25"), ("'EBITDA (ac)", 0, 2.0),
                                                     ("'#Missing", "'#NoAccess", 0))


def test_provenance_names_cube_pov_time_and_route():
    lay = Layout([("All Entities", "un.SY.2BUS"), ("All Scenarios", "All Scenarios")], [], [])
    p = dict(source.provenance("FinMax", lay, "http", dt.datetime(2026, 11, 2, 10, 5)))
    assert p["Source"] == "Oracle EPM Cloud · FinMax" and p["POV"] == "un.SY.2BUS"
    assert p["Fetched"] == "2026-11-02 10:05" and "re-fetch" in p["Route"]


def test_mismatches_compare_only_the_fetched_numbers():
    assert source.mismatches([[None, "a"], ["x", 1.0]], ((None, "a"), ("x", 1.0, "panel"))) == []
    assert source.mismatches([["x", 1.0]], (("x", 2.0),)) == ["row 1 column 2: fetched 1.0, sheet shows 2.0"]


def _adhoc(monkeypatch, seen):
    pull = SimpleNamespace(sheet="Q1 FinMax")
    monkeypatch.setattr(source.userexcel, "pull", lambda h, t, q, conn, pov, rows, cols, title: seen.append(title) or pull)
    return pull


def test_an_http_failure_falls_back_to_the_adhoc_grid_with_a_note(monkeypatch):
    def fail(*a):
        raise svhttp.SvHttpError("token expired")
    monkeypatch.setattr(source.svhttp, "shared", lambda: SimpleNamespace(fetch=fail))
    seen = []
    pull = _adhoc(monkeypatch, seen)
    s = source.pull(None, None, 1, 0, PIN)
    assert s.pull is pull and s.route == "adhoc" and "token expired" in s.note and seen == ["Q1 FinMax"]


def test_an_input_error_is_not_hidden_by_the_fallback(monkeypatch):
    def bad(*a):
        raise InputError("Invalid query: unknown dimension")
    monkeypatch.setattr(source.svhttp, "shared", lambda: SimpleNamespace(fetch=bad))
    with pytest.raises(InputError):
        source.pull(None, None, 1, 0, PIN)


def test_http_success_writes_the_sheet_and_names_later_pulls_by_id(monkeypatch):
    f = svhttp.Fetched(Layout([], [], []), [], 0.1)
    monkeypatch.setattr(source.svhttp, "shared", lambda: SimpleNamespace(fetch=lambda *a: f))
    monkeypatch.setattr(source, "write", lambda t, q, title, pin, fetched: (title, fetched))
    assert source.pull(None, None, 2, 1, PIN) == ("Q2 FinMax m", f)


SPEC = {"version": 1, "question": "EBITDA Sep 26 for Business", "gate": {"outcome": "clear"},
        "assumptions": ["Actual, NOK"],
        "pulls": [{"id": "m", "connection": "FinMax",
                   "rows": [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
                   "columns": [{"dimension": "All Time Periods", "members": ["Sep 26"]}]}],
        "layout": {"title": "Business — EBITDA", "columns": [{"header": "Sep 26"}],
                   "sections": [{"lines": [{"key": "e", "label": "EBITDA",
                                            "cells": [{"pull": "m", "row": ["ac.EBITDA"], "col": ["Sep 26"]}]}]}]}}


def test_skeleton_shows_title_question_status_and_assumptions():
    rows = progress.skeleton(contract.validate(SPEC), "fetching 1 source grid ...")
    assert rows[0] == "Business — EBITDA" and rows[1] == "Question: EBITDA Sep 26 for Business"
    assert rows[progress.STATUS_ROW - 1] == "Status: fetching 1 source grid ..."
    assert rows[-2:] == ["Assumptions", "•  Actual, NOK"]


def test_skeleton_of_a_proxy_answer_says_what_is_answered_instead():
    spec = contract.validate({**SPEC, "gate": {"outcome": "infeasible", "reason": "no product dimension",
                                               "proxy": "contribution margin per segment"}})
    assert progress.skeleton(spec, "x")[-1] == "Answered instead: contribution margin per segment"


def _src(mismatches=(), note=None, pov=(), unknown=()):
    parsed = SimpleNamespace(pov_problems=list(pov), unknown_members=list(unknown))
    return SimpleNamespace(pull=SimpleNamespace(sheet="Q1 FinMax", parsed=parsed), mismatches=list(mismatches),
                           note=note, route="http")


def test_all_green_is_ok_without_findings():
    assert check.review({"checks": {"Sum = YTD": "OK"}, "warnings": [], "lint": []}, [_src()]) == \
        {"ok": True, "findings": []}


def test_a_failed_check_or_a_sheet_that_differs_from_the_fetch_is_critical():
    r = check.review({"checks": {"Sum = YTD": "DIFF"}},
                     [_src(mismatches=["row 3 column 2: fetched 1.0, sheet shows 2.0"])])
    assert not r["ok"] and [f["severity"] for f in r["findings"]] == ["Critical", "Critical"]


def test_a_check_marked_warning_that_differs_is_a_warning_not_critical():
    r = check.review({"checks": {"EBITDA FinPlan = FinMax — Actual": "DIFF", "Sum = YTD": "OK"}}, [],
                     warn_checks={"EBITDA FinPlan = FinMax"})
    assert r == {"ok": True, "findings": [{"severity": "Warning",
                                           "text": "check 'EBITDA FinPlan = FinMax — Actual' is DIFF"}]}
    spec = contract.validate({**SPEC, "layout": {**SPEC["layout"], "checks": [
        {"label": "x", "expr": "e", "severity": "Warning"}]}})
    assert spec.layout.checks[0].severity == "Warning"


def test_fallback_rejected_members_and_missing_checks_are_warnings():
    r = check.review({"checks": {}}, [_src(note="HTTP route failed (x)", pov=["Smart View rejected y"])])
    assert r["ok"] and [f["severity"] for f in r["findings"]] == ["Warning"] * 3


def test_formula_errors_are_critical_and_style_lint_is_info():
    r = check.review({"checks": {"c": "OK"}, "formula_errors": ["ytd"], "lint": ["narrow column"]}, [])
    assert [(f["severity"], f["text"]) for f in r["findings"]] == [
        ("Critical", "formula error on line 'ytd'"), ("Info", "narrow column")]


def test_entry_records_gate_pulls_routes_and_result(monkeypatch, tmp_path):
    spec = contract.validate(SPEC)
    src = SimpleNamespace(pull=SimpleNamespace(sheet="Q1 FinMax", seconds=0.234), route="http")
    e = runlog.entry(spec, 1, [src], {"ok": True, "findings": []}, {"_total": 3.1}, workbook="C:/x/a.xlsx")
    assert e["pulls"] == [{"id": "m", "cube": "FinMax", "sheet": "Q1 FinMax", "route": "http", "seconds": 0.23}]
    assert (e["q"], e["gate"], e["ok"], e["error"]) == ("Q1", "clear", True, None)
    row = runlog.sheet_row(e)
    assert row[1:6] == ("Q1", "EBITDA Sep 26 for Business", "clear", "m=FinMax (http, 0.23 s)", "OK")
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "logs" / "runs.jsonl")
    runlog.append(e)
    assert json.loads((tmp_path / "logs" / "runs.jsonl").read_text(encoding="utf-8"))["q"] == "Q1"


def test_assumptions_and_the_question_lead_the_notes_and_the_raw_layout_is_untouched():
    raw = {**SPEC, "layout": {**SPEC["layout"], "notes": ["n"]}}
    out = pipeline.answer_layout(raw["layout"], contract.validate(raw))
    assert out["notes"] == ["Question: EBITDA Sep 26 for Business", "Assumption: Actual, NOK", "n"]
    assert raw["layout"]["notes"] == ["n"]


def _fake_run(monkeypatch, calls, pull_fails=False):
    t = SimpleNamespace(path="C:/x/a.xlsx", status=lambda text: calls.append(("bar", text)))
    pull = SimpleNamespace(sheet="Q3 FinMax", seconds=0.2, parsed=SimpleNamespace(pov_problems=[], unknown_members=[]))
    monkeypatch.setattr(pipeline.userexcel, "resolve_target", lambda h, target, allow, title: t)
    monkeypatch.setattr(pipeline.userexcel, "next_question", lambda t: 3)
    monkeypatch.setattr(pipeline.userexcel, "rollback", lambda t: ["Q3 Answer"])
    monkeypatch.setattr(pipeline.userexcel, "save_spec",
                        lambda t, q, raw, ids, owned=None: calls.append(("save_spec", q, owned)))
    monkeypatch.setattr(pipeline.progress, "show", lambda t, q, spec, status: calls.append(("show", q)))
    monkeypatch.setattr(pipeline.progress, "set_status", lambda t, q, text: calls.append(("status", text)))

    def pull_(h, t, q, i, pin):
        calls.append(("pull", pin.id))
        if pull_fails:
            raise HostError("Retrieve failed")
        return source.Source(pull, "http", [])
    monkeypatch.setattr(pipeline.source, "pull", pull_)
    monkeypatch.setattr(pipeline.specs, "prepare", lambda layout, ids, sess: calls.append(("prepare", ids)) or layout)
    monkeypatch.setattr(pipeline.answer_mod, "build", lambda t, q, spec, sess, suffix, owned=None, **kw:
                        calls.append(("build", suffix, kw.get("reserve"))) or
                        {"sheet": "Q3 Answer", "checks": {"c": "OK"}, "extent": {"rows": 20, "cols": 6}})
    t.wb = SimpleNamespace(Worksheets=lambda name: name)
    monkeypatch.setattr(pipeline.userexcel, "com", lambda call, *a, **k: call())
    monkeypatch.setattr(pipeline.findings, "render",
                        lambda ws, row, fs: calls.append(("findings", ws, row)) or row + len(fs) + 1)


def test_run_shows_the_answer_sheet_first_then_pulls_then_renders(monkeypatch):
    calls = []
    _fake_run(monkeypatch, calls)
    r = pipeline.run(None, SPEC)
    assert [c[0] for c in calls] == ["show", "status", "pull", "status", "prepare", "build", "findings", "save_spec"]
    assert ("prepare", {"m": "Q3 FinMax"}) in calls and ("build", "Answer", pipeline.findings.RESERVE) in calls
    assert ("findings", "Q3 Answer", 22) in calls
    assert ("save_spec", 3, {"rows": 23, "cols": 6}) in calls and r.raw is SPEC and r.owned == {"rows": 23, "cols": 6}
    assert r.q == 3 and r.review == {"ok": True, "findings": []}


def test_a_failed_findings_table_is_a_warning_not_a_failed_answer(monkeypatch):
    calls = []
    _fake_run(monkeypatch, calls)
    monkeypatch.setattr(pipeline.findings, "render", lambda ws, row, fs: (_ for _ in ()).throw(RuntimeError("busy")))
    r = pipeline.run(None, SPEC)
    assert r.owned == {"rows": 20, "cols": 6} and any("findings" in w for w in r.res["warnings"])


def test_a_failed_pull_removes_the_partial_sheets(monkeypatch):
    calls = []
    _fake_run(monkeypatch, calls, pull_fails=True)
    with pytest.raises(HostError, match="removed the partial sheets: Q3 Answer"):
        pipeline.run(None, SPEC)


@pytest.mark.anyio
async def test_an_ambiguous_question_returns_its_questions_and_touches_no_excel(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")

    async def no_excel(*a, **k):
        raise AssertionError("Excel was touched")
    monkeypatch.setattr(tools_core, "_call_user", no_excel)
    spec = {"version": 1, "question": "How are we doing on costs?",
            "gate": {"outcome": "ambiguous", "questions": ["Which unit?", "Which period?"]}}
    out = await tools_core.sv_answer(spec)
    assert out["gate"] == "ambiguous" and out["questions"] == ["Which unit?", "Which period?"]
    assert json.loads((tmp_path / "runs.jsonl").read_text(encoding="utf-8"))["gate"] == "ambiguous"


@pytest.mark.anyio
async def test_an_invalid_spec_is_a_readable_error():
    with pytest.raises(ToolError, match="not a valid answer spec v1"):
        await tools_core.sv_answer({"question": "x"})



def test_the_pipeline_turns_expectations_into_static_checks(monkeypatch):
    from smartview_mcp import fin_spec
    lay = {"title": "t", "columns": [{"header": "a"}], "sections": [], "checks": [],
           "expect": {"source": "Warehouse", "cells": {"pnl": [[["ac.C30"], ["Sep 26"], 5.0, "f1", 0.5]]}}}
    monkeypatch.setattr(fin_spec, "reconcile_pull", lambda p, cells, source: {"label": f"{source} = {p.sheet}",
                                                                             "static": "OK", "detail": "1 cells"})
    out = pipeline.expect_checks(lay, {"pnl": SimpleNamespace(sheet="Q1 FinMax pnl")})
    assert "expect" not in out and out["checks"][-1] == {"label": "Warehouse = Q1 FinMax pnl", "static": "OK",
                                                          "detail": "1 cells"}


def test_reply_helpers_live_in_core_reply():
    from smartview_mcp.core import reply
    assert callable(reply.finish) and callable(reply.result)
