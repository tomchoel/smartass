"""The new core's walking skeleton (docs/plans/2026-11-02-walking-skeleton.md)."""
import datetime as dt
import json
from types import SimpleNamespace

import pytest
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import config, svhttp, tools_core
from smartview_mcp.excel import adhoc, jobs, sheets, target
from smartview_mcp.core import check, contract, pipeline, progress, runlog, source
from smartview_mcp.grid import Axis, Layout
from smartview_mcp.host import HostError, InputError


def test_fte_rows_are_the_sheet_rows_of_man_year_lines():
    lay = Layout([], [Axis("All Accounts", ["ac.C30", "ac.STA20Y"])], [Axis("All Time Periods", ["Sep 26"])])
    parsed = SimpleNamespace(row_labels=[["C30 REVENUES (ac)"], ["STA20Y Man-years (ac)"]])
    assert sheets.fte_rows(lay, parsed) == [4]

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
    monkeypatch.setattr(adhoc, "pull", lambda h, t, q, conn, pov, rows, cols, title: seen.append(title) or pull)
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
    assert rows[0] == "Business — EBITDA" and rows[1] == "Question (as asked): EBITDA Sep 26 for Business"
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
    assert out["notes"] == ["Question (as asked): EBITDA Sep 26 for Business", "Assumption: Actual, NOK", "n"]
    assert raw["layout"]["notes"] == ["n"]


def _fake_run(monkeypatch, calls, pull_fails=False):
    t = SimpleNamespace(path="C:/x/a.xlsx", status=lambda text: calls.append(("bar", text)))
    pull = SimpleNamespace(sheet="Q3 FinMax", seconds=0.2, parsed=SimpleNamespace(pov_problems=[], unknown_members=[]))
    monkeypatch.setattr(target, "resolve_target", lambda h, target, allow, title: t)
    monkeypatch.setattr(sheets, "next_question", lambda t: 3)
    monkeypatch.setattr(sheets, "rollback", lambda t: ["Q3 Answer"])
    monkeypatch.setattr(sheets, "save_spec",
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
    monkeypatch.setattr(jobs, "com", lambda call, *a, **k: call())
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
           "expect": {"source": "Live read", "cells": {"pnl": [[["ac.C30"], ["Sep 26"], 5.0, "f1", 0.5]]}}}
    monkeypatch.setattr(fin_spec, "reconcile_pull", lambda p, cells, source: {"label": f"{source} = {p.sheet}",
                                                                             "static": "OK", "detail": "1 cells"})
    out = pipeline.expect_checks(lay, {"pnl": SimpleNamespace(sheet="Q1 FinMax pnl")})
    assert "expect" not in out and out["checks"][-1] == {"label": "Live read = Q1 FinMax pnl", "static": "OK",
                                                          "detail": "1 cells"}


def test_reply_helpers_live_in_core_reply():
    from smartview_mcp.core import reply
    assert callable(reply.finish) and callable(reply.result)


# ------------------------------------------------------------------ A4: sv_answer enforces the gate
COSTS = "How are we doing on costs?"
SUPPLIERS = "Which suppliers invoiced most on travel in 2026?"


def _answer_env(monkeypatch, tmp_path, calls):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    monkeypatch.setattr(tools_core, "_closed_cached", lambda: "Sep 26")
    monkeypatch.setattr(tools_core.scope.prefs, "lookup", lambda item, phrase: None)

    async def call_user(job, what):
        calls.append(what)
        return {"sheet": "Q1 Answer"}
    monkeypatch.setattr(tools_core, "_call_user", call_user)


def _recipe_spec(question, name, gate=None, **extra):
    return {"version": 1, "question": question, "gate": gate or {"outcome": "clear"},
            "recipe": {"name": name, "params": {}}, **extra}


@pytest.mark.anyio
async def test_answer_refuses_to_pull_when_the_gate_is_still_ambiguous(monkeypatch, tmp_path):
    calls = []
    _answer_env(monkeypatch, tmp_path, calls)
    out = await tools_core.answer(_recipe_spec(COSTS, "pnl"))
    assert out["gate"] == "ambiguous" and "measure" in out["questions"][0].lower() and calls == []
    assert "answers=" in out["note"]


@pytest.mark.anyio
async def test_answer_pulls_when_answers_resolve_the_gate(monkeypatch, tmp_path):
    calls = []
    _answer_env(monkeypatch, tmp_path, calls)
    out = await tools_core.answer(_recipe_spec(COSTS, "pnl", gate={"outcome": "clear",
                                                                   "answers": {"measure": "opex"}}))
    assert out == {"sheet": "Q1 Answer"} and calls == ["sv_answer"]


@pytest.mark.anyio
async def test_answer_refuses_a_departure_without_deviation(monkeypatch, tmp_path):
    calls = []
    _answer_env(monkeypatch, tmp_path, calls)
    with pytest.raises(InputError, match="departs from the standard analysis: recipe cost_base instead of "
                                         "lines_ranking.*deviation="):
        await tools_core.answer(_recipe_spec(SUPPLIERS, "cost_base"))
    with pytest.raises(InputError, match="a free-form spec where sv_gate named recipe pnl"):
        await tools_core.answer(SPEC)
    assert calls == []


@pytest.mark.anyio
async def test_deviation_is_written_on_the_sheet_and_logged(monkeypatch, tmp_path):
    calls = []
    _answer_env(monkeypatch, tmp_path, calls)
    why = "The user asked for cost centres, not suppliers"
    await tools_core.answer(_recipe_spec(SUPPLIERS, "cost_base", deviation=why))
    assert calls == ["sv_answer"]
    spec = contract.validate({**SPEC, "deviation": why})
    line = f"Deviation from the standard analysis (user's request): {why}"
    assert line in progress.skeleton(spec, "x")
    assert pipeline.answer_layout(SPEC["layout"], spec)["notes"][2] == line
    assert runlog.entry(spec, 1, [], None, {})["deviation"] == why
    with pytest.raises(ValueError):
        contract.validate({**SPEC, "deviation": "x" * 301})


@pytest.mark.anyio
async def test_norwegian_assumption_is_refused_with_its_path(monkeypatch, tmp_path):
    calls = []
    _answer_env(monkeypatch, tmp_path, calls)
    spec = _recipe_spec(SUPPLIERS, "lines_ranking",
                        assumptions=["NOK", "Konsulenter = kontoene C60080 og C60090 (valgt av meg)"])
    with pytest.raises(InputError, match=r"(?s)assumptions\[1\]: Konsulenter.*Write sheet text in English"):
        await tools_core.answer(spec)
    assert calls == []


def test_question_is_shown_as_asked_and_not_checked():
    spec = contract.validate({**SPEC, "question": "Hva er EBITDA for Business i september?"})
    assert progress.skeleton(spec, "x")[1] == "Question (as asked): Hva er EBITDA for Business i september?"
    assert pipeline.answer_layout(SPEC["layout"], spec)["notes"][0].startswith("Question (as asked): Hva")
    assert all(path != "question" for path, _ in contract.sheet_texts(spec.model_dump()))


# ------------------------------------------------------------------ A5: finance defaults, read-back, follow-ups
def _ranking_spec(top: int, other: bool, **extra) -> dict:
    cell = {"pull": "cb", "row": ["{m}"], "col": ["ac.C70"]}
    lines = [{"key": "tot", "label": "Total", "cells": [{"pull": "cb", "row": ["un.SY"], "col": ["ac.C70"]}]}]
    if other:
        lines.insert(0, {"key": "rest", "label": "Other cost centres", "cells": [{"expr": "tot-c0"}]})
    return {"version": 1, "question": "Which cost centres cost the most?", "gate": {"outcome": "clear"},
            "pulls": [{"id": "cb", "connection": "FinMax",
                       "rows": [{"dimension": "All Entities", "members": ["un.SY"]}],
                       "columns": [{"dimension": "All Accounts", "members": ["ac.C70"]}]}],
            "layout": {"title": "Top cost centres", "columns": [{"header": "Costs"}],
                       "sections": [{"title": "Top", "repeat": {"over": {"top": top, "pull": "cb", "by": ["ac.C70"]},
                                                               "line": {"key": "c{i}", "label": "{label}",
                                                                        "cells": [cell]}},
                                     "lines": lines}]}, **extra}


@pytest.mark.anyio
async def test_freeform_ranking_above_top10_without_deviation_is_refused(monkeypatch, tmp_path):
    calls = []
    _answer_env(monkeypatch, tmp_path, calls)
    monkeypatch.setattr(tools_core.scope, "gate", lambda q, closed, answers=None: {"suggested_outcome": "clear",
                                                                                   "recipe": None})
    with pytest.raises(InputError, match=r"departs from the standard analysis: a top-25 ranking \(the standard is "
                                         r"the top 10 \+ Other\); a top-25 ranking without an 'Other' line"):
        await tools_core.answer(_ranking_spec(25, other=False))
    assert calls == []
    await tools_core.answer(_ranking_spec(10, other=True))
    await tools_core.answer(_ranking_spec(25, other=True, deviation="The user asked for the top 25"))
    assert calls == ["sv_answer", "sv_answer"]


def test_reply_key_values_are_the_sheet_values_after_calc():
    from smartview_mcp.core import reply
    prepared = {"columns": [{"header": "Sep 26"}, {"header": "Sep 25"}],
                "sections": [{"title": "S", "lines": [{"key": "e", "label": "EBITDA"}, {"label": "no key"},
                                                      {"key": "r", "label": "Revenue"}]}]}
    res = {"values": {"e": [10.0, None], "r": [3.0, 4.0]}}
    assert reply.key_values(prepared, res) == {"EBITDA": {"Sep 26": 10.0}, "Revenue": {"Sep 26": 3.0, "Sep 25": 4.0}}
    assert list(reply.key_values(prepared, res, limit=1)) == ["EBITDA"]


def test_reply_has_at_most_three_followups():
    from smartview_mcp.core import recipes
    for name in ("pnl", "actual_vs_plan", "monthly_trend", "cost_base", "concept"):
        ups = recipes.followups(name, {"entity": "un.SY.2BUS", "area": "un.SY.2BUS"}, "Sep 26")
        assert 1 <= len(ups) <= 3 and all("{" not in u for u in ups), (name, ups)
    assert any("Sep 26" in u for u in recipes.followups("cost_base", {"parent": "un.SY"}, "Sep 26"))
    assert recipes.followups(None, {}, "Sep 26") == [] and recipes.followups("echo", {}, "Sep 26") == []


def test_reply_result_names_every_sheet_written_and_the_key_values():
    from smartview_mcp.core import reply
    t = SimpleNamespace(path="C:/x/a.xlsx", name="a.xlsx", created=True, wb=SimpleNamespace(ReadOnly=False))
    res = {"sheet": "Q1 Answer", "values": {"e": [1.0]}, "checks": [], "warnings": [], "lint": [],
           "formula_errors": []}
    spec = {"columns": [{"header": "Sep 26"}], "sections": [{"title": "S", "lines": [{"key": "e", "label": "E"}]}]}
    out = reply.result(t, 1, spec, res, [SimpleNamespace(sheet="Q1 FinMax m")], [])
    assert out["sheets"] == ["Q1 Answer", "Q1 FinMax m", "Sources"] and out["key_values"] == {"E": {"Sep 26": 1.0}}


# ------------------------------------------------------------------ B3: topic sheets
class _Cell:
    def __init__(self):
        self.Formula = None


class _Ws:
    def __init__(self, name, log):
        self.Name, self.log, self.cells = name, log, {}

    def Cells(self, r, c):
        return self.cells.setdefault((r, c), _Cell())

    def Move(self, before, after):
        self.log.append(("move", self.Name, after.Name))

    def Activate(self):
        self.log.append(("activate", self.Name))


def test_answer_links_to_topic_sheets_and_back(monkeypatch):
    log = []
    wss = {n: _Ws(n, log) for n in ("Q1 Answer", "Q1 By project")}
    t = SimpleNamespace(wb=SimpleNamespace(Worksheets=lambda n: wss[n]))
    monkeypatch.setattr(progress, "com", lambda call, *a, **k: call())
    progress.link(t, 1, ["By project"])
    assert wss["Q1 Answer"].cells[(3, 2)].Formula == '=HYPERLINK("#\'Q1 By project\'!A1","Q1 By project →")'
    assert wss["Q1 By project"].cells[(3, 2)].Formula == '=HYPERLINK("#\'Q1 Answer\'!A1","← Q1 Answer")'
    assert ("move", "Q1 By project", "Q1 Answer") in log and log[-1] == ("activate", "Q1 Answer")


def test_pipeline_renders_each_extra_on_its_topic_sheet_and_reviews_its_checks(monkeypatch):
    calls = []
    _fake_run(monkeypatch, calls)
    monkeypatch.setattr(pipeline.answer_mod, "build", lambda t, q, spec, sess, suffix, owned=None, **kw:
                        calls.append(("build", suffix, kw.get("reserve"))) or
                        {"sheet": f"Q3 {suffix}", "checks": {"c": "DIFF" if suffix != "Answer" else "OK"},
                         "extent": {"rows": 20, "cols": 6}, "values": {"x": [1.0]}})
    monkeypatch.setattr(pipeline.progress, "link", lambda t, q, topics: calls.append(("link", q, topics)))
    extra = {"topic": "By project", "layout": {"title": "P", "columns": [{"header": "a"}],
                                               "sections": [{"title": "S", "lines": []}]}}
    r = pipeline.run(None, {**SPEC, "extras": [extra]})
    assert ("build", "By project", None) in calls and ("link", 3, ["By project"]) in calls
    assert r.res["checks"]["By project: c"] == "DIFF" and r.res["extra_sheets"] == ["Q3 By project"]
    assert not r.review["ok"]


def test_topic_sheet_is_ours_for_rollback_and_next_question():
    names = ["Q1 Answer", "Q1 By project", "Q1 Lines d1", "Q7 Forecast", "Sources"]
    assert [sheets.question_of(n, names) for n in names] == [1, 1, 1, None, None]
    wb = SimpleNamespace(Worksheets=[SimpleNamespace(Name=n) for n in names + ["Q2 By project"]])
    assert sheets.next_question(SimpleNamespace(wb=wb)) == 2
    assert sheets.question_of("Q2 By project", names) is None                 # no Q2 Answer: the user's sheet
