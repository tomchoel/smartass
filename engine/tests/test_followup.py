"""Phase 5 follow-ups: stored spec + owned range, source plan, follow-up pipeline, sv_answer(followup=...)."""
import copy
import json
from pathlib import Path
from types import SimpleNamespace

import pytest
from mcp.server.mcpserver.exceptions import ToolError
from test_core_skeleton import _fake_run          # tests/ is on sys.path (rootdir import)

from smartview_mcp import tools_core
from smartview_mcp.excel import jobs, sheets, target
from smartview_mcp.core import contract, followup, pipeline

B02 = json.loads((Path(__file__).parent / "fixtures" / "b02_answer_spec.json").read_text(encoding="utf-8"))


class Meta:
    """A fake _svmeta sheet: rows of cell values (column A = key). Like Excel, End(xlUp) on an empty column gives
    row 1 (so the first row written is row 2) and UsedRange starts at the first used row."""
    def __init__(self):
        self.rows: dict[int, dict[int, str]] = {}

    @property
    def UsedRange(self):
        first, n = min(self.rows, default=1), max(self.rows, default=1)
        w = max((max(r) for r in self.rows.values() if r), default=1)
        return SimpleNamespace(Row=first, Value=tuple(tuple(self.rows.get(r, {}).get(c) for c in range(1, w + 1))
                                                      for r in range(first, n + 1)))

    def Range(self, a, b):
        meta = self
        self.range_writes = getattr(self, "range_writes", 0) + 1

        class R:
            def __setattr__(self, k, v):
                for i, val in enumerate(v[0], a.addr[1]):
                    meta.rows.setdefault(a.addr[0], {})[i] = val[1:] if val.startswith("'") else val
        return R()

    def Cells(self, r, c):
        meta = self

        class C:
            addr = (r, c)

            def __setattr__(self, k, v):
                meta.cell_writes = getattr(meta, "cell_writes", 0) + 1
                meta.rows.setdefault(r, {})[c] = v[1:] if isinstance(v, str) and v.startswith("'") else v

            @property
            def End(self):
                return lambda d: SimpleNamespace(Row=max(meta.rows, default=1))
        return C()

    def Rows(self, r):
        return SimpleNamespace(ClearContents=lambda: self.rows.pop(r, None))

    Rows.Count = 1048576


def _meta_target(monkeypatch):
    ws = Meta()
    t = SimpleNamespace(wb=SimpleNamespace(Worksheets=lambda name: ws))
    monkeypatch.setattr(jobs, "com", lambda call, **kw: call())
    monkeypatch.setattr(sheets, "_sheet_exists", lambda t, name: name == sheets.META_SHEET)
    return t, ws


def test_the_spec_is_stored_with_its_owned_range_and_read_back(monkeypatch):
    t, _ = _meta_target(monkeypatch)
    sheets.save_spec(t, 2, B02, {"m": "Q2 FinMax"}, owned={"rows": 40, "cols": 7})
    got = sheets.load_answer(t, 2)
    assert got == {"spec": B02, "ids": {"m": "Q2 FinMax"}, "owned": {"rows": 40, "cols": 7}}
    assert sheets.load_answer(t, 3) is None


def test_stored_questions_lists_only_v1_answer_specs(monkeypatch):
    t, _ = _meta_target(monkeypatch)
    sheets.save_spec(t, 3, B02, {})
    sheets.save_spec(t, 1, {"question": "an old pre-v1 spec"}, {})
    sheets.save_spec(t, 2, B02, {}, owned={"rows": 5, "cols": 3})
    sheets.save_spec(t, 3, B02, {"m": "Q3 FinMax"})                       # overwrites its row
    assert sheets.stored_questions(t) == [2, 3]


def test_saving_a_spec_again_overwrites_its_own_row_and_keeps_the_pull_rows_above(monkeypatch):
    t, ws = _meta_target(monkeypatch)
    ws.Cells(2, 1).Value, ws.Cells(2, 2).Value = "'Q1 FinMax ytd", "'{}"      # save_meta's first row is row 2
    sheets.save_spec(t, 1, B02, {"m": "Q1 FinMax"})
    sheets.save_spec(t, 1, {**B02, "question": "patched"}, {"m": "Q1 FinMax"})
    assert ws.rows[2] == {1: "Q1 FinMax ytd", 2: "{}"}
    assert [r[1] for r in ws.rows.values()].count("SPEC Q1") == 1
    assert sheets.load_answer(t, 1)["spec"]["question"] == "patched"


def _meta_pull(extra=None):
    ax = [SimpleNamespace(dimension="Account", members=["EBITDA"])]
    return SimpleNamespace(sheet="Q1 FinMax", connection="FinMax", parsed=SimpleNamespace(pov_problems=[]),
                           layout=SimpleNamespace(pov={}, rows=ax, cols=ax)), extra


def test_save_meta_is_one_array_write_unless_a_chunk_is_long(monkeypatch):
    t, ws = _meta_target(monkeypatch)
    p, _ = _meta_pull()
    sheets.save_meta(t, p)
    assert ws.range_writes == 1 and getattr(ws, "cell_writes", 0) == 0
    assert json.loads(ws.rows[2][2])["connection"] == "FinMax"
    sheets.save_meta(t, p, {"pad": "x" * 20000})                       # one 20 000-char chunk > 8000
    assert ws.range_writes == 1 and ws.cell_writes == 2                 # name + one chunk, cell by cell
    assert ws.rows[3][1] == "Q1 FinMax" and json.loads(ws.rows[3][2])["pad"] == "x" * 20000


def _pull(pid, members=("EBITDA",), conn="FinMax"):
    return {"id": pid, "connection": conn, "rows": [{"dimension": "Account", "members": list(members)}],
            "columns": [{"dimension": "Period", "members": ["Sep 26"]}]}


def test_unchanged_pulls_are_reused_changed_refetched_removed_dropped():
    old = {"pulls": [_pull("m"), _pull("ytd"), _pull("x")]}
    new = {"pulls": [_pull("m"), _pull("ytd", ("EBITDA", "C30")), _pull("fp", conn="FinPlan")]}
    ids = {"m": "Q2 FinMax", "ytd": "Q2 FinMax ytd", "x": "Q2 FinMax x"}
    plan = followup.plan_sources(old, ids, new, exists=lambda s: True)
    assert plan.reuse == {"m": "Q2 FinMax"}
    assert plan.fetch == ["ytd", "fp"]
    assert plan.replace == {"ytd": "Q2 FinMax ytd"}          # parked, refetched under the same name
    assert plan.drop == ["Q2 FinMax x"]


def test_a_reused_pull_whose_sheet_was_deleted_is_fetched_again():
    old = {"pulls": [_pull("m")]}
    plan = followup.plan_sources(old, {"m": "Q2 FinMax"}, old, exists=lambda s: False)
    assert plan.reuse == {} and plan.fetch == ["m"] and plan.replace == {}


def test_a_corrupted_stored_pull_counts_as_changed():
    old = {"pulls": [{**_pull("m"), "connection": "Nope"}]}
    plan = followup.plan_sources(old, {"m": "Q2 FinMax"}, {"pulls": [_pull("m")]}, exists=lambda s: True)
    assert plan.fetch == ["m"] and plan.replace == {"m": "Q2 FinMax"}


def test_expectations_of_a_changed_pull_are_dropped():
    lay = {"expect": {"source": "Live read", "cells": {"pnl": [[1, 1, 5.0, "f", 0.5]], "fte": [[1, 1, 2.0, "g", 0.5]]}}}
    assert followup.drop_stale_expect(lay, ["fte"])["expect"]["cells"] == {"pnl": [[1, 1, 5.0, "f", 0.5]]}
    assert "expect" not in followup.drop_stale_expect({"expect": {"source": "W", "cells": {"fte": []}}}, ["fte"])
    assert followup.drop_stale_expect({"title": "x"}, ["fte"]) == {"title": "x"}


def test_the_current_question_is_the_active_answer_sheet_else_the_latest_stored(monkeypatch):
    t = SimpleNamespace()
    monkeypatch.setattr(followup, "_active_sheet", lambda t: "Q3 Answer")
    monkeypatch.setattr(sheets, "stored_questions", lambda t: [1, 3, 5])
    assert followup.current_question(t) == 3
    monkeypatch.setattr(followup, "_active_sheet", lambda t: "Sources")
    monkeypatch.setattr(sheets, "stored_questions", lambda t: [1, 4])
    assert followup.current_question(t) == 4
    monkeypatch.setattr(sheets, "stored_questions", lambda t: [])
    assert followup.current_question(t) is None


def _old(raw):
    return {"spec": raw, "ids": {"m": "Q2 FinMax", "ytd": "Q2 FinMax ytd"}, "owned": {"rows": 30, "cols": 7}}


def _stored(*sheets):
    parsed = SimpleNamespace(pov_problems=[], unknown_members=[])
    return {s: SimpleNamespace(sheet=s, seconds=0.0, parsed=parsed) for s in sheets}


def _follow_env(monkeypatch, calls, stored):
    _fake_run(monkeypatch, calls)
    monkeypatch.setattr(sheets, "load_pulls", lambda t, q: stored)
    monkeypatch.setattr(pipeline, "_exists", lambda t, s: True)
    monkeypatch.setattr(pipeline.owned_mod, "backup", lambda t, name: calls.append(("backup", name)) or "~" + name)
    monkeypatch.setattr(pipeline, "_delete", lambda t, names: calls.append(("delete", tuple(names))))
    monkeypatch.setattr(sheets, "park_sheets", lambda t, names: {"~" + n: n for n in names})
    monkeypatch.setattr(sheets, "drop_parked", lambda t, p: calls.append(("drop_parked", tuple(p.values()))))


def test_a_follow_up_reuses_unchanged_pulls_and_redraws_the_same_answer(monkeypatch):
    calls = []
    _follow_env(monkeypatch, calls, _stored("Q2 FinMax", "Q2 FinMax ytd"))
    raw = copy.deepcopy(B02)
    raw["layout"]["title"] = "Changed"
    r = pipeline.run(None, raw, "session", False, q=2, old=_old(B02))
    names = [c[0] for c in calls]
    assert "pull" not in names and "show" not in names and "status" not in names   # no re-fetch, no skeleton
    assert ("backup", "Q2 Answer") in calls and ("delete", ("~Q2 Answer",)) in calls
    assert r.q == 2 and sorted(r.reused) == ["m", "ytd"] and r.owned == {"rows": 23, "cols": 6}   # + findings
    assert [s.route for s in r.sources] == ["reused", "reused"] and r.raw["layout"]["title"] == "Changed"
    assert ("save_spec", 2, {"rows": 23, "cols": 6}) in calls


def test_a_follow_up_that_adds_a_pull_fetches_only_that_pull(monkeypatch):
    calls = []
    _follow_env(monkeypatch, calls, _stored("Q2 FinMax", "Q2 FinMax ytd"))
    fp = {"id": "fp", "connection": "FinPlan", "rows": [{"dimension": "Account", "members": ["EBITDA"]}],
          "columns": [{"dimension": "Scenario", "members": ["Actual"]}]}
    raw = contract.patched(B02, {"pulls": [fp]})
    r = pipeline.run(None, raw, "session", False, q=2, old=_old(B02))
    pulled = [c[1] for c in calls if c[0] == "pull"]
    assert pulled == ["fp"] and [s.route for s in r.sources] == ["reused", "reused", "http"]


def test_a_changed_pull_is_parked_refetched_and_its_old_sheet_dropped(monkeypatch):
    calls = []
    _follow_env(monkeypatch, calls, _stored("Q2 FinMax", "Q2 FinMax ytd"))
    raw = contract.patched(B02, {"pulls": [{"id": "ytd", "title": "changed"}]})
    pipeline.run(None, raw, "session", False, q=2, old=_old(B02))
    assert [c[1] for c in calls if c[0] == "pull"] == ["ytd"]
    assert ("drop_parked", ("Q2 FinMax ytd",)) in calls


def test_a_reused_sheet_without_its_meta_row_is_fetched_again(monkeypatch):
    calls = []
    _follow_env(monkeypatch, calls, _stored("Q2 FinMax"))
    r = pipeline.run(None, copy.deepcopy(B02), "session", False, q=2, old=_old(B02))
    assert [c[1] for c in calls if c[0] == "pull"] == ["ytd"] and r.reused == ["m"]


def test_a_failed_redraw_restores_the_answer_and_the_parked_sources(monkeypatch):
    calls = []
    _follow_env(monkeypatch, calls, {})
    monkeypatch.setattr(pipeline.owned_mod, "restore", lambda t, n, b: calls.append(("restore", n, b)))
    monkeypatch.setattr(sheets, "restore_parked", lambda t, p: calls.append(("unpark", tuple(p.values()))))
    monkeypatch.setattr(pipeline.answer_mod, "build", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("boom")))
    raw = contract.patched(B02, {"pulls": [{"id": "ytd", "title": "changed"}]})
    with pytest.raises(Exception, match="follow-up was not applied"):
        pipeline.run(None, raw, "session", False, q=2, old=_old(B02))
    assert ("restore", "Q2 Answer", "~Q2 Answer") in calls
    assert sorted(next(c[1] for c in calls if c[0] == "unpark")) == ["Q2 FinMax", "Q2 FinMax ytd"]
    assert not any(c[0] in ("delete", "drop_parked") for c in calls)


# ------------------------------------------------------------------ sv_answer(followup=...)
def test_a_follow_up_patches_the_current_answer(monkeypatch):
    monkeypatch.setattr(tools_core.followup_mod, "current_question", lambda t: 2)
    monkeypatch.setattr(sheets, "load_answer", lambda t, q: {"spec": B02, "ids": {}, "owned": None})
    q, old, raw = tools_core._followup_spec(SimpleNamespace(), {"layout": {"title": "New title",
                                                                          "notes": {"append": ["n"]}}}, None)
    assert q == 2 and old["spec"] is B02 and raw["layout"]["title"] == "New title"
    assert raw["layout"]["notes"][-1] == "n"


def test_a_follow_up_without_a_stored_answer_says_ask_a_new_question(monkeypatch):
    monkeypatch.setattr(tools_core.followup_mod, "current_question", lambda t: None)
    with pytest.raises(tools_core.InputError, match="no answer to follow up"):
        tools_core._followup_spec(SimpleNamespace(), {"layout": {"title": "x"}}, None)


def test_a_follow_up_on_an_old_answer_is_refused(monkeypatch):
    monkeypatch.setattr(sheets, "load_answer",
                        lambda t, q: {"spec": {"question": "old pre-v1 spec"}, "ids": {}, "owned": None})
    with pytest.raises(tools_core.InputError, match="Q4 Answer was not drawn by sv_answer"):
        tools_core._followup_spec(SimpleNamespace(), {"layout": {"title": "x"}}, 4)


def test_an_invalid_follow_up_is_an_input_error(monkeypatch):
    monkeypatch.setattr(sheets, "load_answer", lambda t, q: {"spec": B02, "ids": {}, "owned": None})
    with pytest.raises(tools_core.InputError, match="does not give a valid answer spec"):
        tools_core._followup_spec(SimpleNamespace(), {"pulls": [{"id": "m", "drop": True}]}, 2)


@pytest.mark.anyio
async def test_sv_answer_needs_exactly_one_of_spec_and_followup():
    with pytest.raises(ToolError, match="exactly one of question"):
        await tools_core.sv_answer()
    with pytest.raises(ToolError, match="exactly one of question"):
        await tools_core.sv_answer(spec=B02, followup={"layout": {"title": "x"}})


@pytest.mark.anyio
async def test_sv_answer_follow_up_redraws_the_answer_and_returns_the_outline(monkeypatch, tmp_path):
    from smartview_mcp import config
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    t = SimpleNamespace(path="C:/x/a.xlsx", created=False)
    monkeypatch.setattr(tools_core, "_call_user", lambda job, what: _job(job))
    monkeypatch.setattr(target, "resolve_target", lambda h, target, allow: t)
    monkeypatch.setattr(tools_core.followup_mod, "current_question", lambda t: 2)
    monkeypatch.setattr(sheets, "load_answer",
                        lambda t, q: {"spec": B02, "ids": {"m": "Q2 FinMax"}, "owned": None})
    seen, logged = {}, []
    reused = SimpleNamespace(pull=SimpleNamespace(sheet="Q2 FinMax", seconds=0.0), route="reused", note=None)

    def run(h, raw, target, allow, **kw):
        seen.update(kw, raw=raw)
        return pipeline.Run(t, 2, contract.validate(raw), {}, {"warnings": []}, [reused], {"ok": True, "findings": []},
                            [], raw=raw, reused=["m"])
    monkeypatch.setattr(tools_core.pipeline, "run", run)
    monkeypatch.setattr(sheets, "log_source", lambda *a, **k: logged.append(a))
    monkeypatch.setattr(tools_core.runlog, "write_sheet", lambda t, e: None)
    monkeypatch.setattr(tools_core.reply, "finish", lambda *a, **k: None)
    monkeypatch.setattr(tools_core.reply, "result", lambda *a, **k: {"sheet": "Q2 Answer"})
    out = await tools_core.sv_answer(followup={"layout": {"title": "New"}})
    assert seen["q"] == 2 and seen["t"] is t and seen["raw"]["layout"]["title"] == "New"
    assert out["followup_of"] == 2 and out["reused"] == ["m"] and logged == []          # reused sources not re-logged
    assert out["outline"]["pulls"][0] == {"id": "m", "connection": "FinMax"}


async def _job(job):
    return job(None)
