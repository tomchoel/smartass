"""Phase 5 follow-ups: stored spec + owned range, source plan, follow-up pipeline, sv_answer(followup=...)."""
import json
from pathlib import Path
from types import SimpleNamespace

from smartview_mcp import userexcel
from smartview_mcp.core import followup

B02 = json.loads((Path(__file__).parent / "fixtures" / "b02_answer_spec.json").read_text(encoding="utf-8"))


class Meta:
    """A fake _svmeta sheet: rows of cell values (column A = key)."""
    def __init__(self):
        self.rows: dict[int, dict[int, str]] = {}

    @property
    def UsedRange(self):
        n = max(self.rows, default=1)
        w = max((max(r) for r in self.rows.values() if r), default=1)
        return SimpleNamespace(Value=tuple(tuple(self.rows.get(r, {}).get(c) for c in range(1, w + 1))
                                           for r in range(1, n + 1)))

    def Cells(self, r, c):
        meta = self

        class C:
            def __setattr__(self, k, v):
                meta.rows.setdefault(r, {})[c] = v[1:] if isinstance(v, str) and v.startswith("'") else v

            @property
            def End(self):
                return lambda d: SimpleNamespace(Row=max(meta.rows, default=0))
        return C()

    def Rows(self, r):
        return SimpleNamespace(ClearContents=lambda: self.rows.pop(r, None))

    Rows.Count = 1048576


def _meta_target(monkeypatch):
    ws = Meta()
    t = SimpleNamespace(wb=SimpleNamespace(Worksheets=lambda name: ws))
    monkeypatch.setattr(userexcel, "com", lambda call, **kw: call())
    monkeypatch.setattr(userexcel, "_sheet_exists", lambda t, name: name == userexcel.META_SHEET)
    return t, ws


def test_the_spec_is_stored_with_its_owned_range_and_read_back(monkeypatch):
    t, _ = _meta_target(monkeypatch)
    userexcel.save_spec(t, 2, B02, {"m": "Q2 FinMax"}, owned={"rows": 40, "cols": 7})
    got = userexcel.load_answer(t, 2)
    assert got == {"spec": B02, "ids": {"m": "Q2 FinMax"}, "owned": {"rows": 40, "cols": 7}}
    assert userexcel.load_spec(t, 2) == (B02, {"m": "Q2 FinMax"})          # the old reader still works
    assert userexcel.load_answer(t, 3) is None


def test_stored_questions_lists_only_v1_answer_specs(monkeypatch):
    t, _ = _meta_target(monkeypatch)
    userexcel.save_spec(t, 3, B02, {})
    userexcel.save_spec(t, 1, {"question": "an old sv_ws_build spec"}, {})
    userexcel.save_spec(t, 2, B02, {}, owned={"rows": 5, "cols": 3})
    userexcel.save_spec(t, 3, B02, {"m": "Q3 FinMax"})                       # overwrites its row
    assert userexcel.stored_questions(t) == [2, 3]


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
    lay = {"expect": {"source": "Warehouse", "cells": {"pnl": [[1, 1, 5.0, "f", 0.5]], "fte": [[1, 1, 2.0, "g", 0.5]]}}}
    assert followup.drop_stale_expect(lay, ["fte"])["expect"]["cells"] == {"pnl": [[1, 1, 5.0, "f", 0.5]]}
    assert "expect" not in followup.drop_stale_expect({"expect": {"source": "W", "cells": {"fte": []}}}, ["fte"])
    assert followup.drop_stale_expect({"title": "x"}, ["fte"]) == {"title": "x"}


def test_the_current_question_is_the_active_answer_sheet_else_the_latest_stored(monkeypatch):
    t = SimpleNamespace()
    monkeypatch.setattr(followup, "_active_sheet", lambda t: "Q3 Answer")
    monkeypatch.setattr(userexcel, "stored_questions", lambda t: [1, 3, 5])
    assert followup.current_question(t) == 3
    monkeypatch.setattr(followup, "_active_sheet", lambda t: "Kilder")
    monkeypatch.setattr(userexcel, "stored_questions", lambda t: [1, 4])
    assert followup.current_question(t) == 4
    monkeypatch.setattr(userexcel, "stored_questions", lambda t: [])
    assert followup.current_question(t) is None
