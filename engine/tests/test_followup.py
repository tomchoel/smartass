"""Phase 5 follow-ups: stored spec + owned range, source plan, follow-up pipeline, sv_answer(followup=...)."""
import json
from pathlib import Path
from types import SimpleNamespace

from smartview_mcp import userexcel

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
