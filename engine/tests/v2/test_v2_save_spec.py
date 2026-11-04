# engine/tests/v2/test_v2_save_spec.py
from smartass import pipeline
from smartview_mcp.excel import jobs, sheets


class Sheet:
    def __init__(self, name):
        self.Name, self.Visible = name, -1


class Sheets(list):
    def __call__(self, i):
        if isinstance(i, int):
            return self[i - 1]
        return next(s for s in self if s.Name == i)

    @property
    def Count(self):
        return len(self)

    def Add(self, before=None, after=None):
        s = Sheet("Sheet9")
        self.append(s)
        return s


class WB:
    def __init__(self):
        self.Worksheets = Sheets([Sheet("Q1 Answer")])


class T:
    def __init__(self):
        self.wb = WB()


def test_save_spec_creates_the_hidden_meta_sheet_first(monkeypatch):
    monkeypatch.setattr(jobs, "com", lambda call, *a, **k: call())
    seen = {}

    def fake_save(t, q, spec, ids=None, owned=None):
        seen["exists"] = sheets._sheet_exists(t, sheets.META_SHEET)
        seen["args"] = (q, spec)
    monkeypatch.setattr(sheets, "save_spec", fake_save)
    t = T()
    pipeline._save_spec(t, 2, {"a": 1})
    assert seen["exists"] and seen["args"] == (2, {"v2": {"a": 1}})
    meta = t.wb.Worksheets(sheets.META_SHEET)
    assert meta.Visible == 0
    pipeline._save_spec(t, 3, {})                       # second call: no second meta sheet
    assert [s.Name for s in t.wb.Worksheets].count(sheets.META_SHEET) == 1
