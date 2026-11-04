# engine/tests/v2/test_v2_save_spec.py
from smartass import pipeline
from smartass.xl import jobs, sheets


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


def test_arrange_mode_and_activation(monkeypatch):
    from smartass.core import serverkit as kit
    from smartass.xl import sheetorder
    calls = []
    monkeypatch.setattr(sheetorder, "arrange_sheets", lambda wb, mode="full", activate=None: calls.append((mode, activate)))
    monkeypatch.setattr(kit, "_is_our_analysis_file", lambda path: path.startswith("OURS"))

    class Tg:
        wb = object()

        def __init__(self, path, created):
            self.path, self.created = path, created
    assert pipeline._arrange(Tg("C:/u/book.xlsx", False), "Q3 Answer") is None
    assert pipeline._arrange(Tg("OURS/a.xlsx", False), "Q3 Answer") is None
    assert pipeline._arrange(Tg("C:/new.xlsx", True), "Q3 Answer") is None
    assert calls == [("active", "Q3 Answer"), ("full", "Q3 Answer"), ("full", "Q3 Answer")]


def test_arrange_failure_is_a_warning_not_an_error(monkeypatch):
    from smartass.core import serverkit as kit
    from smartass.xl import sheetorder

    def boom(*a, **k):
        raise RuntimeError("excel said no")
    monkeypatch.setattr(sheetorder, "arrange_sheets", boom)
    monkeypatch.setattr(kit, "_is_our_analysis_file", lambda p: True)

    class Tg:
        wb, path, created = object(), "x", True
    msg = pipeline._arrange(Tg(), "Q1 Answer")
    assert "excel said no" in msg
