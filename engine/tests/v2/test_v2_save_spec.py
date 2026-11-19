# engine/tests/v2/test_v2_save_spec.py
import json

import fakexl
from smartass import pipeline
from smartass.xl import sheets


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



# ------------------------------------------------------------------ Task 8B: one read, one write


def _meta_book():
    wb = fakexl.Workbook(["Q1 Answer", "Sources"], styles_ver=False)
    meta = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
    sheets._name_hidden(meta, sheets.META_SHEET)
    return wb, meta, T.__new__(T)


def _row(meta, r):
    return [meta.cells.get((r, c), {}).get("Value") for c in range(1, 6)]


def test_the_spec_is_stored_with_one_read_and_one_write(monkeypatch):
    wb, meta, t = _meta_book()
    t.wb = wb
    meta.put(2, 1, "Q1 FinMax")                         # a save_meta row of v1 above
    sheets.save_spec(t, 1, {"v2": {"question": "x" * 70000}})
    calls = fakexl.count_com(monkeypatch)
    meta.log.clear()
    sheets.save_spec(t, 1, {"v2": {"question": "y"}}, owned=None)
    assert len(calls) == 2 and [k for k, _ in meta.log] == ["Value"]         # A and every chunk: one rectangle
    r = next(r for (r, c), p in meta.cells.items() if c == 1 and p.get("Value") == "SPEC Q1")
    assert r == 3 and _row(meta, r)[2:] == [None, None, None]                 # the longer spec's chunks are gone
    assert json.loads(_row(meta, r)[1]) == {"spec": {"v2": {"question": "y"}}, "ids": {}, "owned": None}
    assert (r, 1) in meta.prefixed and (r, 2) in meta.prefixed


def test_a_long_spec_is_cut_in_chunks_in_the_same_rectangle(monkeypatch):
    wb, meta, t = _meta_book()
    t.wb = wb
    calls = fakexl.count_com(monkeypatch)
    sheets.save_spec(t, 4, {"v2": {"question": "z" * 65000}})
    assert len(calls) == 2
    stored = sheets._stored_specs(t)
    assert stored[4]["spec"]["v2"]["question"] == "z" * 65000
    assert [k for k, _ in meta.log] == ["Value"] and len([c for (r, c) in meta.cells if r == 2]) == 4


def test_the_pipeline_makes_the_hidden_meta_sheet_and_stores_the_spec_in_its_row_two(monkeypatch):
    wb = fakexl.Workbook(["Q1 Answer", "Sources"], styles_ver=False)
    t = T.__new__(T)
    t.wb = wb
    calls = fakexl.count_com(monkeypatch)
    assert pipeline._save_spec(t, 2, {"a": 1}) is True
    assert len(calls) == 4                              # the read, Add, name it hidden, the write
    meta = wb.Worksheets(sheets.META_SHEET)
    assert meta.Visible == 0 and [s.Name for s in wb.order][-1] == sheets.META_SHEET
    assert sheets._stored_specs(t)[2]["spec"] == {"v2": {"a": 1}} and (2, 1) in meta.cells
    assert pipeline._save_spec(t, 3, {}) is False and [s.Name for s in wb.order].count(sheets.META_SHEET) == 1
