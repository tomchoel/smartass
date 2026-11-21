# engine/tests/v2/test_v2_save_spec.py
import json

import fakexl
from smartass import exceljob
from smartass.xl import sheets
from smartass.xl.inproc import runner
from smartass.xl.inproc.batch import Batch


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
    """Our own file fully, the user's workbook only our sheets; the Answer is the active sheet (and landed on after a
    move) only in the active workbook - another workbook is never activated."""
    from smartass.core import serverkit as kit
    from smartass.xl import movein, sheetorder
    calls = []

    def arrange(b, wbn, rows, active, mode="full", activate=None, sources_ours=None):
        calls.append((mode, activate, active, sources_ours))
        return {"order": [n for n, _, _ in rows], "moves": range(0)}
    monkeypatch.setattr(sheetorder, "record_arrange", arrange)
    monkeypatch.setattr(kit, "_is_our_analysis_file", lambda path: path.startswith("OURS"))

    class Tg:
        wb = object()

        def __init__(self, path, created):
            self.path, self.created = path, created

    def book():
        return movein.Book([movein.Sheet("Q3 Answer"), movein.Sheet("Notes")], active="Notes")
    for tg, active in ((Tg("C:/u/book.xlsx", False), False), (Tg("OURS/a.xlsx", False), False),
                       (Tg("C:/new.xlsx", True), True)):
        assert isinstance(exceljob._arrange(tg, "Q3 Answer", book(), None, None, active, True), dict)
    assert calls == [("active", None, None, True), ("full", None, None, True),
                     ("full", "Q3 Answer", "Q3 Answer", True)]


def test_arrange_failure_is_a_warning_not_an_error(monkeypatch):
    from smartass.core import serverkit as kit
    from smartass.xl import sheetorder

    def boom(*a, **k):
        raise RuntimeError("excel said no")
    monkeypatch.setattr(sheetorder, "record_arrange", boom)
    monkeypatch.setattr(kit, "_is_our_analysis_file", lambda p: True)

    class Tg:
        wb, path, created = object(), "x", True
    from smartass.xl import movein
    msg = exceljob._arrange(Tg(), "Q1 Answer", movein.Book([movein.Sheet("Q1 Answer")]), None, None, True, True)
    assert "excel said no" in msg



# ------------------------------------------------------------------ Task 8B: one read, one write


def _meta_book():
    wb = fakexl.Workbook(["Q1 Answer", "Sources"], styles_ver=False)
    meta = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
    sheets._name_hidden(meta, sheets.META_SHEET)
    t = T.__new__(T)
    t.app, t.book_name = wb.Application, wb.Name
    return wb, meta, t


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
    #                                                                         (a batch reads the rows, one writes)
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


def test_the_hidden_meta_sheet_is_made_in_the_batch_that_stores_the_spec_in_its_row_two(monkeypatch):
    wb = fakexl.Workbook(["Q1 Answer", "Sources"], styles_ver=False)
    t = T.__new__(T)
    t.wb, t.app, t.book_name = wb, wb.Application, wb.Name
    calls = fakexl.count_com(monkeypatch)
    assert sheets.save_spec(t, 2, {"v2": {"a": 1}}, ids={}, create=True) is True
    assert len(calls) == 2                              # a batch reads the rows; one adds, names, hides and writes
    meta = wb.Worksheets(sheets.META_SHEET)
    assert meta.Visible == 0 and [s.Name for s in wb.order][-1] == sheets.META_SHEET
    assert sheets._stored_specs(t)[2]["spec"] == {"v2": {"a": 1}} and (2, 1) in meta.cells
    assert sheets.save_spec(t, 3, {}, create=True) is False
    assert [s.Name for s in wb.order].count(sheets.META_SHEET) == 1


def test_the_rows_of_the_stored_specs_are_read_without_the_specs():
    """Only column A of _svmeta's used range is read (the specs grow per answer); a key found is written over as wide
    as the used range, a new one goes to the first free row."""
    wb, meta, t = _meta_book()
    sheets.save_spec(t, 1, {"v2": {"question": "x" * 70000}})
    sheets.save_spec(t, 2, {"v2": {"question": "y"}})
    b = Batch(quiet=False)
    get = sheets.record_spec_rows(b, b.workbook(wb.Name))
    ops = b.to_dict()["ops"]
    assert {"k": "call", "name": "Columns", "args": [1]}.items() <= next(
        op for op in ops if op.get("name") == "Columns" and op["k"] == "call").items()
    assert not [op for op in ops if op.get("name") in ("Value", "Value2") and op["on"] == 3]   # never the used range
    rows = get(runner.run(wb.Application, b))
    assert rows.sheet and rows.keys == ("SPEC Q1", "SPEC Q2") and rows.top == 2 and rows.width == 4
    assert rows.row("SPEC Q2") == (3, 4) and rows.row("SPEC Q7") == (4, 0)
