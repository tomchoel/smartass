"""answer._Cells: the answer sheet is written in one rectangle, not one COM call per cell."""
import pytest

from smartview_mcp import answer
from smartview_mcp.host import HostError


class Cell:
    def __init__(self, ws, r, c, fail=False):
        self.ws, self.r, self.c, self.fail = ws, r, c, fail

    def __setattr__(self, k, v):
        if k in ("Formula2", "Value"):
            if self.fail:
                raise ValueError("bad formula")
            self.ws.cells[(self.r, self.c)] = (k, v)
        object.__setattr__(self, k, v)


class Range:
    def __init__(self, ws, a, b):
        self.ws, self.a, self.b = ws, a, b

    def __setattr__(self, k, v):
        if k == "Formula2":
            if self.ws.batch_fails:
                raise ValueError("one cell in the batch is bad")
            self.ws.batches.append(((self.a.r, self.a.c), (self.b.r, self.b.c), v))
        object.__setattr__(self, k, v)


class Sheet:
    def __init__(self, batch_fails=False, bad=()):
        self.cells, self.batches, self.batch_fails, self.bad = {}, [], batch_fails, set(bad)

    def Cells(self, r, c):
        return Cell(self, r, c, fail=(r, c) in self.bad)

    def Range(self, a, b):
        return Range(self, a, b)


def test_one_rectangle_for_all_cells(monkeypatch):
    monkeypatch.setattr(answer, "com", lambda call, **kw: call())
    ws, warnings = Sheet(), []
    cells = answer._Cells()
    cells.put(1, 2, "'Title")
    cells.put(3, 3, "=SUM(C1:C2)", "EBITDA")
    cells.put(2, 4, 5.0)
    cells.flush(ws, warnings)
    assert ws.batches == [((1, 2), (3, 4), (("'Title", "", ""), ("", "", 5.0), ("", "=SUM(C1:C2)", "")))]
    assert ws.cells == {} and warnings == []


def test_a_refused_batch_falls_back_to_cells_and_warns_only_for_the_bad_one(monkeypatch):
    monkeypatch.setattr(answer, "com", lambda call, **kw: call())
    ws, warnings = Sheet(batch_fails=True, bad={(3, 3)}), []
    cells = answer._Cells()
    cells.put(1, 2, "'Title")
    cells.put(3, 3, "=BAD(", "EBITDA")
    cells.put(2, 4, 5.0)
    cells.flush(ws, warnings)
    assert ws.cells == {(1, 2): ("Value", "'Title"), (2, 4): ("Value", 5.0)}
    assert len(warnings) == 1 and warnings[0].startswith("EBITDA: ")


def test_excel_busy_is_not_swallowed(monkeypatch):
    def busy(call, **kw):
        raise HostError("Excel is busy")
    monkeypatch.setattr(answer, "com", busy)
    cells = answer._Cells()
    cells.put(1, 2, "'x")
    with pytest.raises(HostError):
        cells.flush(Sheet(), [])


def test_nothing_to_write_is_no_call():
    ws = Sheet()
    answer._Cells().flush(ws, [])
    assert ws.batches == [] and ws.cells == {}


def test_check_cells_are_grouped_into_column_runs():
    from smartview_mcp import styling
    assert styling._check_ranges([(10, 4), (11, 4), (12, 4), (15, 4), (3, 6)]) == ["D10:D12", "D15", "F3"]
    assert styling._check_ranges([]) == []
