"""The tool-owned range of an answer sheet: a redraw clears only what the tool drew; user content survives."""
from types import SimpleNamespace

from smartview_mcp import answer_write
from smartview_mcp.core import owned


def test_growth_inserts_rows_only_when_the_band_below_holds_user_content():
    used = lambda r1, c1, r2, c2: (r1, c1, r2, c2) == (31, 1, 40, 9)         # user text in A31:I40
    assert owned.growth({"rows": 30, "cols": 9}, {"rows": 40, "cols": 9}, used) == {"rows": (31, 40), "cols": None}
    assert owned.growth({"rows": 30, "cols": 9}, {"rows": 40, "cols": 9}, lambda *a: False) == \
        {"rows": None, "cols": None}
    assert owned.growth({"rows": 30, "cols": 9}, {"rows": 20, "cols": 9}, used) == {"rows": None, "cols": None}
    assert owned.growth(None, {"rows": 40, "cols": 9}, used) == {"rows": None, "cols": None}


def test_growth_inserts_columns_when_the_band_to_the_right_holds_user_content():
    used = lambda r1, c1, r2, c2: c1 == 10
    g = owned.growth({"rows": 30, "cols": 9}, {"rows": 30, "cols": 12}, used)
    assert g == {"rows": None, "cols": (10, 12)}


class Shape:
    def __init__(self, name):
        self.Name, self.deleted = name, False

    def Delete(self):
        self.deleted = True


def test_only_smartass_charts_are_deleted():
    mine, users = Shape(owned.CHART_NAME), Shape("Chart 7")
    ws = SimpleNamespace(ChartObjects=lambda: [mine, users])
    owned.drop_charts(ws, com=lambda call, **kw: call())
    assert mine.deleted and not users.deleted


def test_without_an_owned_range_the_whole_sheet_is_cleared_as_before():
    calls = []
    ws = SimpleNamespace(Cells=SimpleNamespace(Clear=lambda: calls.append("all")),
                         Range=lambda a, b: SimpleNamespace(Clear=lambda: calls.append(("block", a, b))),
                         ChartObjects=lambda: [])
    ws_cells = lambda r, c: (r, c)
    owned.clear(ws, None, com=lambda call, **kw: call(), cell=ws_cells)
    owned.clear(ws, {"rows": 12, "cols": 6}, com=lambda call, **kw: call(), cell=ws_cells)
    assert calls == ["all", ("block", (1, 1), (12, 6))]


class _Sheet:
    def __init__(self, log):
        self.log = log

    def Cells(self, r, c):
        return SimpleNamespace(r=r, c=c)

    def Range(self, a, b):
        log = self.log

        class R:
            def __setattr__(self, k, v):
                log.append(("write", k))
        return R()


def test_room_is_made_for_the_extent_before_the_cells_are_written(monkeypatch):
    monkeypatch.setattr(answer_write, "com", lambda call, *a, **k: call())
    log = []
    spec = {"title": "T", "columns": [{"header": "A"}, {"header": "B"}],
            "sections": [{"title": "s", "lines": [{"key": "a", "label": "A", "cells": [{"value": 1}, {"value": 2}]}]}],
            "notes": ["n"]}
    out = answer_write._write(SimpleNamespace(status=lambda m: None), 1, spec, {}, _Sheet(log),
                        make_room=lambda ext: log.append(("room", ext)))
    r = out[-2]
    assert log[0] == ("room", {"rows": r - 1, "cols": answer_write.FIRST_COL + 1}) and log[1][0] == "write"


def test_findings_rows_are_severity_then_text_capped_with_a_tail():
    from smartview_mcp.core import findings
    fs = [{"severity": "Critical", "text": "EBITDA check DIFF"}] + \
         [{"severity": "Info", "text": f"note {i}"} for i in range(20)]
    rows = findings.rows(fs, limit=15)
    assert rows[0] == ("Findings", "section")
    assert rows[1] == ("Critical — EBITDA check DIFF", "critical")
    assert rows[-1] == ("… and 6 more (see the tool result)", "note") and len(rows) == 1 + 15 + 1


def test_no_findings_is_one_line():
    from smartview_mcp.core import findings
    assert findings.rows([]) == [("Findings", "section"), ("No findings: all checks OK.", "note")]


def test_render_writes_one_block_with_a_style_per_row_and_returns_the_last_row():
    from smartview_mcp.core import findings
    log = []

    class Cell:
        def __init__(self, r):
            object.__setattr__(self, "r", r)
            object.__setattr__(self, "Font", SimpleNamespace())

        def __setattr__(self, k, v):
            log.append(("cell", self.r, k, v))

    class Rng:
        def __init__(self, a, b):
            object.__setattr__(self, "span", (a.r, b.r))

        def __setattr__(self, k, v):
            log.append(("range", self.span, k, v))
    ws = SimpleNamespace(Cells=lambda r, c: Cell(r), Range=lambda a, b: Rng(a, b))
    fs = [{"severity": "Critical", "text": "c"}, {"severity": "Warning", "text": "w"}]
    last = findings.render(ws, 30, fs, com=lambda call, **kw: call())
    assert last == 32
    assert log[0] == ("range", (30, 32), "Value", (("Findings",), ("Critical — c",), ("Warning — w",)))
    assert ("cell", 30, "Style", "TN Seksjon") in log and ("cell", 32, "Style", "TN Note") in log
    assert len([e for e in log if e[0] == "cell" and e[2] == "Style"]) == 3
