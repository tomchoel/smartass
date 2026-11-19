"""Phase 3 Task 4: the built file's sheets moved into the user's workbook in one copy and finished there
(xl/movein.py): the name pre-check, the hidden read-only temp workbook, the tables, the formula blocks, the chart
label fixes, the user's settings put back, and the park / restore / drop steps of a follow-up. All on fakexl."""
from __future__ import annotations

import logging
from types import SimpleNamespace

import pytest
import xlsxwriter

import fakexl
from smartass.render.xlsx import Base, BuiltFile, FormulaBlock, TableSpec
from smartass.render.xlsx_charts import LabelFix
from smartass.xl import jobs, movein
from smartass.xl.host import HostError, InputError
from smartass.xl.target import Target

USER_SHEETS = ["Q1 Answer", "Q1 Data", "Notes", "Sources", "_svmeta"]
SHEETS = ["Q3 Answer", "Q3 Data", "Q3 Lines b0"]
ANSWER = FormulaBlock("Q3 Answer", "B2:C3", [["'Revenue", "=SUM(Data_Q3[Value])"], ["'=SUM(A1)", "'0047"]])
HEADER = FormulaBlock("Q3 Data", "A1:B2", [["'Q3 Data", ""], ["'Control total", "=SUM(Data_Q3[Value])"]])
LINES = FormulaBlock("Q3 Lines b0", "F1:G1", [["'Lines total", "=SUBTOTAL(9,Lines_Q3_b0[AMOUNT])"]])
TABLES = [TableSpec("Q3 Data", "A4:C6", "Data_Q3", stripes=False),
          TableSpec("Q3 Lines b0", "A4:B6", "Lines_Q3_b0", stripes=True)]
FIXES = [LabelFix(1, 2, 3, position="right"), LabelFix(1, 1, 2, font_color="#1F3864", lift_pt=10.0)]
NAVY_BGR = 0x64381F                                  # "#1F3864" as a COM colour (blue * 65536 + green * 256 + red)


# ------------------------------------------------------------------ the built file and the user's workbook
def _file(path, extra: str | None = None, table: bool = False, link: bool = False):
    """A three-sheet file as render/xlsx.build writes it: values only, no table (unless table=True)."""
    with xlsxwriter.Workbook(str(path), {"strings_to_formulas": False}) as wb:
        answer = wb.add_worksheet("Q3 Answer")
        data = wb.add_worksheet("Q3 Data")
        for j, name in enumerate(("Period", "Account", "Value")):
            data.write_string(3, j, name)
        for i in (4, 5):
            data.write_string(i, 0, "Sep 26")
            data.write_string(i, 1, "0047")
            data.write_number(i, 2, 100.0 * i)
        if table:
            data.add_table("A4:C6", {"name": "Data_Q3", "header_row": True})
        lines = wb.add_worksheet("Q3 Lines b0")
        lines.write_string(3, 0, "ACCOUNT")
        lines.write_string(3, 1, "AMOUNT")
        if extra:
            wb.add_worksheet(extra).write_number(0, 0, 1.0)
        if link:
            answer.write_formula(0, 0, f"='{extra}'!A1")
    return BuiltFile(path=path, sheets=list(SHEETS), tables=list(TABLES), formulas=[ANSWER, HEADER, LINES],
                     label_fixes=list(FIXES))


@pytest.fixture
def built(tmp_path):
    return _file(tmp_path / "answer.xlsx")


@pytest.fixture
def t():
    wb = fakexl.Workbook(USER_SHEETS, styles_ver=False)
    wb.Worksheets("_svmeta").Visible = 0                    # xlSheetHidden, as sheets.save_spec leaves it
    wb.active = wb.Worksheets("Q1 Answer")
    return Target(app=wb.Application, wb=wb, pid=0, hwnd=0)


def _names(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def _tables(wb) -> dict[str, str]:
    return {lo.Name: ws.Name for ws in wb.Worksheets for lo in ws.ListObjects}


def _events(t, *kinds) -> list[tuple]:
    return [e for e in t.app.events if e[1] in kinds]


def _add_table(ws, name: str, ref: str = "A1:B3"):
    lo = ws.ListObjects.Add(1, ws.Range(ref), None, 1)
    lo.Name = name
    return lo


class _Props:
    """A data label (or its Font): every property set is an event of the fake Excel, in order with the others."""

    def __init__(self, ws, key, prefix="", **start):
        object.__setattr__(self, "_ws", ws)
        object.__setattr__(self, "_key", key)
        object.__setattr__(self, "_prefix", prefix)
        for k, v in start.items():
            object.__setattr__(self, k, v)

    def __setattr__(self, k, v):
        fakexl._event(self._ws, "DataLabel", self._key, self._prefix + k, v)
        object.__setattr__(self, k, v)


def charts_on(ws, count: int = 1, series: int = 2, fail=None):
    """ws.ChartObjects(i).Chart.SeriesCollection(s).Points(p).DataLabel for count charts of series series each;
    a chart or series beyond those is a COM error, as in Excel; fail: raised by ChartObjects instead."""
    labels: dict = {}

    def label(c, s, p):
        if s > series:
            raise fakexl.ComError(f"no series {s} in chart {c}")
        if (c, s, p) not in labels:
            lab = _Props(ws, (c, s, p), Top=50.0)
            object.__setattr__(lab, "Font", _Props(ws, (c, s, p), "Font."))
            labels[(c, s, p)] = lab
        return labels[(c, s, p)]

    def chart_objects(c):
        if fail is not None:
            raise fail
        if c > count:
            raise fakexl.ComError(f"no chart {c}")
        points = lambda s: SimpleNamespace(Points=lambda p: SimpleNamespace(DataLabel=label(c, s, p)))  # noqa: E731
        return SimpleNamespace(Chart=SimpleNamespace(SeriesCollection=points))
    ws.ChartObjects = chart_objects
    return labels


def _moved(t, built):
    movein.move_in(t, built, after=t.wb.Worksheets("Sources"))
    charts_on(t.wb.Worksheets("Q3 Answer"))
    t.app.events.clear()


# ------------------------------------------------------------------ move_in
def test_move_in_is_one_copy_and_at_most_eight_com_calls(t, built, monkeypatch):
    opened = []
    real_open = t.app.Workbooks.Open
    monkeypatch.setattr(t.app.Workbooks, "Open", lambda *a: opened.append(real_open(*a)) or opened[-1])
    calls = fakexl.count_com(monkeypatch)
    movein.move_in(t, built, after=t.wb.Worksheets("Sources"))
    assert len(calls) <= 8 and len(_events(t, "Copy")) == 1
    assert _names(t.wb) == ["Q1 Answer", "Q1 Data", "Notes", "Sources", *SHEETS, "_svmeta"]
    assert t.sheets_added == SHEETS
    assert t.wb.LinkSources(1) is None and t.app.Workbooks.Count == 1          # the temp file is closed
    assert opened[0].opened_with == (0, True)                                    # no link update, read-only
    kinds = [e[1] for e in t.app.events if e[1] in ("Window.Visible", "Copy", "Close")]
    assert kinds == ["Window.Visible", "Copy", "Close"] and opened[0].window.Visible is False


def test_the_copied_sheets_carry_the_files_values(t, built):
    movein.move_in(t, built, after=None)
    data = t.wb.Worksheets("Q3 Data")
    assert data.Range("A4:C5").Value == (("Period", "Account", "Value"), ("Sep 26", "0047", 400.0))
    assert _tables(t.wb) == {}                                                   # the tables come with finish


@pytest.mark.parametrize("after", [None, "_svmeta"])
def test_after_a_hidden_sheet_the_copy_lands_after_the_last_visible_one(t, built, after):
    movein.move_in(t, built, after=t.wb.Worksheets(after) if after else None)
    assert _names(t.wb) == ["Q1 Answer", "Q1 Data", "Notes", "Sources", *SHEETS, "_svmeta"]
    assert all(t.wb.Worksheets(n).Visible == -1 for n in SHEETS)


def test_with_no_visible_sheet_up_to_a_hidden_anchor_the_copy_goes_before_the_first_visible_one(built):
    wb = fakexl.Workbook(["_svmeta", "Notes"], styles_ver=False)
    wb.Worksheets("_svmeta").Visible = 0
    wb.active = wb.Worksheets("Notes")
    t = Target(app=wb.Application, wb=wb, pid=0, hwnd=0)
    movein.move_in(t, built, after=wb.Worksheets("_svmeta"))
    assert _names(wb) == ["_svmeta", *SHEETS, "Notes"]


def test_a_link_to_a_sheet_left_behind_is_broken_and_the_users_own_links_stay(t, tmp_path):
    users = r"C:\Users\someone\Budget 2026.xlsx"
    t.wb.links.append(users)                                  # the user's formulas that read another workbook
    built = _file(tmp_path / "linked.xlsx", extra="Scratch", link=True)
    movein.move_in(t, built, after=None)
    assert [e[2] for e in _events(t, "BreakLink")] == [str(built.path)] and t.wb.LinkSources(1) == (users,)
    assert "Scratch" not in _names(t.wb)


def test_first_copies_before_the_first_visible_sheet(t, built):
    t.wb.order[0].Visible = 0                            # the user hid their first sheet
    movein.move_in(t, built, None, first=True)
    assert _names(t.wb) == ["Q1 Answer", *SHEETS, "Q1 Data", "Notes", "Sources", "_svmeta"]


def test_a_built_file_with_a_table_is_refused_before_anything_is_copied(t, tmp_path):
    built = _file(tmp_path / "tabled.xlsx", table=True)
    before = fakexl.snapshot(t.wb)
    with pytest.raises(ValueError, match="Data_Q3") as e:
        movein.move_in(t, built, after=None)
    assert not isinstance(e.value, InputError)
    assert fakexl.snapshot(t.wb) == before and t.app.Workbooks.Count == 1 and t.sheets_added == []
    assert not _events(t, "Copy")


@pytest.mark.parametrize("leftover", ["sheet:Q3 Data", "sheet:q3 answer", "table:Data_Q3",
                                      "table:Lines_Q3_b0@Other sheet"])
def test_a_taken_name_stops_before_anything_is_copied(t, built, leftover):     # Review Focus 2
    kind, _, rest = leftover.partition(":")
    name, _, sheet = rest.partition("@")
    if kind == "sheet":
        t.wb.order.insert(2, fakexl.Worksheet(t.wb, name))
    else:
        if sheet:
            t.wb.order.insert(2, fakexl.Worksheet(t.wb, sheet))
        _add_table(t.wb.Worksheets(sheet or "Q1 Data"), name)
    before = fakexl.snapshot(t.wb)
    with pytest.raises(movein.NameTaken, match=name) as e:
        movein.move_in(t, built, after=None)
    assert isinstance(e.value, InputError) and "book.xlsx" in str(e.value) and "rename or delete" in str(e.value)
    assert fakexl.snapshot(t.wb) == before and t.app.Workbooks.Count == 1 and t.sheets_added == []


def test_taken_names_every_sheet_and_table_already_there(t):
    _add_table(t.wb.Worksheets("Notes"), "lines_q3_b0")
    t.wb.order.insert(0, fakexl.Worksheet(t.wb, "Q3 Data"))
    assert movein.taken(t.wb, SHEETS, ["Data_Q3", "Lines_Q3_b0"]) == [
        "the sheet 'Q3 Data'", "the table 'lines_q3_b0' on sheet 'Notes'"]
    assert movein.taken(t.wb, ["Q4 Answer"], ["Data_Q4"]) == []


def test_the_temp_workbook_is_closed_when_the_copy_fails(t, built, monkeypatch):
    def refuse(self, *a):
        raise fakexl.ComError("Copy method of Sheets class failed")
    monkeypatch.setattr(fakexl._SheetGroup, "Copy", refuse)
    before = fakexl.snapshot(t.wb)
    with pytest.raises(fakexl.ComError, match="Copy method"):
        movein.move_in(t, built, after=None)
    assert t.app.Workbooks.Count == 1 and t.app.ActiveWorkbook is t.wb and fakexl.snapshot(t.wb) == before
    assert t.sheets_added == SHEETS                     # whatever a half-done copy brought in, rollback removes


# ------------------------------------------------------------------ finish
def test_finish_makes_the_tables_then_writes_each_block_once_then_fixes_the_labels(t, built):
    _moved(t, built)
    movein.finish(t, built)
    steps = [(e[0], e[1]) for e in _events(t, "ListObjects.Add", "Formula2", "DataLabel")]
    assert steps[:2] == [("Q3 Data", "ListObjects.Add"), ("Q3 Lines b0", "ListObjects.Add")]
    assert steps[2:5] == [("Q3 Answer", "Formula2"), ("Q3 Data", "Formula2"), ("Q3 Lines b0", "Formula2")]
    assert {s for s in steps[5:]} == {("Q3 Answer", "DataLabel")} and len(steps) > 5
    assert [e[2] for e in _events(t, "Formula2")] == ["B2:C3", "A1:B2", "F1:G1"]
    data, lines = t.wb.Worksheets("Q3 Data").ListObjects(1), t.wb.Worksheets("Q3 Lines b0").ListObjects(1)
    assert _tables(t.wb) == {"Data_Q3": "Q3 Data", "Lines_Q3_b0": "Q3 Lines b0"}
    assert (fakexl.rect_text(*data.rect), fakexl.rect_text(*lines.rect)) == ("A4:C6", "A4:B6")
    assert data.TableStyle == lines.TableStyle == "TableStyleLight1"
    assert data.ShowTableStyleRowStripes is False and lines.ShowTableStyleRowStripes is True
    answer = t.wb.Worksheets("Q3 Answer")
    assert answer.Range("C2").Value == "=SUM(Data_Q3[Value])"


def test_the_label_fixes_set_position_colour_and_lift(t, built):
    movein.move_in(t, built, after=None)
    labels = charts_on(t.wb.Worksheets("Q3 Answer"))
    movein.finish(t, built)
    right, thin = labels[(1, 2, 3)], labels[(1, 1, 2)]
    assert right.Position == -4152 and not hasattr(right.Font, "Color") and right.Top == 50.0
    assert thin.Font.Color == NAVY_BGR and thin.Top == 40.0 and not hasattr(thin, "Position")


def test_a_thin_step_label_at_most_its_lift_from_the_top_is_coloured_but_not_moved(t, built):
    """As charts._labels: only a label more than the lift below the chart's top is lifted (live, 19.11.2026: the
    empty points of a waterfall's other step series kept their position in the COM drawing)."""
    built.label_fixes = [LabelFix(1, 1, 2, font_color="#1F3864", lift_pt=10.0)]
    movein.move_in(t, built, after=None)
    ws = t.wb.Worksheets("Q3 Answer")
    labels = charts_on(ws)
    empty = ws.ChartObjects(1).Chart.SeriesCollection(1).Points(2).DataLabel
    object.__setattr__(empty, "Top", 10.0)
    movein.finish(t, built)
    assert labels[(1, 1, 2)].Font.Color == NAVY_BGR and labels[(1, 1, 2)].Top == 10.0


@pytest.mark.parametrize("position,code", [("right", -4152), ("above", 0), ("below", 1), ("left", -4131)])
def test_each_label_position_is_excels_constant(t, built, position, code):
    built.label_fixes = [LabelFix(1, 1, 1, position=position)]
    movein.move_in(t, built, after=None)
    labels = charts_on(t.wb.Worksheets("Q3 Answer"))
    movein.finish(t, built)
    assert labels[(1, 1, 1)].Position == code


def test_a_label_fix_that_fails_is_skipped_with_a_log_line(t, built, caplog):
    built.label_fixes = [LabelFix(2, 1, 1, position="above"), LabelFix(1, 3, 1, position="left"),
                         LabelFix(1, 1, 2, position="below")]
    movein.move_in(t, built, after=None)
    labels = charts_on(t.wb.Worksheets("Q3 Answer"))
    with caplog.at_level(logging.WARNING):
        movein.finish(t, built)
    assert labels[(1, 1, 2)].Position == 1 and list(labels) == [(1, 1, 2)]
    skipped = [r.getMessage() for r in caplog.records if "label" in r.getMessage()]
    assert len(skipped) == 2 and "chart 2" in skipped[0] and "series 3" in skipped[1]


def test_an_abandoned_job_is_not_taken_for_a_failed_label(t, built):
    movein.move_in(t, built, after=None)
    charts_on(t.wb.Worksheets("Q3 Answer"), fail=jobs.JobAborted("abandoned"))
    with pytest.raises(jobs.JobAborted):
        movein.finish(t, built)


def test_formula_grids_keep_text_as_text(t, built):                             # Review Focus 1
    _moved(t, built)
    movein.finish(t, built)
    answer, data = t.wb.Worksheets("Q3 Answer"), t.wb.Worksheets("Q3 Data")
    assert answer.Range("B3").Value == "=SUM(A1)" and answer.Range("B3").PrefixCharacter == "'"
    assert answer.Range("C3").Value == "0047" and answer.Range("C3").PrefixCharacter == "'"
    assert answer.Range("C2").PrefixCharacter == "" and data.Range("A2").PrefixCharacter == "'"
    assert data.Range("B2").Value == "=SUM(Data_Q3[Value])" and data.Range("B2").PrefixCharacter == ""


def test_text_that_looks_like_a_formula_number_or_date_stays_text(t, built):     # Review Focus 1
    texts = ["=SUM(A1)", "+1", "-2", "@x", "1-2", "3/4", "1E5", "0047", "TRUE"]
    built.formulas = [FormulaBlock("Q3 Answer", "B5:J5", [["'" + x for x in texts]])]
    movein.move_in(t, built, after=None)
    movein.finish(t, built)
    row = [t.wb.Worksheets("Q3 Answer").Cells(5, c) for c in range(2, 11)]
    assert [c.Value for c in row] == texts and {c.PrefixCharacter for c in row} == {"'"}


def test_calculate_does_the_answer_last(t):
    movein.calculate(t, ["Q1 Answer", "Q1 Data", "Notes"])
    assert [e[0] for e in _events(t, "Calculate")] == ["Q1 Data", "Notes", "Q1 Answer"]


# ------------------------------------------------------------------ quiet
class FailingApp(fakexl.Application):
    """Setting fail_on (a name, or (name, value)) raises, as a COM error would."""

    def __setattr__(self, k, v):
        fail = self.__dict__.get("fail_on")
        if fail in (k, (k, v)):
            raise fakexl.ComError(f"cannot set {k}")
        super().__setattr__(k, v)


SETTINGS = ("ScreenUpdating", "EnableEvents", "DisplayAlerts", "Calculation")


@pytest.fixture
def app():
    a = FailingApp()
    user = fakexl.Workbook(["Q1 Answer", "Notes"], app=a, styles_ver=False)
    fakexl.Workbook(["Other"], app=a, styles_ver=False)
    user.Worksheets("Notes").Activate()
    return a


def _state(app):
    return {k: getattr(app, k) for k in SETTINGS} | {"book": app.ActiveWorkbook.Name, "sheet": app.ActiveSheet.Name}


def test_quiet_turns_off_what_slows_a_write_and_puts_it_back(app):
    before = _state(app)
    with movein.quiet(app):
        inside = {k: getattr(app, k) for k in SETTINGS}
        app.Workbooks[1].Worksheets("Other").Activate()              # the copy activates other sheets
    assert inside == {"ScreenUpdating": False, "EnableEvents": False, "DisplayAlerts": False, "Calculation": -4135}
    assert _state(app) == before and app.ActiveWorkbook is app.Workbooks[0]


@pytest.mark.parametrize("fail_at", ["ScreenUpdating", "Calculation", "body"])
def test_quiet_restores_everything_whatever_fails(app, fail_at):               # Review Focus 4
    before = _state(app)
    ran = []
    if fail_at != "body":
        app.fail_on = fail_at
    with pytest.raises((fakexl.ComError, RuntimeError)) as e:
        with movein.quiet(app):
            ran.append(True)
            app.Workbooks[1].Worksheets("Other").Activate()
            raise RuntimeError("the body failed")
    app.fail_on = None
    assert _state(app) == before
    if fail_at == "body":
        assert ran and str(e.value) == "the body failed"
    else:
        assert not ran and str(e.value) == f"cannot set {fail_at}"


def test_a_restore_that_fails_does_not_hide_the_bodys_error_or_stop_the_others(app, caplog):
    before = _state(app)
    with caplog.at_level(logging.WARNING), pytest.raises(RuntimeError, match="the body failed"):
        with movein.quiet(app):
            app.fail_on = ("EnableEvents", True)
            app.Workbooks[1].Worksheets("Other").Activate()
            raise RuntimeError("the body failed")
    app.fail_on = None
    assert _state(app) == before | {"EnableEvents": False}
    assert any("EnableEvents" in r.getMessage() for r in caplog.records)


def test_quiet_survives_an_active_sheet_that_was_deleted(app):
    before = _state(app)
    user = app.Workbooks[0]
    with movein.quiet(app):
        user.Worksheets("Q1 Answer").Activate()
        user.Worksheets("Notes").Delete()
    assert {k: getattr(app, k) for k in SETTINGS} == {k: before[k] for k in SETTINGS}


# ------------------------------------------------------------------ park, restore, drop
@pytest.fixture
def tq():
    """A workbook that holds Q3 already: its three sheets and two tables between the user's own sheets."""
    wb = fakexl.Workbook(["Q1 Answer", "Q3 Answer", "Q3 Data", "Q3 Lines b0", "Notes", "Sources"], styles_ver=False)
    _add_table(wb.Worksheets("Q3 Data"), "Data_Q3", "A4:C6")
    _add_table(wb.Worksheets("Q3 Lines b0"), "Lines_Q3_b0", "A4:B6")
    _add_table(wb.Worksheets("Notes"), "Users_table")
    wb.Worksheets("Q3 Data").put(1, 1, "Q3 Data")
    return Target(app=wb.Application, wb=wb, pid=0, hwnd=0)


def test_park_restore_and_drop(tq):
    before = fakexl.snapshot(tq.wb)
    parked = movein.park(tq, SHEETS, ["Data_Q3", "Lines_Q3_b0"])
    assert parked.sheets == {n: "old " + n for n in SHEETS}
    assert parked.tables == {"Data_Q3": "Data_Q3_old", "Lines_Q3_b0": "Lines_Q3_b0_old"}
    assert _names(tq.wb) == ["Q1 Answer", *("old " + n for n in SHEETS), "Notes", "Sources"]
    assert _tables(tq.wb) == {"Data_Q3_old": "old Q3 Data", "Lines_Q3_b0_old": "old Q3 Lines b0",
                              "Users_table": "Notes"}
    movein.restore(tq, parked)
    assert fakexl.snapshot(tq.wb) == before

    parked = movein.park(tq, SHEETS, ["Data_Q3", "Lines_Q3_b0"])
    tq.app.events.clear()
    movein.drop(tq, parked)
    assert [e[2] for e in _events(tq, "Delete")] == [["old " + n for n in SHEETS]]
    assert _names(tq.wb) == ["Q1 Answer", "Notes", "Sources"] and _tables(tq.wb) == {"Users_table": "Notes"}
    assert tq.app.DisplayAlerts is True


def test_park_takes_a_free_name_when_an_old_one_is_left_over(tq):
    tq.wb.order.append(fakexl.Worksheet(tq.wb, "old Q3 Answer"))
    _add_table(tq.wb.Worksheets("Notes"), "Data_Q3_old", "D1:E3")
    parked = movein.park(tq, ["Q3 Answer", "Q3 Data"], ["Data_Q3"])
    assert parked.sheets == {"Q3 Answer": "old Q3 Answer 2", "Q3 Data": "old Q3 Data"}
    assert parked.tables == {"Data_Q3": "Data_Q3_old2"}


def test_a_parked_name_stays_within_excels_31_characters(tq):
    tq.wb.order.append(fakexl.Worksheet(tq.wb, "Q3 Lines a_very_long_alias_x"))
    parked = movein.park(tq, ["Q3 Lines a_very_long_alias_x"], [])
    assert parked.sheets == {"Q3 Lines a_very_long_alias_x": "old Q3 Lines a_very_long_alias_"}
    assert len(parked.sheets["Q3 Lines a_very_long_alias_x"]) == 31


def test_a_park_that_fails_half_way_puts_back_what_it_renamed(tq):
    lo = tq.wb.Worksheets("Q3 Lines b0").ListObjects(1)

    class Stubborn(fakexl.ListObject):
        def __setattr__(self, k, v):
            if k == "Name" and "Name" in self.__dict__:
                raise fakexl.ComError("cannot rename the table")
            super().__setattr__(k, v)
    lo.__class__ = Stubborn
    before = fakexl.snapshot(tq.wb)
    with pytest.raises(fakexl.ComError, match="rename the table"):
        movein.park(tq, SHEETS, ["Data_Q3", "Lines_Q3_b0"])
    assert fakexl.snapshot(tq.wb) == before


def test_park_refuses_a_name_the_workbook_does_not_have_and_renames_nothing(tq):
    before = fakexl.snapshot(tq.wb)
    with pytest.raises(InputError, match="Q3 Lines b1"):
        movein.park(tq, ["Q3 Answer", "Q3 Lines b1"], ["Data_Q3"])
    with pytest.raises(InputError, match="Lines_Q3_b1"):
        movein.park(tq, ["Q3 Answer"], ["Lines_Q3_b1"])
    assert fakexl.snapshot(tq.wb) == before


def test_restore_puts_back_every_name_it_can_and_then_says_which_it_could_not(tq):
    before = fakexl.snapshot(tq.wb)
    parked = movein.park(tq, SHEETS, ["Data_Q3", "Lines_Q3_b0"])
    tq.wb.order.append(fakexl.Worksheet(tq.wb, "Q3 Data"))          # a new sheet the rollback could not remove
    with pytest.raises(HostError, match="Q3 Data") as e:
        movein.restore(tq, parked)
    assert "Q3 Answer" not in str(e.value)
    tq.wb.order.pop()
    assert {n for n in _names(tq.wb)} == {*before["order"]} - {"Q3 Data"} | {"old Q3 Data"}
    assert _tables(tq.wb) == {"Data_Q3": "old Q3 Data", "Lines_Q3_b0": "Q3 Lines b0", "Users_table": "Notes"}



# ------------------------------------------------------------------ the destination's base (Task 6)
def _destination(normal=("Telenor", 11.0), height=14.25, width=8.38, digit=8):
    wb = fakexl.Workbook(["Notes", "Sources"])
    wb.normal = normal
    ws = wb.order[0]
    ws.StandardHeight, ws.StandardWidth, ws.digit_px = height, width, digit
    return Target(app=wb.Application, wb=wb, pid=0, hwnd=0)


def test_the_base_is_the_destinations_normal_font_and_standard_sizes_in_one_call(monkeypatch):
    t = _destination()
    calls = fakexl.count_com(monkeypatch)
    assert movein.base(t) == Base("Telenor", 11, row_px=19, col_px=72, date_sep=".")   # 14.25 pt; 8.38 x 8 px + 5
    assert len(calls) == 1
    assert movein.base(_destination(("Aptos Narrow", 11.0), 15.0, 8.43, 7)) == Base("Aptos Narrow", 11, 20, 64)


def test_a_last_column_that_is_not_standard_leaves_xlsxwriters_column_unit(caplog):
    t = _destination()
    t.wb.order[0].colw[fakexl.MAXC] = 20
    with caplog.at_level(logging.WARNING, logger="smartview.movein"):
        assert movein.base(t) == Base("Telenor", 11, row_px=19, col_px=64)
    assert "standard column" in caplog.text
