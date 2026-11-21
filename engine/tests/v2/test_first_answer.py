# engine/tests/v2/test_first_answer.py
"""Phase 3c Task 1: the first answer of a new workbook. The renderer writes the whole workbook - the answer's sheets,
the furnished Sources log and the sensitivity label - to its reserved name in the analysis folder; Excel only opens
it, and the job goes on with the look, the bring-in (no park, no insert), the close-out and the save. Nothing is
added from a template, saved as, inserted or furnished in Excel. A failure once the file is reserved closes the
workbook without saving and deletes the file. All on the fake Excel (the COM executor); every figure is invented."""
from __future__ import annotations

import dataclasses
from pathlib import Path
from types import SimpleNamespace

import pytest
from xlsxwriter.theme import THEME_XML_2023
import test_file_job
from test_file_job import BASE, _batches, _job, _pre

import fakexl
from smartass import exceljob, pipeline
from smartass.core import config, session
from smartass.core.timing import BATCH_PREFIX
from smartass.render import xlsx
from smartass.xl import jobs, movein, sources, styling
from smartass.xl import target as xltarget
from smartass.xl.host import HostError
from smartass.xl.inproc import runner
from smartass.xl.inproc.runner import BatchError

tmp_home, calculated = test_file_job.tmp_home, test_file_job.calculated      # the fixtures, shared
FURNITURE = ("AutoFilter", "ColumnWidth", "PageSetup", "PrintCommunication", "FreezePanes", "SplitRow", "Names.Add",
             'Range("A:N")', f"Columns({sources.LIMITS_COL}).Font", f"Columns({sources.TOTAL_COL}).NumberFormat")
#               what sources._furnish and its print setup write


@pytest.fixture
def new_book(monkeypatch, tmp_home, tmp_path):
    """The user's Excel with their own workbook open and no analysis workbook yet: target 'new' chooses it, and
    the job builds the new workbook in the analysis folder. The base of a new workbook is known (a look of one made
    from the TN template stored it)."""
    app = fakexl.Application()
    mine = fakexl.Workbook(["Mine"], app=app, styles_ver=False)
    mine.Name, mine.FullName = "mine.xlsx", r"C:\Users\me\mine.xlsx"
    folder = tmp_path / "analyser"
    folder.mkdir()
    monkeypatch.setattr(session, "project_dir", lambda: folder)
    sessions: list = []
    monkeypatch.setattr(session, "set_workbook", sessions.append)
    monkeypatch.setattr(config, "LABEL_ID", "f604d2c9")
    monkeypatch.setattr(config, "LABEL_SITE", "1676489c")
    monkeypatch.setattr(jobs, "hold_turn", lambda *a, **k: True)
    monkeypatch.setattr(styling, "template_theme", lambda: THEME_XML_2023)
    xltarget.remember_new_base(BASE)
    asked: list = []

    def target(h, target, title, build_new=False):
        asked.append(build_new)
        return xltarget._new_target(None, title, [(7, 70, app)], build_new=build_new)
    monkeypatch.setattr(exceljob, "_target", target)
    monkeypatch.setattr(exceljob, "remember", lambda pre, out: None)
    return SimpleNamespace(app=app, mine=mine, folder=folder, sessions=sessions, asked=asked)


def _runs(monkeypatch) -> list:
    runs, real = [], runner.run
    monkeypatch.setattr(runner, "run", lambda app, b, *a, **k: runs.append(b) or real(app, b, *a, **k))
    return runs


def _writes(b) -> list[str]:
    """The texts of a batch's sets and calls (its reads left out)."""
    d = b.to_dict()
    return [b.text(i) for i, op in enumerate(d["ops"] + d["finally"]) if op["k"] in ("set", "call")]


def _furniture(runs) -> list[str]:
    return [x for b in runs for x in _writes(b) if any(f in x for f in FURNITURE)]


def _made(folder) -> list[str]:
    return sorted(p.name for p in folder.iterdir())


def test_the_first_answer_is_its_built_file_opened_in_four_batches(new_book, calculated, monkeypatch):
    """No Workbooks.Add, no SaveAs, no insert, no Sources furniture in Excel: the file holds it all. COM calls: the
    open workbooks' names (the file name), the Excel's International (the base's date format), the Open and its
    read-only check (each waits for a busy Excel on its own), then one per batch - look, bring in (the finish only),
    close out, save."""
    calls = fakexl.count_com(monkeypatch)
    runs = _runs(monkeypatch)
    out, timings = _job(_pre(), calculated)
    app, folder = new_book.app, new_book.folder
    path = Path(out["workbook"])
    wb = app.Workbooks(path.name)
    assert new_book.asked == [True] and path.parent == folder and _made(folder) == [path.name]
    assert out["q"] == 1 and out["checks"]["sheet_equals_model"] == "OK" and out["sheets"] == ["Q1 Answer", "Q1 Data"]
    assert [ws.Name for ws in wb.order] == ["Q1 Answer", "Q1 Data", "Sources", "_svmeta"]
    assert [lo.Name for ws in wb.order for lo in ws.ListObjects] == ["Data_Q1"]
    kinds = [k for _, k, *_ in app.events]
    assert kinds.count("Open") == 1 and not {"Sheets.Add", "SaveAs", "Add"} & set(kinds)
    assert _batches(timings) == {"draw.look": 1, "draw.close out": 1, "save": 1} and timings[BATCH_PREFIX + "draw"] == 3
    assert len(calls) == 8 and len(runs) == 4
    assert "draw.build" in timings and "draw.open" in timings
    assert _furniture(runs) == []
    assert wb.saved == 1 and new_book.sessions == [str(path)]
    log = wb.Worksheets("Sources")
    assert log.Range("A2").Value and log.Range("B2").Value == "Q1" and log.Range("A3").Value is None
    assert app.ActiveWorkbook is wb and wb.active.Name == "Q1 Answer"


def test_a_first_answer_with_a_lines_sheet_builds_it_into_the_new_workbook(new_book, calculated, monkeypatch):
    """The Lines sheet in the built workbook, before the log; its table made after the open; its Sources row and
    link; its dates in the chosen Excel's date format (nb-NO: the built-in short date)."""
    from test_file_job import SUPPLIERS
    from test_v2_pipeline import LINE_CELL, lines_only_spec

    from smartass.model.lines import LineSet
    seen = {}
    real = xlsx.build

    def build(path, *a, **k):
        seen.update(k)
        return real(path, *a, **k)
    monkeypatch.setattr(xlsx, "build", build)
    ls = LineSet("b0", list(SUPPLIERS), ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=-100.0, cell=LINE_CELL)
    out, _ = _job(_pre(lines_only_spec(), lines={"b0": ls}), calculated)
    wb = new_book.app.Workbooks(Path(out["workbook"]).name)
    assert out["checks"]["sheet_equals_model"] == "OK" and seen["base"].date_format == movein.SHORT_DATE
    assert [ws.Name for ws in wb.order] == ["Q1 Answer", "Q1 Data", "Q1 Lines b0", "Sources", "_svmeta"]
    assert [lo.Name for ws in wb.order for lo in ws.ListObjects] == ["Data_Q1", "Lines_Q1_b0"]
    log = wb.Worksheets("Sources")
    assert [log.Range(f"C{r}").Value for r in (2, 3)] == ["b0", None]            # the drill's row (no query)
    assert "#'Q1 Lines b0'!A1" in str(log.Range("M2").Value)


def test_a_second_answer_finds_the_new_workbooks_log_furnished(new_book, calculated, monkeypatch):
    """The furniture's names came with the file: the next answer into the workbook appends its rows only."""
    out, _ = _job(_pre(), calculated)
    wb = new_book.app.Workbooks(Path(out["workbook"]).name)
    t = xltarget.Target(app=new_book.app, wb=wb, pid=7, hwnd=70, created=True)
    monkeypatch.setattr(exceljob, "_target", lambda h, target, title, build_new=False: t)
    runs = _runs(monkeypatch)
    out2, _ = _job(_pre(), calculated)
    assert out2["q"] == 2 and out2["checks"]["sheet_equals_model"] == "OK"
    assert _furniture(runs) == []
    assert wb.Worksheets("Sources").Range("B3").Value == "Q2"


def test_a_failure_after_the_open_closes_the_new_workbook_unsaved_and_deletes_its_file(new_book, calculated,
                                                                                        monkeypatch):
    def refuse(self, *a):
        raise fakexl.ComError("a table was refused")
    monkeypatch.setattr(fakexl._ListObjects, "Add", refuse)
    with pytest.raises(BatchError, match="a table was refused.*the new workbook .*xlsx was closed without saving and "
                                         "its file deleted"):
        _job(_pre(), calculated)
    app = new_book.app
    assert list(app.Workbooks) == [new_book.mine] and app.ActiveWorkbook is new_book.mine
    assert _made(new_book.folder) == [] and new_book.sessions == []
    assert (app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation) == (True, True, True, -4105)


def test_a_failed_save_deletes_the_new_workbook_too(new_book, calculated, monkeypatch):
    def save(self):
        raise fakexl.ComError("the save was cancelled")
    monkeypatch.setattr(fakexl.Workbook, "Save", save)
    with pytest.raises(BatchError, match="cancelled.*closed without saving and its file deleted"):
        _job(_pre(), calculated)
    assert _made(new_book.folder) == [] and list(new_book.app.Workbooks) == [new_book.mine]


@pytest.mark.parametrize("route", ["built", "added and saved as", "no Excel"])
def test_every_route_of_a_new_workbook_marks_its_target_new(new_book, monkeypatch, route):
    """Final review I1: the job saves a workbook it created in the foreground - the built file, one added from the
    template and saved as, and one the hidden Excel made when no Excel was open; the session's workbook is not new."""
    if route == "no Excel":
        made = new_book.folder / "made.xlsx"
        monkeypatch.setattr(xltarget, "_new_analysis_file", lambda h, title: made)
        monkeypatch.setattr(xltarget, "_open_path", lambda h, path, excels, *a: xltarget.Target(
            app=new_book.app, wb=None, pid=8, hwnd=0))
        t = xltarget._new_target(None, "T", [])
    else:
        monkeypatch.setattr(xltarget, "_save_new", lambda app, wb, title: new_book.folder / "added.xlsx")
        monkeypatch.setattr(fakexl._Workbooks, "Add", lambda self, *a: fakexl.Workbook(
            ["Sources"], app=new_book.app, styles_ver=False), raising=False)
        t = xltarget._new_target(None, "T", [(7, 70, new_book.app)], build_new=route == "built")
    assert t.new is True and t.created is True and (t.made is not None) == (route == "built")
    assert xltarget.Target(app=None, wb=None, pid=0, hwnd=0, created=True).new is False


def test_a_build_that_fails_deletes_the_reserved_file_and_opens_nothing(new_book, calculated, monkeypatch):
    def too_many(path, *a, **k):
        path.write_bytes(b"PK half a file")
        raise xlsx.TooManyRows("sheet 'Q1 Data' would need 1 100 000 rows; an Excel sheet has 1 048 576")
    monkeypatch.setattr(xlsx, "build", too_many)
    with pytest.raises(pipeline.GuardError, match="1 100 000 rows") as e:
        _job(_pre(), calculated)
    assert "narrow" in e.value.fix and "was not opened; its file was deleted" in str(e.value)
    assert _made(new_book.folder) == [] and list(new_book.app.Workbooks) == [new_book.mine]
    assert not [k for _, k, *_ in new_book.app.events if k == "Open"]


def test_a_failed_open_deletes_the_reserved_file(new_book, calculated, monkeypatch):
    def refuse(self, path, *a):
        raise fakexl.ComError("Open method of Workbooks class failed")
    monkeypatch.setattr(fakexl._Workbooks, "Open", refuse)
    with pytest.raises(fakexl.ComError, match="Open method"):
        _job(_pre(), calculated)
    assert _made(new_book.folder) == [] and new_book.sessions == []


def test_the_built_file_carries_the_label_and_the_new_workbooks_base(new_book, calculated, monkeypatch):
    import zipfile
    seen = {}
    real = xlsx.build

    def build(path, *a, **k):
        seen.update(k)
        built = real(path, *a, **k)
        with zipfile.ZipFile(path) as z:
            seen["custom"] = z.read("docProps/custom.xml").decode()
        return built
    monkeypatch.setattr(xlsx, "build", build)
    _job(_pre(), calculated)
    assert seen["base"] == xlsx.Base("Telenor", 11, 19, 72, movein.SHORT_DATE)      # nb-NO: the built-in short date
    assert seen["book"].properties.keys() == xltarget.label_properties().keys()
    assert seen["book"].theme == THEME_XML_2023
    assert 'name="MSIP_Label_f604d2c9_Enabled"><vt:lpwstr>true<' in seen["custom"]


@pytest.mark.parametrize("label, base, theme, want", [("f604", BASE, True, True), ("", BASE, True, False),
                                                      ("site", BASE, True, False),
                                                      ("f604", None, True, False), ("f604", BASE, False, False),
                                                      ("f604", dataclasses.replace(BASE, col_px=73), True, False)])
def test_the_job_builds_a_new_workbook_only_with_a_label_a_known_base_and_the_templates_theme(tmp_home, monkeypatch,
                                                                                              label, base, theme, want):
    """Without our label id and site the file could not be saved; without a stored base (or with a standard column XlsxWriter
    has no digit width for) the file would not measure as the template's workbooks do, and without the template's
    theme its links and table styles would not look as theirs: then a new workbook is made today's way (its look
    stores the base)."""
    monkeypatch.setattr(config, "LABEL_ID", "" if label == "" else "f604")
    monkeypatch.setattr(config, "LABEL_SITE", "" if label == "site" else "1676")
    monkeypatch.setattr(styling, "template_theme", lambda: THEME_XML_2023 if theme else None)
    if base is not None:
        xltarget.remember_new_base(base)
    asked = []

    def target(h, target, title, build_new=False):
        asked.append(build_new)
        raise HostError("stop here")
    monkeypatch.setattr(exceljob, "_target", target)
    with pytest.raises(HostError, match="stop here"):
        exceljob.excel_job(None, _pre())
    assert asked == [want]


def test_a_new_workbook_made_from_the_template_stores_the_base_its_look_found(tmp_home, calculated, monkeypatch):
    wb = fakexl.Workbook(["Sources"])
    wb.Name, wb.FullName = "2026-11-20_0930 answer.xlsx", str(tmp_home / "answer.xlsx")
    wb.normal = ("Telenor", 11.0)
    t = xltarget.Target(app=wb.Application, wb=wb, pid=0, hwnd=0, created=True, from_template=True)
    monkeypatch.setattr(exceljob, "_target", lambda h, target, title, build_new=False: t)
    monkeypatch.setattr(exceljob, "remember", lambda pre, out: None)
    assert xltarget.new_base() is None
    _job(_pre(), calculated)
    got = xltarget.new_base()
    assert (got.font_name, got.font_size) == ("Telenor", 11) and got.row_px == 20


def test_a_stored_base_of_another_template_version_is_not_used(tmp_home, monkeypatch):
    xltarget.remember_new_base(BASE)
    assert xltarget.new_base() == BASE
    monkeypatch.setattr(styling, "_styles_version", lambda: "other")
    assert xltarget.new_base() is None



def test_a_built_workbook_whose_base_is_not_the_stored_one_drops_it_and_keeps_the_answer(new_book, calculated,
                                                                                         monkeypatch, caplog):
    """Review minor 3: the look of a built workbook reads the base Excel gives it; when that is not the stored one
    (a default font changed, a template the stored base no longer describes) the answer stays, a warning is logged and
    the stored base is dropped - the next new workbook is made from the template once and stores a fresh one."""
    import logging
    real = exceljob._look

    def look(job):
        found = real(job)
        found.base = dataclasses.replace(found.base, col_px=64)
        return found
    monkeypatch.setattr(exceljob, "_look", look)
    with caplog.at_level(logging.WARNING):
        out, _ = _job(_pre(), calculated)
    assert out["checks"]["sheet_equals_model"] == "OK" and new_book.sessions == [out["workbook"]]
    assert xltarget.new_base() is None
    assert any("stored base" in r.getMessage() for r in caplog.records)


def test_a_built_workbook_whose_base_is_the_stored_one_keeps_it(new_book, calculated):
    _job(_pre(), calculated)
    assert xltarget.new_base() == BASE


def test_the_templates_theme_is_read_once_per_version_of_the_file(tmp_home, monkeypatch):
    """Review minor 7: _new_book runs before every answer; the template's zip is read again only when the file
    changes (its path or its time)."""
    import os
    import zipfile

    from xlsxwriter.theme import THEME_XML_2007
    monkeypatch.setattr(config, "SVKIT", tmp_home)
    path = styling.template_path()

    def put(theme, ns):
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("xl/theme/theme1.xml", theme)
        os.utime(path, ns=(ns, ns))
    put(THEME_XML_2023, 1_700_000_000_000_000_000)
    opened, real = [], zipfile.ZipFile
    monkeypatch.setattr(zipfile, "ZipFile", lambda *a, **k: opened.append(a[0]) or real(*a, **k))
    assert styling.template_theme() == THEME_XML_2023 == styling.template_theme() and len(opened) == 1
    monkeypatch.setattr(zipfile, "ZipFile", real)
    put(THEME_XML_2007, 1_700_000_001_000_000_000)
    assert styling.template_theme() == THEME_XML_2007
