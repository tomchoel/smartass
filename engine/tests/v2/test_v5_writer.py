# engine/tests/v2/test_v5_writer.py
"""Plan 5 Task 8: the answer writer - the Sheet = model status cell after the readback, the print setup and the
column widths (fakes only, no Excel). Every figure is invented."""
import datetime

from fakexl import Workbook
from test_v5_layout import GROUPED, run

from smartass.render import excel
from smartass.render.excel import answer_footer, answer_grid, answer_roles, sheet_status, write_status
from smartass.render.layout import SheetPlan
from smartass.xl import styling


def _ws():
    return Workbook(["Q3 Answer"]).order[0]


def test_status_cell_ok_without_mismatches_and_diff_otherwise():
    ws = _ws()
    p = SheetPlan(status_cell="D30")
    assert write_status(ws, p, []) == "OK"
    assert ws.cells[(30, 4)]["Value"] == "OK"
    assert write_status(ws, p, ["D5: sheet 1, model 2", "E5: sheet 1, model 3"]) == "DIFF in 2 cells"
    assert ws.cells[(30, 4)]["Value"] == "DIFF in 2 cells"
    assert sheet_status([]) == "OK" and sheet_status(["x"]) == "DIFF in 1 cells"


def test_status_cell_is_the_layouts_and_coloured_like_a_check():
    _, _, p = run([GROUPED])
    r = int(p.status_cell[1:])
    assert p.cells[(r, 3)] == "Sheet = model" and p.status_cell.startswith("D")
    row = next(x for x in p.rows if x["row"] == r)
    assert row["kind"] == "check" and row["status_col"] == "D"


def test_print_setup_properties():
    ws = _ws()
    styling.print_setup(ws, "G", 14, "Source: Oracle EPM Cloud · fetched 2026-11-05 · page &P of &N")
    ps = ws.PageSetup
    assert ps.PrintArea == "$A$1:$G$14"
    assert (ps.Orientation, ps.Zoom, ps.FitToPagesWide, ps.FitToPagesTall) == (2, False, 1, False)
    assert ps.PrintTitleRows == "" and ps.PrintGridlines is False
    assert ps.CenterFooter == "Source: Oracle EPM Cloud · fetched 2026-11-05 · page &P of &N"
    assert ps.LeftFooter == "" and ps.RightFooter == ""            # one page number, not v1's 'Side &P av &N'
    assert ws.Parent.Application.PrintCommunication is True       # batched and switched back on


def test_print_setup_is_skipped_when_unchanged_and_redone_for_a_new_footer():
    ws = _ws()
    styling.print_setup(ws, "G", 14, "F1")
    ws.PageSetup.Orientation = 1                                  # changed by hand
    styling.print_setup(ws, "G", 14, "F1")
    assert ws.PageSetup.Orientation == 1                          # same key: one call, skipped
    styling.print_setup(ws, "G", 14, "F2")
    assert ws.PageSetup.Orientation == 2 and ws.PageSetup.CenterFooter == "F2"


def test_footer_text():
    when = datetime.datetime(2026, 11, 5, 9, 30)
    assert answer_footer(when) == "Source: Oracle EPM Cloud · fetched 2026-11-05 · page &P of &N"


def test_a_laid_out_answer_prints_its_tables_without_helpers_and_lints_clean():
    _, _, p = run([GROUPED])
    p.footer = answer_footer(datetime.datetime(2026, 11, 5))
    roles = answer_roles(p)
    assert roles["print_footer"] == p.footer
    ws = _ws()
    r0, c0, g = answer_grid(p)
    for i, row in enumerate(g):
        for j, v in enumerate(row):
            ws.put(r0 + i, c0 + j, v)
    assert styling.apply_report(ws, roles) == []
    ps = ws.PageSetup
    assert ps.PrintArea == f"$A$1:${p.print_last_col}${p.last_row}" and ps.CenterFooter == p.footer
    assert ps.PrintTitleRows == "" and ps.RightFooter == ""


def test_without_a_footer_the_v1_page_setup_stays():
    _, _, p = run([GROUPED])
    roles = answer_roles(p)
    assert "print_footer" not in roles


def test_widths_b_is_9_and_gap_columns_2():
    _, _, p = run([GROUPED])
    assert p.widths["A"] == 2 and p.widths["B"] == 9
    gaps = [k for k, w in p.widths.items() if w == 2 and k != "A"]
    assert len(gaps) == 1                                         # month | YTD: one gap column, 2 wide
    lines = [x["row"] for x in p.rows if x.get("kind") == "line" and x.get("num", {}).get("E")]
    assert lines and all((r, excel.col_number(gaps[0])) not in p.cells for r in lines)   # gaps hold no cells


def test_the_answer_roles_pass_widths_and_code_cols():
    _, _, p = run([GROUPED])
    roles = answer_roles(p)
    assert roles["code_cols"] == ["B"]
    assert all(roles["col_widths"][k] == w for k, w in p.widths.items())


def test_the_job_closes_out_with_the_status_after_the_readback_and_sets_the_footer(monkeypatch, tmp_path):
    """The read-back's mismatches go to the close-out batch, which writes the status (write_status)."""
    from test_file_job import Wired
    from test_v2_pipeline import FakeProvider
    from v2util import facts, resolver, spec_dict

    from smartass import exceljob, pipeline
    from smartass.core import config
    from smartass.render import xlsx
    monkeypatch.setattr(config, "HOME", tmp_path)
    wired, seen = Wired(monkeypatch), {}
    log = wired.calls
    real = xlsx.build

    def build(path, name, plan, *a, **k):
        seen.update(plan=plan)
        return real(path, name, plan, *a, **k)
    monkeypatch.setattr(xlsx, "build", build)
    monkeypatch.setattr(exceljob, "compare", lambda *a: log.append("compare") or ["D5: sheet 1, model 2"])
    monkeypatch.setattr(exceljob, "_close_out", wired.rec("close_out", lambda job, built, plan, pre, header, stamp,
                                                          read, m, book: seen.update(m=m) or (None, None, [])))
    out = exceljob.excel_job(None, pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts())), "session")
    assert log.index("bring_in") + 1 == log.index("compare") and log.index("compare") + 1 == log.index("close_out")
    assert seen["m"] == ["D5: sheet 1, model 2"] and out["checks"]["sheet_equals_model"] == "DIFF in 1 cells"
    assert seen["plan"].footer.startswith("Source: Oracle EPM Cloud · fetched ")


def test_v1_page_setup_keeps_a_v2_footer_and_adds_no_second_page_number():
    """xl_tidy on a v2 answer (no print_footer role): v1's 'Side &P av &N' would sit beside 'page &P of &N'."""
    _, _, p = run([GROUPED])
    p.footer = answer_footer(datetime.datetime(2026, 11, 5))
    ws = _ws()
    styling.print_setup(ws, "G", 14, p.footer)
    styling._page_setup(ws, ws.Parent.Application, 14, 7, {})
    ps = ws.PageSetup
    assert ps.CenterFooter == p.footer and ps.RightFooter == "" and ps.LeftFooter == ""
    fresh = _ws()                                                 # a v1 sheet without a centre footer: as before
    styling._page_setup(fresh, fresh.Parent.Application, 14, 7, {})
    assert fresh.PageSetup.RightFooter == '&"Arial"&8Side &P av &N'
