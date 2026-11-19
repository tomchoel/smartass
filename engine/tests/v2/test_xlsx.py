# engine/tests/v2/test_xlsx.py
"""Phase 3 Task 2: an answer's sheets built as an .xlsx file (render/xlsx.py): the look and the bulk data, no Excel
Table, no Answer contents; what must happen after the move-in (tables, formula blocks) is returned. The file is
read back with openpyxl, and through the XML inside the zip where openpyxl cannot tell. Every figure is invented."""
import datetime
import re
import zipfile

import openpyxl
import pytest
from test_look import CASES, small_sheet
from test_v2_pipeline import FakeProvider
from test_v5_layout import GROUPED, run
from v2util import facts, resolver, spec_dict

from smartass import pipeline
from smartass.core import config
from smartass.epm.drill_lines import KEY_LABELS
from smartass.render import xlsx
from smartass.render.datasheet import header_rows, lines_header
from smartass.render.excel import SHEET_FOOTER, VALUE_FORMAT, answer_footer, answer_grid, data_grid
from smartass.render.layout import col_letter, col_number
from smartass.render.look import answer_look, style_props
from smartass.render.xlsx import DataSheet, LinesSheet, build
from smartass.xl.drill_xl import AMT_FMT, WIDTHS, header_width
from smartass.xl.vendor import tn_rapportstil as tn

FOOTER = answer_footer(datetime.datetime(2026, 11, 5))
WHEN = datetime.datetime(2026, 11, 5, 9, 30, tzinfo=datetime.timezone(datetime.timedelta(hours=1)))
SOME_SUM_KEY = "b0.m2.0"                  # D28 of the P&L table: its grand total row, TN Heltall · Hovedsum
COLUMNS = ["Period", "GL_POSTED_DATE", "ACCOUNT", "AP_CUSTOMER_NAME", "AMOUNT", "DEBET_AMOUNT", "LINE",
           "GL_DESCRIPTION_LINE"]        # DEBET_AMOUNT and LINE are not key columns
KEY_COLUMNS = [c for c in COLUMNS if c in KEY_LABELS]
SUPPLIER_COL = COLUMNS.index("AP_CUSTOMER_NAME") + 1
ACCOUNT_COL = 3                           # Cube, Query, Account, ...
POSTED = datetime.datetime(2026, 9, 30, 10, 25, tzinfo=datetime.timezone.utc)
BORDER = {1: "thin", 2: "medium", 6: "double"}
VALIGN = {"top": "top", "vcenter": "center", "bottom": None}


# ------------------------------------------------------------------ the inputs
def _plan():
    _, _, p = run([GROUPED])
    p.footer = FOOTER
    return p


def _data(account=None, grid=None) -> DataSheet:
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    head = header_rows(3, pre.spec, pre, WHEN, "tom")
    g = data_grid(pre.frame)
    if account is not None:
        g[1] = g[1][:ACCOUNT_COL - 1] + (account,) + g[1][ACCOUNT_COL:]
    return DataSheet("Q3 Data", head, grid if grid is not None else g)


def _line(i: int, supplier="Accenture AS") -> dict:
    return {"Period": "Sep 26", "GL_POSTED_DATE": POSTED, "ACCOUNT": 4100 + i, "AP_CUSTOMER_NAME": supplier,
            "AMOUNT": -1234.5 - i, "DEBET_AMOUNT": None, "LINE": i, "GL_DESCRIPTION_LINE": f"Consulting {i}"}


def _header_block(table="Lines_Q3_b0") -> list[list]:
    """The F1:K3 rectangle as pipeline._lines_header writes it: text with its leading apostrophe."""
    g = lines_header(table, {"Account": "ac.C50S", "Period": "Sep 26"}, "2026-11-05 09:30 +0100 · tom")
    return [["'" + x if isinstance(x, str) and x and not x.startswith("=") else x for x in r] for r in g]


def _lines(rows) -> LinesSheet:
    return LinesSheet(name="Q3 Lines b0", table="Lines_Q3_b0", title="Ledger lines behind ac.C50S / Sep 26",
                      source="Source: Oracle EPM Cloud, FinLite drill-through (Smart View).", columns=list(COLUMNS),
                      rows=rows, cube_value=-66_471_995.74, header_block=_header_block())


def build_with(tmp_path, fact_account=None, line_supplier=None, lines=None, plan=None, grid=None):
    rows = lines if lines is not None else [_line(i, line_supplier or "Accenture AS") for i in range(3)]
    return build(tmp_path / "answer.xlsx", "Q3 Answer", plan or _plan(), _data(fact_account, grid), [_lines(rows)])


@pytest.fixture
def plan():
    return _plan()


@pytest.fixture
def built(tmp_path, plan):
    return build(tmp_path / "answer.xlsx", "Q3 Answer", plan, _data(), [_lines([_line(i) for i in range(3)])])


def _xml(path, n: int) -> str:
    """The XML of the n-th sheet (1-based, in file order)."""
    with zipfile.ZipFile(path) as z:
        return z.read(f"xl/worksheets/sheet{n}.xml").decode("utf-8")


def _widths(ws) -> dict[str, float]:
    """Column letter -> width; openpyxl keeps a run of equal columns under its first letter."""
    return {col_letter(i): d.width for d in ws.column_dimensions.values() for i in range(d.min, d.max + 1)}


def _unmerged(path, tmp_path):
    """A copy of the file without its merges: openpyxl reads a merged cell other than the first as a blank
    MergedCell, without the format the file gives it."""
    out = tmp_path / "unmerged.xlsx"
    with zipfile.ZipFile(path) as z, zipfile.ZipFile(out, "w") as o:
        for n in z.namelist():
            data = z.read(n)
            if n.startswith("xl/worksheets/"):
                data = re.sub(rb"<mergeCells.*?</mergeCells>", b"", data)
            o.writestr(n, data)
    return out


def _rgb(colour) -> str | None:
    return "#" + colour.rgb[-6:] if colour is not None and isinstance(colour.rgb, str) else None


# ------------------------------------------------------------------ the file
def test_the_built_file_has_no_table_and_three_sheets_in_order(built):
    with zipfile.ZipFile(built.path) as z:
        assert not [n for n in z.namelist() if n.startswith("xl/tables/")]
    assert openpyxl.load_workbook(built.path).sheetnames == ["Q3 Answer", "Q3 Data", "Q3 Lines b0"] == built.sheets


def test_the_answer_sheet_carries_the_look_and_no_contents(built, plan):
    ws = openpyxl.load_workbook(built.path)["Q3 Answer"]
    assert all(c.value is None for row in ws.iter_rows() for c in row)
    r, c = plan.at(SOME_SUM_KEY)
    cell = ws.cell(r, c)
    assert (cell.font.name, cell.font.sz, cell.font.b) == ("Arial", 10, True)
    assert cell.number_format == '#,##0_);(#,##0);"–"_)'
    assert cell.border.top.style == "thin" and ws.row_dimensions[1].height == 28
    assert ws.sheet_properties.tabColor.rgb.endswith("1F3864") and ws.sheet_view.showGridLines is False
    assert ws.page_setup.orientation == "landscape" and ws.print_area is not None


def test_what_must_happen_after_the_move_in_is_returned(built, plan):
    assert [(t.sheet, t.name, t.stripes) for t in built.tables] == [("Q3 Data", "Data_Q3", False),
                                                                     ("Q3 Lines b0", "Lines_Q3_b0", True)]
    r0, c0, grid = answer_grid(plan)
    blocks = {(b.sheet, b.ref.split(":")[0]) for b in built.formulas}
    assert blocks == {("Q3 Answer", f"{col_letter(c0)}{r0}"), ("Q3 Data", "A1"), ("Q3 Lines b0", "F1")}
    answer = next(b for b in built.formulas if b.sheet == "Q3 Answer")
    assert answer.grid == grid
    assert answer.ref == f"{col_letter(c0)}{r0}:{col_letter(c0 + len(grid[0]) - 1)}{r0 + len(grid) - 1}"


@pytest.mark.parametrize("case", CASES, ids=lambda f: f.__name__)
def test_every_answer_cell_has_the_format_of_its_look(tmp_path, case):
    """Every cell of A1:<last_col><last_row> (blank ones too: map 10 item 15) carries its look as a direct
    format: font, number format, alignment, wrap, indent, borders, fill; and the widths, heights and merges."""
    p, _ = case()
    p.footer = p.footer or FOOTER
    lk = answer_look(p)
    path = build_with(tmp_path, plan=p).path
    ws = openpyxl.load_workbook(_unmerged(path, tmp_path))["Q3 Answer"]
    for (r, c), cl in lk.cells.items():
        want = style_props(cl.style)
        cell = ws.cell(r, c)
        f, a = cell.font, cell.alignment
        assert cell.value is None
        assert (f.name, f.sz, bool(f.b), bool(f.i)) == ("Arial", want["font_size"], want["bold"], want["italic"]), \
            (r, c)
        assert _rgb(f.color) == (cl.font_color or want["font_color"]), (r, c)
        assert cell.number_format == (cl.num_format or want.get("num_format") or "General"), (r, c)
        assert a.horizontal == (cl.align or want.get("align")), (r, c)
        assert a.vertical == VALIGN[cl.valign or want["valign"]], (r, c)
        assert bool(a.wrap_text) == bool(cl.wrap or want["text_wrap"]), (r, c)
        assert (a.indent or 0) == want.get("indent", 0), (r, c)
        assert bool(f.u) == bool(want.get("underline")), (r, c)
        for edge in ("top", "bottom"):
            side = getattr(cell.border, edge)
            assert side.style == BORDER.get(want.get(edge)), (r, c, edge)
            if want.get(edge):
                assert _rgb(side.color) == want[f"{edge}_color"], (r, c, edge)
        assert (_rgb(cell.fill.fgColor) if cell.fill.fill_type == "solid" else None) == want.get("bg_color"), (r, c)
    assert {r: ws.row_dimensions[r].height for r in lk.heights} == lk.heights
    assert all(ws.row_dimensions[r].customHeight for r in lk.heights)          # fixed, as RowHeight sets them
    widths = _widths(ws)
    for col, w in lk.widths.items():
        assert widths[col] == pytest.approx(w, abs=0.75), col
    ws = openpyxl.load_workbook(path)["Q3 Answer"]
    merged = sorted((m.min_row, m.min_col, m.max_col) for m in ws.merged_cells.ranges)
    assert merged == lk.merges and all(m.min_row == m.max_row for m in ws.merged_cells.ranges)


def test_merges_and_check_colours_reach_the_file(tmp_path):
    p, _ = small_sheet()
    p.footer = FOOTER
    lk = answer_look(p)
    assert len(lk.merges) >= 2 and lk.status_runs              # a wrapped title, a Key point, a check run
    ws = openpyxl.load_workbook(build_with(tmp_path, plan=p).path)["Q3 Answer"]
    assert all(ws.cell(r, c).alignment.vertical == "top" for r, c, _ in lk.merges)
    rules = {str(cf.sqref): [(x.type, x.operator, x.formula, _rgb(x.dxf.font.color), x.dxf.font.b,
                              _rgb(x.dxf.fill.bgColor) if x.dxf.fill is not None else None) for x in cf.rules]
             for cf in ws.conditional_formatting}
    want = {f"{col_letter(c)}{a}" + (f":{col_letter(c)}{b}" if b > a else ""):
            [("cellIs", "equal", ['"OK"'], "#00703C", True, None),
             ("cellIs", "notEqual", ['"OK"'], "#C00000", True, "#FDE9E9")] for c, a, b in lk.status_runs}
    assert rules == want


def test_the_answer_prints_as_today(built, plan):
    lk = answer_look(plan)
    ws = openpyxl.load_workbook(built.path)["Q3 Answer"]
    assert ws.print_area == "'Q3 Answer'!" + lk.print_area
    assert (ws.page_setup.fitToWidth or 1) == 1 and ws.page_setup.fitToHeight == 0        # 1 wide is the default
    assert ws.sheet_properties.pageSetUpPr.fitToPage
    assert ws.page_margins.left == ws.page_margins.right == pytest.approx(1.5 / 2.54)
    assert ws.print_options.horizontalCentered and not ws.print_options.gridLines
    assert lk.breaks and [b.id for b in ws.row_breaks.brk] == [r - 1 for r in lk.breaks]   # a break above row r
    xml = _xml(built.path, 1)
    assert '<oddHeader>&amp;C&amp;"Arial"&amp;8Intern</oddHeader>' in xml
    assert f"<oddFooter>&amp;C{FOOTER.replace('&', '&amp;')}</oddFooter>" in xml


# ------------------------------------------------------------------ Data
def _shown(v, value_col: bool) -> str:
    """A table cell's displayed text, as AutoFit measures it."""
    return xlsx.value_text(v) if value_col else str(v)


def test_the_data_sheet_holds_the_facts_as_values_below_an_empty_header(built):
    data = _data()
    r0, g = data.header.table_row, data.grid
    w, end = len(g[0]), r0 + len(g) - 1
    ws = openpyxl.load_workbook(built.path)["Q3 Data"]
    assert all(ws.cell(r, c).value is None for r in range(1, r0) for c in range(1, 5))     # the header: a block
    for i, row in enumerate(g):
        for j, v in enumerate(row):
            cell = ws.cell(r0 + i, j + 1)
            assert (cell.font.name, cell.font.sz) == ("Arial", 9)
            if i == 0:
                assert cell.value == v and cell.font.b and cell.number_format == ("@" if j < w - 1 else "General")
            elif j < w - 1:
                assert cell.value == (v or None) and cell.number_format == "@" and not cell.font.b
                assert v == "" or cell.data_type == "s"
            else:
                assert cell.value == v and isinstance(cell.value, (int, float)) and cell.number_format == VALUE_FORMAT
                assert _rgb(cell.font.color) == "#0000C0" and cell.alignment.horizontal == "right"
    assert ws.freeze_panes == f"A{r0 + 1}" and ws.sheet_view.showGridLines is False
    assert ws.sheet_properties.tabColor.rgb.endswith("A6A6A6")
    for j in range(w):
        fit = min(max(tn.text_units(_shown(row[j], i > 0 and j == w - 1), 9)
                      for i, row in enumerate(g) if row[j] not in (None, "")), 40)
        want = max(fit, 16) if j < 4 else fit
        assert _widths(ws)[col_letter(j + 1)] == pytest.approx(want, abs=0.75), col_letter(j + 1)
    assert ws.print_area == f"'Q3 Data'!$A$1:${col_letter(max(4, w))}${end}"
    assert ws.page_setup.orientation == "landscape" and ws.sheet_properties.pageSetUpPr.fitToPage
    assert ws.page_margins.left == pytest.approx(1.5 / 2.54) and ws.print_options.horizontalCentered
    xml = _xml(built.path, 2)
    assert f"<oddFooter>&amp;C{SHEET_FOOTER.replace('&', '&amp;')}</oddFooter>" in xml
    assert '<oddHeader>&amp;C&amp;"Arial"&amp;8Intern</oddHeader>' in xml
    assert built.tables[0].ref == f"A{r0}:{col_letter(w)}{end}"


def test_a_data_table_without_facts_keeps_one_empty_body_row(tmp_path):
    data = _data()
    built = build_with(tmp_path, grid=data.grid[:1])
    r0 = data.header.table_row
    assert built.tables[0].ref == f"A{r0}:{col_letter(len(data.grid[0]))}{r0 + 1}"


def test_value_text_is_the_number_as_the_value_format_shows_it():
    assert [xlsx.value_text(v) for v in (1234567.4, -1234.6, 0.0, 0.4, None)] == ["1,234,567", "(1,235)", "–", "0",
                                                                                  ""]


def test_the_data_header_has_its_look_in_the_file_and_its_contents_as_a_formula_block(built):
    data = _data()
    h = len(data.header)
    ws = openpyxl.load_workbook(built.path)["Q3 Data"]
    numeric = [i + 1 for i, (lab, _) in enumerate(data.header) if lab in ("Control total", "Tolerance (NOK)")]
    assert len(numeric) == 2
    for r in range(1, h + 1):
        for c in range(1, 5):
            f = ws.cell(r, c).font
            assert f.name == "Arial"
            if (r, c) == (1, 1):
                assert (f.sz, f.b, _rgb(f.color)) == (14, True, "#1F3864")
            else:
                assert (f.sz, bool(f.b)) == (9, False)
                assert _rgb(f.color) == ("#767676" if c == 1 else "#262626"), (r, c)
            fmt = VALUE_FORMAT if r in numeric and c in (2, 3) else "General"
            assert ws.cell(r, c).number_format == fmt, (r, c)                # never "@": a formula stays a formula
    block = next(b for b in built.formulas if b.sheet == "Q3 Data")
    assert block.ref == f"A1:D{h}" and len(block.grid) == h and all(len(r) == 4 for r in block.grid)
    labels = dict(data.header)
    assert block.grid[0] == ["'" + labels["Title"], "", "", ""]
    assert block.grid[1] == ["'Question", "'" + labels["Question"], "", ""]
    static, live, status = labels["Control total"]
    assert block.grid[numeric[0] - 1] == ["'Control total", static, live, status]
    assert block.grid[numeric[1] - 1] == ["'Tolerance (NOK)", 1.0, "", ""]
    rows = next(i for i, (lab, _) in enumerate(data.header) if lab == "Rows")
    assert block.grid[rows][1] == labels["Rows"] and labels["Rows"].startswith("=COUNTIFS(")


@pytest.mark.parametrize("text", ["=SUM(A1)", "+47 900 00 000", "-", "@home", "0047", "1-2", "3/4", "1E5", "TRUE"])
def test_text_stays_text(tmp_path, text):                                 # Review Focus 1
    data = _data()
    built = build_with(tmp_path, fact_account=text, line_supplier=text)
    wb = openpyxl.load_workbook(built.path)
    fact, line = wb["Q3 Data"].cell(data.header.table_row + 1, ACCOUNT_COL), wb["Q3 Lines b0"].cell(5, SUPPLIER_COL)
    assert fact.value == text and fact.data_type == "s"
    assert line.value == text and line.data_type == "s"
    with zipfile.ZipFile(built.path) as z:                                   # no formula anywhere in the file
        assert not [n for n in z.namelist() if n.startswith("xl/worksheets/") and b"<f>" in z.read(n)]


# ------------------------------------------------------------------ Lines
def test_the_lines_sheet_as_today(built):
    ws = openpyxl.load_workbook(built.path)["Q3 Lines b0"]
    ls = _lines([])
    assert (ws["A1"].value, ws["A1"].font.b) == (ls.title, True)
    assert ws["A2"].value == ls.source and _rgb(ws["A2"].font.color) == "#000000"
    assert ws["C3"].value == "Cube value" and ws["D3"].value == -66_471_995.74 and ws["D3"].number_format == AMT_FMT
    assert [ws.cell(4, j + 1).value for j in range(len(COLUMNS))] == COLUMNS
    for i in range(3):
        line = _line(i)
        got = {c: ws.cell(5 + i, j + 1) for j, c in enumerate(COLUMNS)}
        assert got["Period"].value == "Sep 26" and got["Period"].data_type == "s"
        assert got["GL_POSTED_DATE"].value == POSTED.replace(tzinfo=None)                 # a serial, tz dropped
        assert got["GL_POSTED_DATE"].number_format == "dd.mm.yyyy"
        assert got["ACCOUNT"].value == str(line["ACCOUNT"]) and got["ACCOUNT"].data_type == "s"   # TEXT_LABELS
        assert got["LINE"].value == str(i) and got["LINE"].data_type == "s"
        assert got["AMOUNT"].value == line["AMOUNT"] and got["AMOUNT"].number_format == AMT_FMT
        assert got["DEBET_AMOUNT"].value is None and got["DEBET_AMOUNT"].number_format == AMT_FMT
        assert got["GL_DESCRIPTION_LINE"].number_format == "General"
    for j, c in enumerate(COLUMNS):
        want = max(WIDTHS.get(c, 14), header_width(c))
        assert _widths(ws)[col_letter(j + 1)] == pytest.approx(want, abs=0.75), c
    assert ws.freeze_panes == "A5" and ws.print_title_rows == "$4:$4"
    assert ws.sheet_properties.tabColor.rgb.endswith("A6A6A6") and ws.sheet_view.showGridLines is not False
    assert ws.print_area == f"'Q3 Lines b0'!$A$1:$K${4 + 3}"                          # the table and F1:K3
    setup = ws.sheet_properties.pageSetUpPr
    assert (setup is None or not setup.fitToPage) and ws.page_setup.orientation == "landscape"   # 100 %
    assert f"<oddFooter>&amp;C{SHEET_FOOTER.replace('&', '&amp;')}</oddFooter>" in _xml(built.path, 3)
    for r in range(1, 4):
        for c in range(col_number("F"), col_number("K") + 1):
            cell = ws.cell(r, c)
            assert cell.value is None and (cell.font.name, cell.font.sz) == ("Arial", 9)
            grey = c == col_number("F") or (r, c) in ((3, col_number("H")), (3, col_number("J")))
            assert _rgb(cell.font.color) == ("#767676" if grey else "#262626"), (r, c)
            assert cell.number_format == (VALUE_FORMAT if (r, c) == (3, col_number("G")) else "General")
    block = next(b for b in built.formulas if b.sheet == "Q3 Lines b0")
    assert block.ref == "F1:K3" and block.grid == _header_block()
    assert built.tables[1].ref == f"A4:{col_letter(len(COLUMNS))}7"


@pytest.mark.parametrize("n, cols", [(0, "all"), (50, "all"), (51, "key")])
def test_lines_at_the_edges(tmp_path, n, cols, monkeypatch):               # Review Focus 5
    monkeypatch.setattr(config, "DRILL_XL_FULL", 50)                       # keep the test fast
    built = build_with(tmp_path, lines=[_line(i) for i in range(n)])
    want = COLUMNS if cols == "all" else KEY_COLUMNS
    ws = openpyxl.load_workbook(built.path)["Q3 Lines b0"]
    assert [ws.cell(4, j + 1).value for j in range(len(COLUMNS))] == want + [None] * (len(COLUMNS) - len(want))
    # 0 lines: the header row plus one empty body row (ListObjects.Add needs a body row)
    assert built.tables[1].ref == f"A4:{col_letter(len(want))}{4 + max(n, 1)}"
    assert all(ws.cell(5 + n, j + 1).value is None for j in range(len(COLUMNS)))      # nothing below the lines
    if n:
        last = {c: ws.cell(4 + n, j + 1).value for j, c in enumerate(want)}
        assert last["GL_DESCRIPTION_LINE"] == f"Consulting {n - 1}" and last["AMOUNT"] == -1234.5 - (n - 1)
    else:
        assert all(ws.cell(5, j + 1).value is None for j in range(len(want)))


class FakeRows:
    """As many rows as len() says, none of which may be read: the refusal comes before any row is touched."""

    def __init__(self, n: int):
        self.n = n

    def __len__(self):
        return self.n

    def __iter__(self):
        raise AssertionError("the rows were read")

    def __getitem__(self, i):
        raise AssertionError("the rows were read")


def test_more_lines_than_excel_has_rows_is_refused(tmp_path):
    with pytest.raises(xlsx.TooManyRows, match="1 048 576"):
        build_with(tmp_path, lines=FakeRows(1_048_573))                    # 4 header rows + lines > the sheet
    assert not (tmp_path / "answer.xlsx").exists()


def test_lines_that_end_on_the_last_row_of_the_sheet_are_not_refused():
    xlsx.check_rows("Q3 Lines b0", 4 + 1_048_572)                         # row 1 048 576: no error
    with pytest.raises(xlsx.TooManyRows, match="Q3 Lines b0"):
        xlsx.check_rows("Q3 Lines b0", 4 + 1_048_573)


def test_more_facts_than_excel_has_rows_is_refused(tmp_path):
    r0 = _data().header.table_row
    with pytest.raises(xlsx.TooManyRows, match="1 048 576"):
        build_with(tmp_path, grid=FakeRows(1_048_576 - r0 + 2))           # the column row + facts one past the end
    assert not (tmp_path / "answer.xlsx").exists()


def test_the_answer_name_must_carry_its_question_number(tmp_path):
    with pytest.raises(ValueError, match="Q<n> Answer"):
        build(tmp_path / "x.xlsx", "Answer", _plan(), _data(), [])
    assert not (tmp_path / "x.xlsx").exists()


def test_an_answer_without_lines_has_two_sheets_and_one_table(tmp_path):
    data = _data()
    built = build(tmp_path / "a.xlsx", "Q7 Answer", _plan(), DataSheet("Q7 Data", data.header, data.grid), [])
    assert built.sheets == ["Q7 Answer", "Q7 Data"] and [(t.sheet, t.name) for t in built.tables] == [("Q7 Data",
                                                                                                         "Data_Q7")]
    assert re.fullmatch(r"A\d+:[A-Z]+\d+", built.tables[0].ref)
    assert openpyxl.load_workbook(built.path).sheetnames == built.sheets
