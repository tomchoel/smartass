# Light and fast, phase 3: file-built answers - Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** An answer reaches the user's Excel through about twenty COM calls instead of about 280: its sheets are built as an .xlsx file, moved into the open workbook, and finished there.

**Architecture:** `render/layout.py` keeps describing the contents. A new pure `render/look.py` gives the look of every Answer cell (what `xl/styling.apply_report` applies through COM today). `render/xlsx.py` writes the look, the charts and the bulk data to a temporary file with XlsxWriter and returns what must be done after the move-in (tables, formula blocks). `xl/movein.py` moves the sheets in with one copy call and finishes them. The new job lives beside the old one behind `config.RENDER` until a live comparison and the speed gate have passed; then the old writers are deleted.

**Tech Stack:** Python 3.12, XlsxWriter (new dependency), pywin32 COM, pytest with `engine/tests/fakexl.py`, openpyxl (already a dependency) to inspect built files in tests.

**Spec:** `docs/specs/2026-11-19-light-fast-design.md`, section "2. Excel output" (refined 2026-11-19). **Map of today's code:** `docs/audit/2026-11-19-render-map.md` - every task names the sections it builds on; the exact colours, formats, sizes and rules are there and are used verbatim.

## Global Constraints

- **No Excel Table in the built file.** A sheet that contains a table is never copied between workbooks (it crashes Excel on the next replacement). Tables are created in the workbook after the copy.
- Every text cell is written as text: `write_string` in the file, a leading apostrophe in a `Formula2` grid. A string is a formula only when the plan says so (starts with "=" and is not in `plan.literal`).
- The user's Excel is left as it was found: calculation mode, `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, the active workbook and sheet - also when the job fails. Nothing of the user's is saved or closed; only smartass's own answer file is saved.
- A failed job leaves the workbook as it was: sheets it brought in are removed, sheets it parked are restored under their names.
- Same numbers: `checks.sheet_equals_model` on every answer; reply fields and tool parameters unchanged.
- Stage names inside the Excel job are fixed (the bench reads them): `open target`, `draw` with `draw.build`, `draw.move in`, `draw.finish`, `draw.read back`, `draw.sources`, `draw.arrange`, and `save`.
- The temporary file is written under `config.HOME / "tmp"` (never OneDrive) and deleted when the job ends.
- Tests never start Excel; live checks are scripts under `scripts/` run on the owner's PC.
- Line length 120; `uv run --frozen ruff check .` and `uv run --frozen pytest -q` (from `engine/`) green before every commit; commits end with `Co-Authored-By: Claude Fable 5.2 <noreply@anthropic.com>`.
- Branch `feat/file-render` (worktree `.worktrees/file-render`). Never push; the controller opens the PR.

## Review Focus

1. **Text that looks like something else:** a label, a supplier name or a fact that starts with "=", "+", "-" or "@", or looks like a number or a date ("0047", "1-2", "3/4", "1E5"), arrives in Excel as that exact text - in the file and through `Formula2`. (Tasks 2 and 4)
2. **Names already taken:** a leftover "Q5 Data" sheet, or a table "Data_Q5" / "Lines_Q5_b0" on any sheet, stops the job with a message that names it, before anything is copied - never a silent "Q5 Data (2)". (Task 4)
3. **A failure after the move-in** (Excel busy, a formula refused, the read-back raising): the new sheets are gone, parked sheets are back under their names, and the error reaches the caller once. (Tasks 5 and 7)
4. **The user's Excel state** (calculation, screen updating, events, alerts, active sheet) is restored on success and on every failure path, including a failure inside the restore of another setting. (Task 4)
5. **Lines sets at the edges:** zero lines, exactly `config.DRILL_XL_FULL` lines, one more than that (key columns only), and a set above Excel's 1 048 576 rows (refused with a message, not a corrupt file). (Tasks 2 and 5)

---

## File structure

| File | Responsibility |
|---|---|
| `engine/smartass/render/look.py` (new) | Pure: the look of the Answer block (style per cell, heights, widths, merges, status runs, page setup, breaks) from a `SheetPlan`. |
| `engine/smartass/render/xlsx.py` (new) | Builds the temporary file (Answer look, Data, Lines) and returns the tables and formula blocks to apply after the move-in. |
| `engine/smartass/render/xlsx_charts.py` (new) | `ChartSpec` -> an XlsxWriter chart. |
| `engine/smartass/xl/movein.py` (new) | COM: quiet mode, name pre-check, the one-call copy, tables, formula blocks, park / restore / drop. |
| `engine/smartass/pipeline.py` (modify) | `file_job` and `file_redraw_job` beside `excel_job` / `redraw_job`; `config.RENDER` picks. |
| `scripts/compare_renders.py` (new, deleted in Task 9) | Live: draws each test-set case both ways and diffs the sheets. |
| deleted in Task 9 | `xl/owned.py`, `xl/drill_xl.write_lines` and helpers only it uses, the COM writers in `render/excel.py` and `render/charts.py` that the file replaced, and their tests. |

---

### Task 1: `render/look.py` - the Answer's look as a pure function

Map: `3 (the plan object), `4.1-4.3 (styles, formats, heights, merges, page setup), `4.6 (what reads Excel back).

**Files:**
- Create: `engine/smartass/render/look.py`
- Test: `engine/tests/v2/test_look.py`

**Interfaces:**
- Consumes: `render.layout.SheetPlan`, `render.excel.answer_roles(plan)` and `keep_spans(plan)`, `xl.styling`'s pure helpers (`wrap_lines`, `keep_breaks`, the role-to-style rules of `apply_report`), `xl.vendor.tn_rapportstil` (`style_defs()`, `text_units`).
- Produces:

```python
@dataclass(frozen=True)
class CellLook:
    style: str                      # a TN named style, e.g. "TN Heltall · Sum"; "TN Tom" for a blank cell
    num_format: str | None = None   # en-US code laid over the style (map `4.2 overrides)
    font_color: str | None = None   # "#767676" on code cells and helper columns
    align: str | None = None        # "left" on code cells
    wrap: bool | None = None        # True on a merged text or title cell
    valign: str | None = None       # "top" on a merged text cell

@dataclass
class SheetLook:
    cells: dict[tuple[int, int], CellLook]      # every cell of A1:<last_col><last_row>, 1-based (row, col)
    heights: dict[int, float]                   # row -> points
    widths: dict[str, float]                    # column letter -> characters
    merges: list[tuple[int, int, int]]          # (row, first col, last col)
    status_runs: list[tuple[int, int, int]]     # (col, first row, last row): the OK / not-OK conditional formats
    text_rows: list[int]                        # Key point rows whose height depends on the calculated text
    print_area: str                             # "$A$1:$N$60"
    breaks: list[int]                           # rows that start a new printed page
    footer: str                                 # plan.footer

def answer_look(plan: SheetPlan, texts: dict[tuple[int, int], str] | None = None) -> SheetLook
def text_heights(plan: SheetPlan, look: SheetLook, block: tuple) -> dict[int, float]
def style_props(style: str) -> dict
```

  - `texts`: the displayed text of cells whose height depends on it (title, Key points). Without it a Key point row gets the height of its formula's static text (a lower bound); `text_heights` gives the final heights from the block read back from Excel (`block` = the `Value2` rectangle from A1), only for rows whose height differs from `look.heights`.
  - `style_props(style)`: the XlsxWriter format properties of a TN named style, from `tn.style_defs()`: `font_name`, `font_size`, `bold`, `italic`, `font_color`, `align`, `valign`, `text_wrap`, `indent`, `num_format`, `top`/`bottom` and their colours, `bg_color` (map `4.2).

- [ ] **Step 1: Write the failing differential test.** The oracle is today's COM styling run on the fake Excel:

```python
@pytest.mark.parametrize("case", CASES)      # at least: a P&L table with sum and grand rows, a trend with a chart
def test_look_equals_what_apply_report_does(case):   # zone, a lines ranking, an answer with checks and notes
    plan, values = case()                            # values: {(row, col): calculated value} for text and number cells
    ws = fakexl_sheet_with(plan, values)             # the grid written as write_answer does
    styling.apply_report(ws, answer_roles(plan))
    look = answer_look(plan, texts={rc: v for rc, v in values.items() if isinstance(v, str)})
    for (r, c), cl in look.cells.items():
        cell = ws.Cells(r, c)
        assert cell.Style == cl.style, (r, c)
        assert shown_format(cell) == (cl.num_format or style_props(cl.style).get("num_format")), (r, c)
        assert font_colour(cell) == (cl.font_color or style_props(cl.style)["font_color"]), (r, c)
    assert {r: ws.Rows(r).RowHeight for r in look.heights} == look.heights
    assert {c: ws.Columns(c).ColumnWidth for c in look.widths} == look.widths
    assert merged_ranges(ws) == sorted(look.merges)
    assert status_formats(ws) == sorted(look.status_runs)
```

  Build the cases from the plans the existing tests already use (`tests/v2/test_v5_layout.py`, `tests/test_excel_calls.py` and its pinned `fixtures/excel_calls/report_first.json`). The small helpers (`fakexl_sheet_with`, `shown_format`, `font_colour`, `merged_ranges`, `status_formats`) live in the test file and read `fakexl`'s recorded state.

- [ ] **Step 2: Add the pure tests.**

```python
def test_every_cell_of_the_block_has_a_look_and_blank_cells_are_tn_tom(): ...
    # len(look.cells) == plan.last_row * col_number(plan.last_col); an unlisted row is all "TN Tom"
def test_print_area_breaks_and_footer_come_from_the_plan(): ...
    # print_area == f"$A$1:undefinedundefined"; breaks == keep_breaks(heights, keep_spans(plan))
def test_text_heights_only_returns_rows_that_change(): ...
    # a Key point whose calculated text wraps to 3 lines: {row: 3 * 13 + 2}; a one-line point: not in the result
def test_style_props_gives_the_concrete_format_of_a_named_style():
    p = style_props("TN Heltall · Hovedsum")
    assert (p["font_name"], p["font_size"], p["bold"], p["align"]) == ("Arial", 10, True, "right")
    assert p["num_format"] == '#,##0_);(#,##0);"–"_)' and p["top"] == 1 and p["bottom"] == 6   # thin, double
    assert style_props("TN Tittel")["font_color"] == "#1F3864" and style_props("TN Tittel")["bottom"] == 2   # medium
```

- [ ] **Step 3: Run** `uv run --frozen pytest -q tests/v2/test_look.py`. Expected: FAIL, `No module named 'smartass.render.look'`.
- [ ] **Step 4: Implement `look.py`** by moving `apply_report`'s decisions (not its COM calls) into pure code. Do not change `apply_report`: `xl_tidy` keeps using it, and it is this task's oracle.
- [ ] **Step 5: Run the whole suite.** Expected: green.
- [ ] **Step 6: Commit** `feat(render): the Answer's look as a pure function`.

### Task 2: `render/xlsx.py` - build the file

Map: `2 (the sheets), `4.3-4.4 (page setup, Data look), `6.1 (Lines), `10 items 13 and 15 (XlsxWriter traps).

**Files:**
- Create: `engine/smartass/render/xlsx.py`
- Modify: `engine/pyproject.toml` (add `xlsxwriter` to `dependencies`, then `uv lock` - this task may change `engine/uv.lock`)
- Test: `engine/tests/v2/test_xlsx.py`

**Interfaces:**
- Consumes: `look.answer_look`, `look.style_props`, `look.CellLook` (Task 1); `render.excel.answer_grid(plan)`, `data_grid(frame)`; `render.datasheet.Header`; the Lines constants of `xl/drill_xl.py` (`KEY_LABELS`, `TEXT_LABELS`, `AMOUNT_LABELS`, `DATE_LABELS`, `WIDTHS`, `AMT_FMT`) and `config.DRILL_XL_FULL`.
- Produces:

```python
@dataclass
class TableSpec:
    sheet: str; ref: str; name: str          # ref "A14:J120" includes the header row
    stripes: bool                            # False for Data, True for Lines (as today)

@dataclass
class FormulaBlock:
    sheet: str; ref: str; grid: list[list]   # written with ONE Range.Formula2 after the tables exist

@dataclass
class DataSheet:
    name: str; header: Header; grid: list[tuple]      # data_grid(frame): the column row, then one row per fact

@dataclass
class LinesSheet:
    name: str; table: str; title: str; source: str
    columns: list[str]; rows: list[dict]; cube_value: float | None
    header_block: list[list]                           # the F1:K3 rectangle _lines_header writes today

@dataclass
class BuiltFile:
    path: Path; sheets: list[str]; tables: list[TableSpec]; formulas: list[FormulaBlock]

def build(path: Path, answer_name: str, plan: SheetPlan, data: DataSheet, lines: list[LinesSheet]) -> BuiltFile
```

  - Sheet order in the file and in `BuiltFile.sheets`: Answer, Data, then the Lines sheets.
  - **Answer:** no cell contents. Every cell of the block gets its format (one XlsxWriter format per distinct `CellLook`; `style_props` plus the overrides); widths, heights, merges, the two conditional formats per status run, tab colour #1F3864, gridlines hidden, landscape, fit to 1 page wide, margins 1.5 cm left and right, centred horizontally, header `&"Arial"&8Intern`, footer `plan.footer`, print area, page breaks. Charts are added by Task 3 (this task leaves a hook: `build` calls `xlsx_charts.add_charts(workbook, worksheet, plan)` when that module exists).
    `FormulaBlock` for the Answer: `answer_grid(plan)` on its rectangle.
  - **Data:** tab colour #A6A6A6; the fact rows as values at `A<header.table_row>` (label columns text with number format "@", the Value column numbers with `#,##0;(#,##0);"–"`, font #0000C0, right); table header row bold; Arial 9; widths from the cell texts (`tn.text_units(text, 9)`, at most 40, columns A-D at least 16); frozen below the table header row; gridlines hidden; print setup as `sheet_print` does. The header rows 1..h: their formats in the file, their contents as one `FormulaBlock` on `A1:D<h>` (they hold formulas on the table). `TableSpec(name=f"Data_Q{q}", stripes=False)`.
  - **Lines:** tab colour #A6A6A6; A1 title bold, A2 source, C3 "Cube value", D3 the cube value in `AMT_FMT`; row 4 the column names; the lines from row 5 (key columns only above `config.DRILL_XL_FULL` rows); amount and date formats, widths, frozen below row 4, row 4 repeated when printed. `FormulaBlock` on `F1:K3` = `header_block`. `TableSpec(ref="A4:<last><4+n>", stripes=True)`.

- [ ] **Step 1: Write the failing tests** (build to `tmp_path`, read back with openpyxl and the zip):

```python
def test_the_built_file_has_no_table_and_three_sheets_in_order(built):
    with zipfile.ZipFile(built.path) as z:
        assert not [n for n in z.namelist() if n.startswith("xl/tables/")]
    assert openpyxl.load_workbook(built.path).sheetnames == ["Q3 Answer", "Q3 Data", "Q3 Lines b0"] == built.sheets

def test_the_answer_sheet_carries_the_look_and_no_contents(built, plan):
    ws = openpyxl.load_workbook(built.path)["Q3 Answer"]
    assert all(c.value is None for row in ws.iter_rows() for c in row)
    r, c = plan.at(SOME_SUM_KEY)
    cell = ws.cell(r, c)
    assert (cell.font.name, cell.font.sz, cell.font.b) == ("Arial", 10, True) and cell.number_format == '#,##0_);(#,##0);"–"_)'
    assert cell.border.top.style == "thin" and ws.row_dimensions[1].height == 28
    assert ws.sheet_properties.tabColor.rgb.endswith("1F3864") and ws.sheet_view.showGridLines is False
    assert ws.page_setup.orientation == "landscape" and ws.print_area is not None

def test_what_must_happen_after_the_move_in_is_returned(built, plan):
    assert [(t.sheet, t.name, t.stripes) for t in built.tables] == [("Q3 Data", "Data_Q3", False), ("Q3 Lines b0", "Lines_Q3_b0", True)]
    blocks = {(b.sheet, b.ref.split(":")[0]) for b in built.formulas}
    assert blocks == {("Q3 Answer", "B1"), ("Q3 Data", "A1"), ("Q3 Lines b0", "F1")}
    answer = next(b for b in built.formulas if b.sheet == "Q3 Answer")
    assert answer.grid == answer_grid(plan)[2]

@pytest.mark.parametrize("text", ["=SUM(A1)", "+47 900 00 000", "-", "@home", "0047", "1-2", "3/4", "1E5", "TRUE"])
def test_text_stays_text(tmp_path, text):                                 # Review Focus 1
    built = build_with(tmp_path, fact_account=text, line_supplier=text)
    wb = openpyxl.load_workbook(built.path)
    assert wb["Q3 Data"].cell(DATA_ROW, ACCOUNT_COL).value == text and wb["Q3 Data"].cell(DATA_ROW, ACCOUNT_COL).data_type == "s"
    assert wb["Q3 Lines b0"].cell(5, SUPPLIER_COL).value == text

@pytest.mark.parametrize("n, cols", [(0, "all"), (config.DRILL_XL_FULL, "all"), (config.DRILL_XL_FULL + 1, "key")])
def test_lines_at_the_edges(tmp_path, n, cols, monkeypatch):               # Review Focus 5
    monkeypatch.setattr(config, "DRILL_XL_FULL", 50)                       # keep the test fast
    ...   # 0 lines: table ref is the header row plus one empty row ("A4:<last>5"); 51 lines: only KEY_LABELS columns

def test_more_lines_than_excel_has_rows_is_refused(tmp_path):
    with pytest.raises(xlsx.TooManyRows, match="1 048 576"):
        build_with(tmp_path, lines=FakeRows(1_048_573))                    # 4 header rows + lines > the sheet
```

- [ ] **Step 2: Run** them. Expected: FAIL, module missing.
- [ ] **Step 3: Implement `xlsx.py`.** `xlsxwriter.Workbook(path, {"strings_to_formulas": False, "strings_to_numbers": False, "strings_to_urls": False})`; text through `write_string`, numbers through `write_number`, dates as serial numbers with the date format. `TooManyRows(ValueError)` before any file is written.
- [ ] **Step 4: Run the whole suite.** Expected: green.
- [ ] **Step 5: Commit** `feat(render): build an answer's sheets as an .xlsx file`.

### Task 3: `render/xlsx_charts.py` - charts in the file

Map: `5 (every kind, colour, label and axis rule; `5.6 lists what to verify in XlsxWriter).

**Files:**
- Create: `engine/smartass/render/xlsx_charts.py`
- Modify: `engine/smartass/render/xlsx.py` (call it for the Answer sheet)
- Test: `engine/tests/v2/test_xlsx_charts.py`

**Interfaces:**
- Consumes: `render.layout.ChartSpec`; the pure rules of `render/charts.py` (`XL_TYPE` kinds, `HUES`, `PARTS`, `tint`, `hue`, `label_format`, `union_ref`, the legend rule), which stay where they are.
- Produces: `add_charts(workbook, worksheet, sheet_name: str, plan: SheetPlan) -> tuple[list[str], list[str]]` = (titles of the charts drawn, lint lines). A spec without series is skipped. A chart is placed at `spec.anchor`, 480 x 260 pt, named "smartass chart".
- Known limits, each reported as one lint line when it applies, never as an error: the per-point label position of line charts (if XlsxWriter's custom labels take no position, labelled points keep the series position), and the 10 pt lift of labels on thin waterfall steps (dark ink without the lift).

- [ ] **Step 1: Write the failing tests.** One per kind (`column`, `bar`, `line`, `stacked`, `stacked100`, `combo`, `variance`, `waterfall`), each building a file from a `ChartSpec` and asserting on the chart read back with openpyxl (or the chart XML): the chart type and grouping, every series' values and category reference (`'Q3 Answer'!$D$5:$D$9`; a range with gaps as a union), the fill or line colour per role (map `5.3), gap 50 and overlap 100 where the map says, the data-label number format (`label_format`), the title text, the legend rule, hidden value-axis gridlines and labels. Plus:

```python
def test_a_spec_without_series_is_skipped_and_not_reported_as_drawn(): ...
def test_a_chart_that_cannot_be_built_is_one_lint_line_and_the_others_are_drawn(): ...
def test_base_series_of_a_waterfall_are_invisible(): ...          # no fill, no line
def test_highlighted_points_get_the_accent_and_the_rest_grey(): ...   # #1C16C5 / #A6A6A6
```

- [ ] **Step 2: Run** them. Expected: FAIL.
- [ ] **Step 3: Implement.** Verify each needed XlsxWriter feature against its documentation for the installed version before relying on it; where one is missing, apply the stated limit and its lint line.
- [ ] **Step 4: Run the whole suite.** Expected: green.
- [ ] **Step 5: Commit** `feat(render): charts in the built file`.

### Task 4: `xl/movein.py` - move the sheets in and finish them

Map: `1.1 (Target), `1.4 (helpers), `10 items 3, 7, 9, 10, 14.

**Files:**
- Create: `engine/smartass/xl/movein.py`
- Modify: `engine/tests/fakexl.py` (only what the tests need: `Workbooks.Open` of a built file as a fake workbook, `Worksheets([...]).Copy`, `LinkSources` / `BreakLink`, `ListObjects.Add`, sheet `Delete` and rename)
- Test: `engine/tests/test_movein.py`

**Interfaces:**
- Consumes: `xlsx.BuiltFile`, `TableSpec`, `FormulaBlock` (Task 2); `xl.jobs.com` for every COM call; `xl.target.Target` (`t.app`, `t.wb`, `t.sheets_added`).
- Produces:

```python
class NameTaken(InputError): ...     # a sheet or table name of the answer already exists in the workbook

@contextmanager
def quiet(app): ...                  # ScreenUpdating, EnableEvents, DisplayAlerts off, Calculation manual; all four
                                     # restored in reverse order, each in its own try; the active sheet re-activated
def taken(wb, sheets: list[str], tables: list[str]) -> list[str]
def move_in(t, built: BuiltFile, after) -> None      # NameTaken first; open read-only + hidden, ONE Copy, close,
                                                     # break links; every copied name appended to t.sheets_added
def finish(t, built: BuiltFile) -> None              # each TableSpec: ListObjects.Add + Name + TableStyleLight1
                                                     # (+ stripes); then each FormulaBlock: one Formula2
def calculate(t, sheets: list[str]) -> None          # Worksheet.Calculate per sheet (Data, Lines, then Answer)

@dataclass
class Parked:
    sheets: dict[str, str]           # old name -> parked name ("old Q3 Answer")
    tables: dict[str, str]           # old name -> parked name ("Data_Q3_old")

def park(t, sheets: list[str], tables: list[str]) -> Parked
def restore(t, parked: Parked) -> None               # after a failure: names back
def drop(t, parked: Parked) -> None                  # after success: the old sheets deleted in ONE call
```

- [ ] **Step 1: Write the failing tests** on `fakexl`:

```python
def test_move_in_is_one_copy_and_at_most_eight_com_calls(t, built):
    with fakexl.count_com() as n:
        movein.move_in(t, built, after=t.wb.Worksheets("Sources"))
    assert n.total <= 8 and n.named("Copy") == 1
    assert [ws.Name for ws in t.wb.Worksheets][-3:] == built.sheets and t.sheets_added[-3:] == built.sheets
    assert t.wb.LinkSources(1) is None and t.app.Workbooks.Count == 1          # the temp file is closed

def test_finish_makes_the_tables_then_writes_each_block_once(t, built): ...
    # ListObjects names == ["Data_Q3", "Lines_Q3_b0"]; Data: no row stripes; each FormulaBlock one Formula2 write;
    # the tables exist before the first Formula2 (fakexl refuses a formula on a missing table)

@pytest.mark.parametrize("leftover", ["sheet:Q3 Data", "table:Data_Q3", "table:Lines_Q3_b0@Other sheet"])
def test_a_taken_name_stops_before_anything_is_copied(t, built, leftover):     # Review Focus 2
    with pytest.raises(movein.NameTaken, match=leftover.split(":")[1].split("@")[0]):
        movein.move_in(t, built, after=None)
    assert fakexl.snapshot(t.wb) == BEFORE and t.app.Workbooks.Count == 1

@pytest.mark.parametrize("fail_at", ["ScreenUpdating", "Calculation", "body"])
def test_quiet_restores_everything_whatever_fails(app, fail_at):               # Review Focus 4
    ...   # the four settings and the active sheet equal their values before; the original error is the one raised

def test_formula_grids_keep_text_as_text(t, built): ...                        # Review Focus 1: "'=SUM(A1)" stays text
def test_park_restore_and_drop(t): ...
    # park: sheets "old Q3 Answer", table "Data_Q3_old"; restore: names back, snapshot == BEFORE; drop: one Delete call
def test_the_temp_workbook_is_closed_when_the_copy_fails(t, built): ...
```

- [ ] **Step 2: Run** them. Expected: FAIL.
- [ ] **Step 3: Implement.** `Workbooks.Open(path, 0, True, ..., AddToMru=False)` positional (late binding ignores keywords, map `1.2 step 4); hide the temp workbook's window before the copy; positional `Copy(None, after)`.
- [ ] **Step 4: Run the whole suite.** Expected: green.
- [ ] **Step 5: Commit** `feat(xl): move built sheets into the workbook and finish them`.

### Task 5: `pipeline.file_job` - a new answer the file way

Map: `1.2 (today's job, step by step), `7 (read-back and checks), `10 items 2, 5, 16, 17.

**Files:**
- Modify: `engine/smartass/core/config.py` (`RENDER = os.environ.get("SV_RENDER", "com")`, values `"com"` | `"file"`), `engine/smartass/pipeline.py`, `engine/smartass/server.py` (pick the job by `config.RENDER`), `engine/smartass/render/check.py` (the chart rule counts `plan`'s drawn charts when no COM chart was drawn), `scripts/bench_answers.py` (columns for the new stages)
- Test: `engine/tests/v2/test_file_job.py`

**Interfaces:**
- Consumes: Tasks 1-4; today's `_target`, `autosave_paused`, `_next_q`, `layout(...)` call exactly as `_draw` makes it (map `1.2 step 6.5), `header_rows`, `read_block`, `compare`, `write_status`, `_write_sources`, `_save_spec`, `_arrange`, `_check`, `_save`, `_rollback`, `reply`, `remember`.
- Produces: `pipeline.file_job(h, pre, target="session") -> dict` with the same reply as `excel_job`. Sequence (stage names from the Global Constraints):
  1. `open target`; `autosave_paused`; `q = _next_q(t)`.
  2. `draw.build`: header, plan, footer as today; `LinesSheet` per line set; `xlsx.build(config.HOME / "tmp" / f"{uuid}.xlsx", ...)`.
  3. inside `movein.quiet(t.app)` - `draw.move in`: `move_in` after the question's place (the anchor `xl.sheets.add_sheet` would use); `draw.finish`: `finish`, `calculate`; `draw.read back`: `read_block` + `compare`, `write_status`, `look.text_heights` applied to the rows it returns (one call per run of equal heights).
  4. `draw.sources`: `_write_sources`; `_save_spec` without `owned`; `draw.arrange`: `_arrange`; `_check`.
  5. `save`; on any exception `_rollback(t)` (the copied sheets are in `t.sheets_added`) and re-raise; the temp file is deleted in a `finally`.

- [ ] **Step 1: Write the failing tests** (seams monkeypatched as `tests/v2/test_v2_pipeline_job.py`'s `wire()` does, plus `fakexl` for the end-to-end one):

```python
def test_the_file_job_runs_its_steps_in_order(wired):
    out = pipeline.file_job(H, PRE)
    assert wired.calls == ["target", "next_q", "build", "move_in", "finish", "calculate", "read_block", "status",
                           "sources", "save_spec", "arrange", "check", "save", "remember"]
    assert out["sheets"] == ["Q1 Answer", "Q1 Data"] and out["checks"]["sheet_equals_model"] == "OK"

def test_a_failure_after_the_move_in_removes_the_new_sheets_and_the_temp_file(wired, tmp_home):   # Review Focus 3
    wired.fail("finish", HostError("Excel is busy"))
    with pytest.raises(HostError, match="busy"):
        pipeline.file_job(H, PRE)
    assert wired.rolled_back == ["Q1 Answer", "Q1 Data"] and not list((tmp_home / "tmp").glob("*.xlsx"))
    assert "save" not in wired.calls and "remember" not in wired.calls

def test_the_whole_job_on_the_fake_excel_makes_at_most_45_com_calls(fake_target, pre_one_query):
    with fakexl.count_com() as n:
        out = pipeline.file_job(H, pre_one_query)
    assert n.total <= 45 and out["checks"]["sheet_equals_model"] in ("OK", "DIFF in 0 cells")
    ...   # stage names recorded: open target, draw.build, draw.move in, draw.finish, draw.read back, draw.sources,
          # draw.arrange, save

def test_config_render_picks_the_job(monkeypatch): ...            # "com" -> excel_job, "file" -> file_job, in server.py
def test_a_lines_block_with_no_lines_draws_an_empty_table(fake_target, pre_lines_empty): ...      # Review Focus 5
```

- [ ] **Step 2: Run** them. Expected: FAIL.
- [ ] **Step 3: Implement `file_job`** and the `config.RENDER` switch (default stays `"com"`).
- [ ] **Step 4: Run the whole suite.** Expected: green.
- [ ] **Step 5: Live smoke** on the owner's PC: `SV_RENDER=file` and `scripts\bench_answers.py --rounds 1 --case b02`; the answer must open in Excel with `sheet_equals_model` OK. Report the stage seconds. If Excel shows a dialog or the job cannot proceed, STOP and report.
- [ ] **Step 6: Commit** `feat(pipeline): file_job - a new answer through a built file (SV_RENDER=file)`.

### Task 6: Live comparison of the two renderers (parity gate)

**Files:**
- Create: `scripts/compare_renders.py`, `docs/audit/2026-11-19-render-comparison.md`

**Interfaces:**
- Produces: `python scripts/compare_renders.py [--case b01 ...]`: draws each test-set case with `SV_RENDER=com` into one new workbook and with `SV_RENDER=file` into another (the three cases without a chart block as `bench_answers.py` runs them), then for each pair of Q sheets diffs, cell by cell over the used range: `Formula2`, `Value2` (numbers within 0.005), number format (en-US), font name / size / bold / italic / colour, fill, top and bottom border, horizontal alignment, wrap; and per sheet: row heights, column widths, merges, freeze row, tab colour, print area, orientation, table names, chart count and per chart its type, series formulas and title. Prints counts and addresses of differences (no cell values), and exits 1 when any difference is outside the allowed list.
- Allowed differences (anything else is a defect to fix in Tasks 1-5): the cell style **name** ("Normal" against "TN ..."); the hidden sheet name `smartass_pagesetup`; column widths of the Data sheet within 1.0 character (computed against AutoFit); the two chart limits of Task 3.

- [ ] **Step 1: Write the script** (it reads Excel through the same job runner the bench uses; it changes nothing but the two workbooks it creates).
- [ ] **Step 2: Run it live** over all nine cases. Fix every difference outside the allowed list in the module that owns it (each fix with a unit test that pins it), and re-run until the script exits 0.
- [ ] **Step 3: Write the audit file:** the final output, the list of fixes made, and screenshots are not needed - the owner compares the charts of one case by eye afterwards (name the two workbooks).
- [ ] **Step 4: Commit** `docs(audit): file-built answers match the COM renderer` (plus the fix commits).

### Task 7: `pipeline.file_redraw_job` - follow-ups and refresh replace the Q sheets

Map: `1.3 (today's redraw), `8 (follow-ups, `_svmeta`), `10 item 8.

**Files:**
- Modify: `engine/smartass/pipeline.py`, `engine/smartass/server.py`
- Test: `engine/tests/v2/test_file_redraw.py`

**Interfaces:**
- Consumes: `movein.park`, `restore`, `drop` (Task 4); `file_job`'s draw steps (Task 5) factored so both jobs share them; `xl.sheets.question_of` to find the sheets of Q<n>.
- Produces: `pipeline.file_redraw_job(h, pre, q, owned, target) -> dict` (`owned` is accepted and ignored), reply as `redraw_job` with `"followup": True`. Sequence: `open target`; `autosave_paused`; the sheets of Q<n> = "Q<n> Answer", "Q<n> Data" and every "Q<n> Lines ..." present, and their tables; `park` them; then the draw steps of `file_job` with this `q`; on success `drop`; on any failure remove the new sheets, `restore`, and raise `RestoreError` only when the restore itself failed (as today). A Q<n> whose Answer or Data sheet is missing is an `InputError` that names it.

- [ ] **Step 1: Write the failing tests:**

```python
def test_a_follow_up_replaces_every_sheet_of_the_question(fake_target_with_q3, pre_q3_without_lines):
    out = pipeline.file_redraw_job(H, pre_q3_without_lines, 3, None, PATH)
    names = [ws.Name for ws in fake_target_with_q3.wb.Worksheets]
    assert "Q3 Lines b0" not in names and not [n for n in names if n.startswith("old ")]     # the dropped block's sheet goes
    assert [lo for lo in all_tables(fake_target_with_q3.wb)] == ["Data_Q3"] and out["followup"] is True
    assert sheet_position(names, "Q3 Answer") == POSITION_BEFORE                               # same place in the workbook

def test_a_failed_follow_up_leaves_the_old_answer_exactly_as_it_was(fake_target_with_q3, pre_q3, wired):   # Review Focus 3
    before = fakexl.snapshot(fake_target_with_q3.wb)
    wired.fail("finish", HostError("Excel is busy"))
    with pytest.raises(HostError):
        pipeline.file_redraw_job(H, pre_q3, 3, None, PATH)
    assert fakexl.snapshot(fake_target_with_q3.wb) == before

def test_a_restore_that_fails_is_a_restore_error(...): ...
def test_other_questions_and_the_users_own_sheets_are_untouched(...): ...     # snapshot of every sheet not of Q3 equal
def test_a_question_without_its_data_sheet_is_refused_by_name(...): ...
def test_fin_refresh_and_the_follow_up_path_use_the_file_job_when_configured(monkeypatch): ...
```

- [ ] **Step 2: Run** them. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run the whole suite.** Expected: green.
- [ ] **Step 5: Live check:** `SV_RENDER=file`: one answer, then six follow-ups on it in a row through `fin_analyze(patch=...)`'s path (a script under `scripts/`, kept: `scripts/check_followups.py`), in the owner's Excel; each must end with `sheet_equals_model` OK and Excel must still answer. Report the seconds per follow-up.
- [ ] **Step 6: Commit** `feat(pipeline): follow-ups and refresh replace the Q sheets (file renderer)`.

### Task 8: Slim the workbook-wide tail

Map: `6.2 (Sources log), `8.3 (`_svmeta`), `1.2 steps 6.14 and 7 (arrange, save).

**Files:**
- Modify: `engine/smartass/xl/sources.py`, `engine/smartass/xl/sheets.py` (`save_spec`), `engine/smartass/pipeline.py` (`_save`)
- Test: `engine/tests/v2/test_v5_sources.py`, `engine/tests/v2/test_v2_save_spec.py`, `engine/tests/v2/test_v2_autosave.py`

**Interfaces:**
- Produces (same functions, same results on the sheet, fewer calls):
  - `sources.write_sources`: the new rows written with one `Value2` rectangle; fixed column widths instead of `AutoFit` (set only when the sheet is created); fonts and the control-total format set for the new rows only; freeze and print setup only when the sheet is created. At most 12 COM calls for an answer with one query into an existing log.
  - `sheets.save_spec`: the row's cells written with one `Value2` rectangle (A plus every chunk).
  - `pipeline._save(t)`: when AutoSave was on before the job and is switched back on for our own file, no explicit `Save` (AutoSave writes it); otherwise `wb.Save()` as today. The reply is unchanged.

- [ ] **Step 1: Write the failing tests:** call counts with `fakexl.count_com()` (`<= 12` for the log, `<= 3` for `save_spec`), the existing content assertions unchanged, and

```python
def test_no_explicit_save_when_autosave_takes_over(): ...     # AutoSaveOn True before -> Save not called, AutoSaveOn True after
def test_an_explicit_save_when_autosave_is_off_or_unknown(): ...   # AutoSaveOn False, or the property raising
def test_a_users_workbook_is_never_saved(): ...               # as today
```

- [ ] **Step 2: Run** them. Expected: FAIL on the counts.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run the whole suite.** Expected: green.
- [ ] **Step 5: Live check of the save rule:** draw one answer with `SV_RENDER=file` into a new answer workbook and confirm the file on disk has the new sheets within 30 s without an explicit `Save` (open the local file read-only with openpyxl after the wait). If it does not, keep the explicit `Save` and say so in the report.
- [ ] **Step 6: Commit** `perf(xl): Sources log, spec store and save in a handful of calls`.

### Task 8A: `xl/fastcom.py` - a late-bound COM wrapper that learns each interface once

Added by the controller on 2026-11-19 after probing the owner's Excel (ledger: "Controller probes"). It comes before
the rest of the performance work because it speeds up every COM call the engine makes, in both renderers and in the
`xl_*` tools.

**What was measured (owner's Excel, 10 workbooks open):** a pure COM round trip costs 0.4 ms and a raw
`IDispatch.Invoke` that returns an object 2.2-3.5 ms, but the same step through pywin32's standard late-bound wrapper
(`win32com.client.dynamic.CDispatch`) costs 7-14 ms: for every object a call returns it fetches the type information
again, builds a new description, and binds each attribute name again on first use (`_LazyAddAttr_`: four
`ITypeComp.Bind` calls per name), and it compiles a new method stub per object. A `jobs.com` call is typically 3-5
such hops, which is the 56-73 ms per call seen in both renderers. A prototype that shares what is learnt per
interface gave x2.5-3.2 on chains like `ws.Range("B2").Font.Name` with identical results on every case tried.

**Files:**
- Create: `engine/smartass/xl/fastcom.py`
- Modify: every place the engine turns a raw COM pointer or a standard pywin32 object into the Excel object it works
  with (find them: `win32com.client.Dispatch`, `DispatchEx`, `GetObject`, `GetActiveObject`, `dynamic.Dispatch`,
  `AccessibleObjectFromWindow` / `ObjectFromLresult` in `engine/smartass` - at least `xl/instances.py`, `xl/host.py`,
  `xl/target.py`); `docs/troubleshooting.md` (one short entry for the switch)
- Create: `scripts/check_fastcom.py` (live equivalence + timing check; kept)
- Test: `engine/tests/test_fastcom.py`

**Interfaces:**
- Produces:
  - `fastcom.wrap(obj, name: str | None = None)` -> the object wrapped fast. `obj` is a `PyIDispatch` or any pywin32
    dispatch object (its `_oleobj_` is used). Wrapping an already fast object returns it unchanged. With the
    environment variable `SV_FASTCOM=0` (read once, as `config.FASTCOM`) it returns the standard wrapper
    (`win32com.client.dynamic.Dispatch`): the switch exists to rule the wrapper out when diagnosing.
  - `class fastcom.Fast(dynamic.CDispatch)`: pywin32's own late-bound wrapper with one change - what it learns about
    an interface is shared by every object of that interface. Everything an attribute get, a call, a method result
    or an iteration returns that is a COM object is wrapped `Fast` too.
- Behaviour that must NOT change (the engine's 5 000 lines of Excel code rely on pywin32's late-bound semantics):
  the same attribute names, property / method distinction, optional arguments, default members
  (`Worksheets("name")`, `Cells(r, c)`), property puts, iteration, `pythoncom.com_error` raised unchanged, tuple /
  date / currency conversions. No early binding (`gencache` / makepy) - its generated classes rename parameterised
  properties (`GetOffset`, `GetEnd` ...) and would break existing code that only real Excel exercises.

**Design that the prototype used (controller's spike; read `engine/.venv/Lib/site-packages/win32com/client/dynamic.py`
and `build.py` of the pinned pywin32 before writing it):**

```python
_IFACE = {}     # interface IID -> shared state: the olerepr (name maps) and the compiled method stubs

class Fast(dynamic.CDispatch):
    def __init__(self, ob, iface, userName=None, typeinfo=None): ...
        # CDispatch.__init__(self, ob, iface.olerepr, userName, lazydata=...) and _builtMethods_ = iface.methods
    def _wrap_dispatch_(self, ob, userName=None, returnCLSID=None):
        return wrap(ob, userName)
    def _make_method_(self, name): ...
        # as CDispatch._make_method_, but "Dispatch" in the generated stub's globals is a function that calls
        # wrap(), and the compiled function is stored in the shared iface.methods

def wrap(ob, userName=None):
    ti = ob.GetTypeInfo(); iid = ti.GetTypeAttr()[0]          # two cheap calls per new object
    iface = _IFACE.get(iid) or first_time(iid, ob, ti)        # dynamic.MakeOleRepr(ob, ti, lazydata) once per IID
    return Fast(ob, iface, userName, ti)
    # any com_error on the way: fall back to dynamic.Dispatch(ob, userName)
```

- **Threading (required):** user-Excel jobs run on their own STA threads (`xl/host.py`), so a COM pointer learnt on
  one thread must never be used on another. The shared state per interface is pure Python only (the olerepr's maps,
  the compiled stubs, guarded by a lock where it is mutated). The type information used to bind a name that is not
  in the shared maps yet comes from **the object at hand** (its own `GetTypeInfo()` / `GetTypeComp()`, fetched on
  that thread, the typecomp only when a new name must be bound) - never from a pointer cached by another thread.
- The shared olerepr is filled lazily by whichever object first uses a name; after warm-up no type-information
  call is made beyond the two in `wrap`.

- [ ] **Step 1: Write the failing unit tests** against a COM server every Windows has, in-process, no Excel:
  `Scripting.FileSystemObject` and `Scripting.Dictionary` (`scrrun.dll`). Compare `fastcom.wrap(raw)` with
  `dynamic.Dispatch(raw)` on the same operations:

```python
def test_results_equal_the_standard_wrapper(tmp_path):      # GetFolder(p).Files.Count, .Name, .Path, SubFolders
def test_objects_returned_by_gets_calls_and_iteration_are_fast(tmp_path):   # isinstance(x, fastcom.Fast) for each
def test_an_interface_is_learnt_once(tmp_path):
    # wrap two Folder objects; after using the same names on both, len(fastcom._IFACE) and the shared method-stub
    # dict did not grow for the second, and (counting with a spy on ITypeComp.Bind via a wrapper class or by
    # monkeypatching Fast._LazyAddAttr_) no bind happened for the second object
def test_a_property_put_and_a_method_with_arguments_work():  # Dictionary: d.Add("k", 1); d.Item("k"); d.CompareMode = 1
def test_a_com_error_is_raised_unchanged():                   # fso.GetFolder(missing) -> pythoncom.com_error, same hresult
def test_wrapping_twice_returns_the_same_object_and_the_switch_gives_the_standard_wrapper(monkeypatch):
def test_an_object_wrapped_on_one_thread_does_not_hand_its_type_information_to_another():
    # learn the interface on thread A (CoInitialize there), then on thread B wrap a NEW object of the same
    # interface and use a name NOT yet bound: it must work, and the typeinfo used must be thread B's own object
def test_an_object_without_type_information_falls_back_to_the_standard_wrapper():   # a fake PyIDispatch-like raising com_error
```

- [ ] **Step 2: Run** `uv run --frozen pytest -q tests/test_fastcom.py`. Expected: FAIL, module missing.
- [ ] **Step 3: Implement `fastcom.py`**, then wrap at every root so everything downstream is fast. A root is wrapped
  once; nothing else in the engine changes.
- [ ] **Step 4: Run the whole suite.** Expected: green (the fakexl tests are unaffected: fakes are not COM objects -
  `wrap` must pass through anything that is neither a `PyIDispatch` nor has `_oleobj_`).
- [ ] **Step 5: Write and run `scripts/check_fastcom.py` (live, in a HIDDEN Excel of its own via `DispatchEx`, never
  the user's Excel; a scratch workbook that is never saved).** For each idiom the engine uses, run it through the
  standard wrapper and through `fastcom` and require equal results (or the same `com_error` hresult), then time both:
  ranges (`Range("A1:C3")`, `Range(Cells(1,1), Cells(3,3))`, `Cells(r, c)`, `Rows(n)` / `Columns(n)`, `UsedRange`,
  `Offset`, `Resize`, `End(-4162)`, `EntireRow`), `Value` / `Value2` / `Formula` / `Formula2` get and set with
  scalars and 2D tuples, `NumberFormat` / `NumberFormatLocal`, `Font` / `Interior` / `Borders(9)` sets,
  `ColumnWidth` / `RowHeight`, `Merge` / `MergeCells`, `FormatConditions.Add` + `Delete`, `Worksheets.Add(None,
  after)` / `Name` / `Tab.Color` / `Visible` / `Move` / `Delete`, `Worksheets([names]).Copy(None, after)` with a
  list argument, `ListObjects.Add(1, rng, None, 1)` + `Name` + `TableStyle` + `Range` / `DataBodyRange`,
  `Names.Add` and iteration over `wb.Names` / `Worksheets` / `ListObjects`, `Hyperlinks.Add`, `AutoFilter`,
  `PageSetup` gets and sets, `HPageBreaks.Add`, `ChartObjects()` / `Shapes.AddChart2` + a series' `Formula` /
  `Points(i).DataLabel`, `Application.Calculation` / `ScreenUpdating` / `EnableEvents` / `DisplayAlerts` /
  `International` (as the engine reads it) / `CentimetersToPoints` / `WorksheetFunction.CountA`, `Workbooks.Add` /
  `Open(path, 0, True)` / `Close(False)` / `LinkSources(1)`, a missing sheet and a missing table (same error). Print
  one line per idiom: equal yes/no and the two medians; exit 1 on any inequality. Take the list of idioms from the
  code (grep `com(lambda` in `engine/smartass`) rather than only from this list.
- [ ] **Step 6: Live in the owner's Excel** (the same safety rules as the earlier live tasks): run
  `scripts/check_followups.py` once with `SV_RENDER=file` and once with `SV_RENDER=com`; both must end with every
  answer `sheet_equals_model` OK. Report the Excel-job seconds and COM calls of each next to the figures from before
  this task (file renderer follow-up: 12-15 s at 121 calls; COM renderer new answer: 15-45 s at about 280 calls).
- [ ] **Step 7: Commit** `perf(xl): a late-bound COM wrapper that learns each interface once`.

### Task 9: Speed gate, switch over, delete the old writers

**Files:**
- Create: `docs/audit/2026-11-19-file-render-bench.md`
- Modify: `engine/smartass/core/config.py` (`RENDER` removed), `pipeline.py` and `server.py` (the file jobs become `excel_job` / `redraw_job`), docs (`docs/DESIGN.md`, `skills/smartass/reference/operations.md`, `CHANGELOG.md`)
- Delete: `engine/smartass/xl/owned.py`, the COM writers the file replaced in `render/excel.py` (`write_data`, `rewrite_data`, `rewrite_grid`, `_write_header`, `_style`, `write_answer`), `render/charts.py`'s COM drawing (`_place` ... `write_charts`; its pure rules stay), `xl/drill_xl.write_lines` and what only it uses, `pipeline._write_lines`, `_lines_header`, `_backup`, `_restore`, `_make_room`, `_clear`; `scripts/compare_renders.py`; their tests (map `9.2 column "if the COM writer is replaced")

**Interfaces:**
- Consumes: Task 6's audit says the comparison exits 0; Tasks 7-8 are merged into the branch.
- Gate (spec): `scripts/bench_answers.py --rounds 3` with `SV_RENDER=file`, live: median `excel job._total` at most **3 s** without a lines sheet and **5 s** with one, in every round.

- [ ] **Step 1: Run the bench live** and write the audit file: the baseline table beside the new one, per round, with the stage table.
- [ ] **Step 2: If the gate is missed, STOP** and report the stage table to the controller - do not delete anything.
- [ ] **Step 3: Write the failing test** in `engine/tests/test_doctor.py`:

```python
def test_the_com_writers_are_gone():
    import smartass, pathlib
    root = pathlib.Path(smartass.__file__).parent
    assert not (root / "xl" / "owned.py").exists()
    src = (root / "pipeline.py").read_text(encoding="utf-8")
    assert "RENDER" not in src and "owned" not in src and "_backup(" not in src
    assert not hasattr(__import__("smartass.render.excel", fromlist=["x"]), "write_answer")
```

- [ ] **Step 4: Switch and delete.** A test of deleted behaviour is deleted; a test of behaviour that still exists moves to the new job. `xl/styling.apply_report` and `xl_tidy` stay.
- [ ] **Step 5: Run** `uv run --frozen ruff check . && uv run --frozen pytest -q`. Expected: green; state the test counts before and after.
- [ ] **Step 6: Update the docs** to what is left, with the measured seconds.
- [ ] **Step 7: Commit** `feat(render)!: answers are built as a file and moved in; the COM writers are gone`.
