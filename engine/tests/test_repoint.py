"""xl/repoint.py: a formula's references to a follow-up's parked sheets and tables named as the new answer's - exactly
those references (light and fast phase 3b, Task 3). Pure text: the cases Excel writes in Formula2, FormulaArray, a
series formula and a name's RefersToLocal."""
from __future__ import annotations

import pytest

from smartass.xl.repoint import Repoint, plan, repoint, scan_texts

SHEETS = {"old q3 answer": "Q3 Answer", "old q3 data": "Q3 Data", "old o'brien": "O'Brien"}
TABLES = {"data_q3_old": "Data_Q3", "lines_q3_b0_old": "Lines_Q3_b0"}


@pytest.mark.parametrize("before, after", [
    ("='old Q3 Answer'!E12", "='Q3 Answer'!E12"),                                        # a cell
    ("=SUM('old Q3 Answer'!A1:A3)", "=SUM('Q3 Answer'!A1:A3)"),                          # a cross-sheet SUM
    ("=SUM(Data_Q3_old[Value])", "=SUM(Data_Q3[Value])"),                                # a column of the table
    ("=ROWS(Data_Q3_old)", "=ROWS(Data_Q3)"),                                            # the table itself
    ("=Data_Q3_old[[#Headers],[Value]]", "=Data_Q3[[#Headers],[Value]]"),
    ("=@Data_Q3_old[Value]", "=@Data_Q3[Value]"),
    ("=SORT('old Q3 Answer'!A1:A3)", "=SORT('Q3 Answer'!A1:A3)"),                        # a dynamic array
    ("=SUM('Summary:old Q3 Answer'!E12)", "=SUM('Summary:Q3 Answer'!E12)"),              # a 3D reference's end
    ("='old Q3 Answer'!E12+'old Q3 Data'!B2*Lines_Q3_b0_old[AMOUNT]",
     "='Q3 Answer'!E12+'Q3 Data'!B2*Lines_Q3_b0[AMOUNT]"),
    ("='OLD Q3 ANSWER'!E12", "='Q3 Answer'!E12"),                                        # names without case
    ("='old O''Brien'!A1", "='O''Brien'!A1"),                                            # a quote in the name
    ("=SERIES('old Q3 Answer'!$B$1,'old Q3 Answer'!$A$2:$A$5,'old Q3 Answer'!$B$2:$B$5,1)",
     "=SERIES('Q3 Answer'!$B$1,'Q3 Answer'!$A$2:$A$5,'Q3 Answer'!$B$2:$B$5,1)"),       # a chart series
    ("=SUMMER(Data_Q3_old[Verdi];'old Q3 Answer'!$E$12)",
     "=SUMMER(Data_Q3[Verdi];'Q3 Answer'!$E$12)"),                                      # a name's RefersToLocal
    ("=SØK(\"x\";'old Q3 Answer'!A1)", "=SØK(\"x\";'Q3 Answer'!A1)"),
    ("='old Q3 Answer'!A1:A3*{1;2;3}", "='Q3 Answer'!A1:A3*{1;2;3}"),                    # FormulaArray's constant
])
def test_a_reference_to_a_parked_sheet_or_table_names_the_new_one(before, after):
    assert repoint(before, SHEETS, TABLES) == after


@pytest.mark.parametrize("formula", [
    "='old Q3 Answer 2'!A1",                    # a sheet whose name merely contains a parked one
    "='my old Q3 Answer'!A1",
    "=\"'old Q3 Answer'!E12\"&1",              # text inside a string literal (Excel leaves it too)
    "=\"a\"\"'old Q3 Answer'!A1\"\"\"&1",            # ... with doubled quotes in it
    "=INDIRECT(\"Data_Q3_old[Value]\")",
    "=SUM(Data_Q3_old2[Value])",                # another table whose name starts like it
    "=SUM(MyData_Q3_old[Value])",
    "=Data_Q3_old.x+Data_Q3_old_1",             # defined names that contain it
    "=SUM(T[Data_Q3_old])",                     # a column of another table that is called so
    "=T[a']Data_Q3_old]",                       # ... with an escaped bracket before it
    "='[Other.xlsx]old Q3 Answer'!A1",          # another workbook's sheet
    "='C:\\x\\[Other.xlsx]old Q3 Answer'!A1",
    "=Other.xlsx!Data_Q3_old",                  # another workbook's name
    "=Summary!Data_Q3_old",                     # a sheet's own name
    "='old Q3 Lines b0'!A1",                    # a parked sheet the new answer has no namesake for
    "=1+2",
    "",
])
def test_anything_else_is_left_as_it_is(formula):
    assert repoint(formula, SHEETS, TABLES) == formula


def test_a_parked_sheet_without_a_namesake_stays_in_a_formula_that_is_re_pointed_otherwise():
    assert (repoint("='old Q3 Lines b0'!A1+'old Q3 Answer'!E12", SHEETS, TABLES)
            == "='old Q3 Lines b0'!A1+'Q3 Answer'!E12")


def test_the_scan_looks_for_the_parked_sheets_common_start_and_each_table():
    """The gate texts are loose (the rewrite decides exactly): the parked sheets' common start, each parked table."""
    assert scan_texts(["old Q3 Answer", "old Q3 Data"], ["Data_Q3_old"]) == ["old Q3 ", "Data_Q3_old"]
    assert scan_texts(["old Q3 Answer"], []) == ["old Q3 Answer"]
    assert scan_texts(["old Q3 Answer", "x"], ["Data_Q3_old", "Lines_Q3_b0_old"]) == [
        "old Q3 Answer", "x", "Data_Q3_old", "Lines_Q3_b0_old"]       # no common start: each name


def test_the_plan_rewrites_what_mentions_a_parked_name_and_leaves_out_the_skipped_sheets_own_names():
    rows = [("cell", "Summary", "A1", "='old Q3 Answer'!E12"),
            ("cell", "Summary", "A2", "=\"'old Q3 Answer'!E12\""),                  # nothing to rewrite
            ("array", "Summary", "D1:D3", "='old Q3 Answer'!A1:A3*2"),
            ("series", "Summary", "1/2", "=SERIES(,,'old Q3 Data'!$B$2:$B$5,1)"),
            ("chart", "Chart1", "1", "=SERIES(,,Data_Q3_old[Value],1)"),
            ("name", "TotalQ", "4", "='old Q3 Answer'!$E$12"),
            ("name", "'old Q3 Answer'!Print_Area", "5", "='old Q3 Answer'!$A$1:$H$40"),   # the parked sheet's own
            ("name", "'Q3 Answer'!Print_Area", "6", "='Q3 Answer'!$A$1:$H$40"),
            ("name", "Summary!Local", "7", "=Data_Q3_old[Value]"),
            ("unread", "Locked", "", "a protected sheet")]
    got, unread = plan(rows, SHEETS, TABLES, ["old Q3 Answer", "old Q3 Data", "Q3 Answer", "Q3 Data"])
    assert got == [Repoint("cell", "Summary", "A1", "='old Q3 Answer'!E12", "='Q3 Answer'!E12"),
                   Repoint("array", "Summary", "D1:D3", "='old Q3 Answer'!A1:A3*2", "='Q3 Answer'!A1:A3*2"),
                   Repoint("series", "Summary", "1/2", "=SERIES(,,'old Q3 Data'!$B$2:$B$5,1)",
                           "=SERIES(,,'Q3 Data'!$B$2:$B$5,1)"),
                   Repoint("chart", "Chart1", "1", "=SERIES(,,Data_Q3_old[Value],1)", "=SERIES(,,Data_Q3[Value],1)"),
                   Repoint("name", "TotalQ", "4", "='old Q3 Answer'!$E$12", "='Q3 Answer'!$E$12"),
                   Repoint("name", "Summary!Local", "7", "=Data_Q3_old[Value]", "=Data_Q3[Value]")]
    assert unread == [("Locked", "a protected sheet")]
    assert [r.where_text for r in got] == ["Summary!A1", "Summary!D1:D3", "a chart series on Summary",
                                          "a chart series on Chart1", "the name TotalQ", "the name Summary!Local"]
