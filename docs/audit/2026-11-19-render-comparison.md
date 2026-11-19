# Live comparison of the two renderers (phase 3, Task 6)

Date: 19.11.2026, on the owner's PC (Excel 16.0, nb-NO). The parity gate is `scripts/compare_renders.py`. It draws
the nine test-set cases twice, once with `SV_RENDER=com` (pipeline.excel_job, the reference) and once with
`SV_RENDER=file` (pipeline.file_job). Each renderer draws into its own new workbook in the analyser folder. The
three cases without a chart block run as `bench_answers.py` runs them. The script then reads every Q sheet pair back
through the user's Excel and diffs them. It prints counts, addresses, property names and format values. It never
prints a cell value, and it prints a formula only when the cell holds one.

Command of the final run (from the repo root):

    engine\.venv\Scripts\python.exe scripts\compare_renders.py

Exit code of the final run: **1**. Three difference groups are outside the allowed list. They are listed under
"Proposed allowances" below, and the script keeps failing on them until the controller rules.

## The two workbooks of the final run (left open, to compare by eye)

- COM renderer: `2026-11-19_1952 Wholesale P&L year to date vs last year.xlsx`
- File renderer: `2026-11-19_1958 Wholesale P&L year to date vs last year.xlsx`

Both are in the analyser folder and hold Q1 to Q9 (b01 to b16, in the test-set order). The script closed the six
workbooks of the earlier iteration runs after it had read them.

## Final output

```
cases compared: 9; cells per side: 148197; texts with the run's timestamp masked: 18; failed reads: 0
difference groups: 51, not allowed: 3
DIFF | Answer | chart height | com 260.0 -> file 260.2 | 9 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 chart at $B$8, b02 chart at $B$7, b05 chart at $B$10, b06 chart at $B$10, b07 chart at $B$7, b09 chart at $B$10 ...
DIFF | Answer | chart overlap | com 100 -> file 0 | 1 in b07: b07 chart at $B$7
DIFF | Data | column widths | com 39.38 -> file 37.13 | 1 in b09: b09 B
allowed (the hidden sheet name smartass_pagesetup) | Answer | sheet name | com ["'Q# Answer'!smartass_pagesetup", '="a1|$A$1:$H$86|Intern|Source: Oracle EPM Cloud · fetched 2026-11-19 · page &P of &N|breaks 26,76"', False] -> file None | 1 in b01: b01 'Q# Answer'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Answer | sheet name | com ["'Q# Answer'!smartass_pagesetup", '="a1|$A$1:$G$50|Intern|Source: Oracle EPM Cloud · fetched 2026-11-19 · page &P of &N|breaks 29"', False] -> file None | 2 in b02 b13: b02 'Q# Answer'!smartass_pagesetup, b13 'Q# Answer'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Answer | sheet name | com ["'Q# Answer'!smartass_pagesetup", '="a1|$A$1:$I$64|Intern|Source: Oracle EPM Cloud · fetched 2026-11-19 · page &P of &N|breaks 28,57"', False] -> file None | 1 in b05: b05 'Q# Answer'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Answer | sheet name | com ["'Q# Answer'!smartass_pagesetup", '="a1|$A$1:$H$64|Intern|Source: Oracle EPM Cloud · fetched 2026-11-19 · page &P of &N|breaks 28,51"', False] -> file None | 1 in b06: b06 'Q# Answer'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Answer | sheet name | com ["'Q# Answer'!smartass_pagesetup", '="a1|$A$1:$F$59|Intern|Source: Oracle EPM Cloud · fetched 2026-11-19 · page &P of &N|breaks 25,49"', False] -> file None | 1 in b07: b07 'Q# Answer'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Answer | sheet name | com ["'Q# Answer'!smartass_pagesetup", '="a1|$A$1:$G$76|Intern|Source: Oracle EPM Cloud · fetched 2026-11-19 · page &P of &N|breaks 28,51"', False] -> file None | 1 in b09: b09 'Q# Answer'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Answer | sheet name | com ["'Q# Answer'!smartass_pagesetup", '="a1|$A$1:$G$57|Intern|Source: Oracle EPM Cloud · fetched 2026-11-19 · page &P of &N|breaks 28,52"', False] -> file None | 1 in b15: b15 'Q# Answer'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Answer | sheet name | com ["'Q# Answer'!smartass_pagesetup", '="a1|$A$1:$F$68|Intern|Source: Oracle EPM Cloud · fetched 2026-11-19 · page &P of &N|breaks 28"', False] -> file None | 1 in b16: b16 'Q# Answer'!smartass_pagesetup
allowed (the cell style name) | Answer | style | com 'TN Tittel' -> file 'Normal' | 84 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 A1, b01 B1, b01 C1, b01 D1, b01 E1, b01 F1 ...
allowed (the cell style name) | Answer | style | com 'TN Undertittel' -> file 'Normal' | 84 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 A2, b01 B2, b01 C2, b01 D2, b01 E2, b01 F2 ...
allowed (the cell style name) | Answer | style | com 'TN Tom' -> file 'Normal' | 2392 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 A3, b01 B3, b01 C3, b01 D3, b01 E3, b01 F3 ...
allowed (the cell style name) | Answer | style | com 'TN Seksjon' -> file 'Normal' | 538 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 A4, b01 B4, b01 C4, b01 D4, b01 E4, b01 F4 ...
allowed (the cell style name) | Answer | style | com 'TN Tekst' -> file 'Normal' | 220 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 B5, b01 C5, b01 D5, b01 E5, b01 F5, b01 G5 ...
allowed (the cell style name) | Answer | style | com 'TN Kolonne' -> file 'Normal' | 196 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 A27, b01 B27, b01 C27, b01 H27, b01 I27, b01 J27 ...
allowed (the cell style name) | Answer | style | com 'TN Kolonne tall' -> file 'Normal' | 80 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 D27, b01 E27, b01 F27, b01 G27, b01 D32, b01 E32 ...
allowed (the cell style name) | Answer | style | com 'TN Etikett' -> file 'Normal' | 659 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 B28, b01 C28, b01 H28, b01 I28, b01 J28, b01 K28 ...
allowed (the cell style name) | Answer | style | com 'TN Heltall' -> file 'Normal' | 263 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 D28, b01 E28, b01 F28, b01 D29, b01 E29, b01 F29 ...
allowed (the cell style name) | Answer | style | com 'TN Prosent 1d' -> file 'Normal' | 69 in b01 b02 b05 b06 b09 b13 b15 b16: b01 G28, b01 G29, b01 G33, b01 G34, b01 G36, b01 G38 ...
allowed (the cell style name) | Answer | style | com 'TN Etikett · Sum' -> file 'Normal' | 46 in b01 b05 b09 b16: b01 B35, b01 C35, b01 I35, b01 J35, b01 K35, b01 L35 ...
allowed (the cell style name) | Answer | style | com 'TN Heltall · Sum' -> file 'Normal' | 33 in b01 b05 b09 b16: b01 D35, b01 E35, b01 F35, b01 H35, b01 D37, b01 E37 ...
allowed (the cell style name) | Answer | style | com 'TN Prosent 1d · Sum' -> file 'Normal' | 4 in b01 b09 b16: b01 G35, b01 G37, b09 G33, b16 E42
allowed (the cell style name) | Answer | style | com 'TN Etikett · Hovedsum' -> file 'Normal' | 27 in b01 b05 b06 b07 b09: b01 B39, b01 C39, b01 I39, b01 J39, b01 K39, b01 L39 ...
allowed (the cell style name) | Answer | style | com 'TN Heltall · Hovedsum' -> file 'Normal' | 35 in b01 b05 b06 b07 b09: b01 D39, b01 E39, b01 F39, b01 H39, b01 D49, b01 J49 ...
allowed (the cell style name) | Answer | style | com 'TN Prosent 1d · Hovedsum' -> file 'Normal' | 4 in b01 b06 b09: b01 G39, b06 G41, b06 H41, b09 G40
allowed (the cell style name) | Answer | style | com 'TN Tekst tallkolonne' -> file 'Normal' | 11 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 D52, b02 D30, b05 D38, b06 D45, b06 D46, b07 D40 ...
allowed (the cell style name) | Answer | style | com 'TN Note' -> file 'Normal' | 881 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 A55, b01 B55, b01 C55, b01 D55, b01 E55, b01 F55 ...
allowed (the cell style name) | Answer | style | com 'TN Underseksjon' -> file 'Normal' | 45 in b09: b09 B30, b09 C30, b09 D30, b09 E30, b09 F30, b09 G30 ...
allowed (Data column width within 1.0 character) | Data | column widths | com 6.88 -> file 6.75 | 9 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 E, b02 E, b05 E, b06 E, b07 E, b09 F ...
allowed (Data column width within 1.0 character) | Data | column widths | com 8.63 -> file 8.5 | 8 in b01 b02 b06 b07 b09 b13 b15 b16: b01 F, b02 F, b06 F, b07 F, b09 G, b13 F ...
allowed (Data column width within 1.0 character) | Data | column widths | com 7.75 -> file 7.63 | 9 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 G, b02 G, b05 G, b06 G, b07 G, b09 H ...
allowed (Data column width within 1.0 character) | Data | column widths | com 8 -> file 7.88 | 9 in b01 b02 b05 b06 b07 b09 b13 b15 b16: b01 H, b02 H, b05 H, b06 H, b07 H, b09 I ...
allowed (Data column width within 1.0 character) | Data | column widths | com 9 -> file 8.63 | 2 in b01 b09: b01 I, b09 J
allowed (Data column width within 1.0 character) | Data | column widths | com 8.88 -> file 8.5 | 1 in b02: b02 I
allowed (Data column width within 1.0 character) | Data | column widths | com 18.88 -> file 18.75 | 1 in b05: b05 C
allowed (Data column width within 1.0 character) | Data | column widths | com 10.38 -> file 10.13 | 1 in b05: b05 F
allowed (Data column width within 1.0 character) | Data | column widths | com 9.88 -> file 9.5 | 3 in b05 b06 b15: b05 I, b06 I, b15 I
allowed (Data column width within 1.0 character) | Data | column widths | com 10.13 -> file 9.75 | 1 in b07: b07 I
allowed (Data column width within 1.0 character) | Data | column widths | com 8 -> file 7.63 | 1 in b13: b13 I
allowed (Data column width within 1.0 character) | Data | column widths | com 6.25 -> file 6.13 | 1 in b16: b16 I
allowed (the hidden sheet name smartass_pagesetup) | Data | sheet name | com ["'Q# Data'!smartass_pagesetup", '="a1|$A$1:$I$134|Intern|Source: Oracle EPM Cloud · &A · page &P of &N"', False] -> file None | 1 in b01: b01 'Q# Data'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Data | sheet name | com ["'Q# Data'!smartass_pagesetup", '="a1|$A$1:$I$32|Intern|Source: Oracle EPM Cloud · &A · page &P of &N"', False] -> file None | 1 in b02: b02 'Q# Data'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Data | sheet name | com ["'Q# Data'!smartass_pagesetup", '="a1|$A$1:$I$194|Intern|Source: Oracle EPM Cloud · &A · page &P of &N"', False] -> file None | 1 in b05: b05 'Q# Data'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Data | sheet name | com ["'Q# Data'!smartass_pagesetup", '="a1|$A$1:$I$246|Intern|Source: Oracle EPM Cloud · &A · page &P of &N"', False] -> file None | 1 in b06: b06 'Q# Data'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Data | sheet name | com ["'Q# Data'!smartass_pagesetup", '="a1|$A$1:$I$34|Intern|Source: Oracle EPM Cloud · &A · page &P of &N"', False] -> file None | 2 in b07 b13: b07 'Q# Data'!smartass_pagesetup, b13 'Q# Data'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Data | sheet name | com ["'Q# Data'!smartass_pagesetup", '="a1|$A$1:$J$641|Intern|Source: Oracle EPM Cloud · &A · page &P of &N"', False] -> file None | 1 in b09: b09 'Q# Data'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Data | sheet name | com ["'Q# Data'!smartass_pagesetup", '="a1|$A$1:$I$228|Intern|Source: Oracle EPM Cloud · &A · page &P of &N"', False] -> file None | 1 in b15: b15 'Q# Data'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Data | sheet name | com ["'Q# Data'!smartass_pagesetup", '="a1|$A$1:$I$7|Intern|Source: Oracle EPM Cloud · &A · page &P of &N"', False] -> file None | 1 in b16: b16 'Q# Data'!smartass_pagesetup
allowed (the hidden sheet name smartass_pagesetup) | Lines b0 | sheet name | com ["'Q# Lines b0'!smartass_pagesetup", '="a1|$A$1:$AS$2843|Intern|Source: Oracle EPM Cloud · &A · page &P of &N|zoom100|$4:$4"', False] -> file None | 1 in b16: b16 'Q# Lines b0'!smartass_pagesetup
```

All 9 cases drew in both renderers, and sheet = model was OK in all 18 answers. The replies did not differ: keys,
sheets, chart titles, layout findings, warnings (numbers masked) and the status matched.

Run artefacts are not counted as differences, and none was added to the allowed list:
- The fetch timestamp (`datasheet.fetched_text`) sits in the Data header's Fetched cell. 18 texts were masked
  before hashing, 9 per side.
- Q numbers in sheet, table and name references are written as `Q#`. They are the same on both sides in this run.
- Sheet positions are not compared, because the two answers are in separate workbooks (controller ruling).

## Fixes made

| commit | what differed (COM -> file) | how much (first seen) |
|---|---|---|
| a170ada | Clustered column charts: Overlap -27 (AddChart2's default) -> 0. | the 3 column charts (b02, b05, b13) |
| b6ced67 | The file was built in XlsxWriter's Calibri 11 (7 px digits, 15 pt rows, a theme font). A copied sheet keeps its widths in pixels, and its unformatted cells take the destination's Normal style (Telenor 11, 8 px digits, 14.25 pt rows). Results: Answer column widths 16 -> 14, 46 -> 40.25 and 2 -> 1.75; Data rows 14.25 -> 15 pt; blank cells in the theme font (Aptos Narrow); the style "Normal 2". The fix: `movein.base` reads the destination's Normal style and standard sizes (one COM call), and `xlsx.build` writes them as the file's defaults. This also closes candidates (a), Lines A1/A2 in XlsxWriter's font, and (e), blank cells. | b02 alone: 7 Answer widths, 4 Data widths, 31 Data rows, 69 blank Data cells; every case alike |
| e93fbb4 | Clustered bar charts: Overlap 0 -> -27 (a170ada went too far: AddChart2's bar keeps 0). Data labels: #404040 (AddChart2's label text) -> #595959 (the axis grey). | 2 bar charts, each with one series (b06, b16); the labels of 5 series and of 3 labelled line points in the 4 cases re-read (b01, b02, b06, b07) |
| 7025aa9 | Thin waterfall steps: COM lifts a label only when its Top is more than the lift. The LabelFix lifted the empty points of the other step series too (Position Custom against Center). | 9 point labels (b01, b09) |
| 93774d3 | Data widths: the file's text widths are 7 px units, and AutoFit works in the destination's 8 px digits and measures the header bold, so b05 C came out 21.38 against 18.88. This is candidate (b). A Data table without facts had its empty body row formatted ('@', the value format, blue). The Lines date was a literal 'dd.mm.yyyy'. | 3 Data widths beyond 1 character (b05 C, F; b09 E); 9 cells (b16) |
| ac6c117 | Lines date: drill_xl's local 'dd.mm.åååå' is stored by Excel as its built-in short date 14 (read in en-US: 'm/d/yyyy'). The file wrote a code. `movein.date_format` decides it from Application.International in base's one call. | 5678 cells (b16) |
| 55f99c2 | Excel's merge gives the merged cells the bottom border of the cells above as their own top border. The first Key point under its section heading therefore had a navy top border in the COM drawing and none in the file. It is the same line on screen, and the cells now carry the same borders. | 54 cells (all nine cases) |
| e11d92f | A Data table without facts: `table_font` covers the whole table range, so the empty body row is Arial 9. After 93774d3 the file left it in the Normal style. The separator row above the table (A5:D5, Arial in the COM drawing) matched as well after this fix. I did not trace why. | 13 font-name and 12 font-size cells (b16) |

Every fix has unit tests (engine/tests/v2/test_xlsx.py, test_xlsx_charts.py, test_file_job.py and
engine/tests/test_movein.py). The COM writers are unchanged.

Candidate (d), the file job reports no style lint: no difference showed. The replies' warnings matched in all nine
cases, because apply_report found nothing to lint in any of them. The file job would still miss a style lint finding
if apply_report ever reported one.

## Proposed allowances (the script still fails on them)

1. **Chart height 260.25 pt against 260 (all 9 charts).** XlsxWriter sizes a chart in whole pixels: 260 pt is
   346.67 px, and XlsxWriter writes 347 px. The difference is a third of a pixel, so it cannot be seen. The file
   cannot hold 260 pt exactly. The alternative to an allowance is one COM call per chart after the move-in
   (`ChartObject.Height = spec.height` in movein.finish). That would add a call to a job whose point is fewer calls.
   Proposed rule: chart width and height within 0.75 pt (one pixel).
2. **Overlap of a line chart's group: 100 against 0 (b07).** AddChart2 leaves 100 in the COM chart's line group. The
   file format has no overlap for a line chart (`c:lineChart` has no `c:overlap`), so the file cannot write it, and
   a line has no bars for it to act on. Proposed rule: Overlap and GapWidth are not compared on a line chart group.
3. **Data column B of b09: 37.13 against AutoFit's 39.38 (2.25 characters).** The column holds one 48-character
   query id in lowercase with underscores. `tn.text_units` measures it at about 297 px, and AutoFit gives the column
   315 px. A per-character model cannot know the screen's glyph hinting, and every other Data width of the nine
   cases is within 1.0 character. Options: a wider tolerance for texts beyond 40 characters, or AutoFit through COM
   after the move-in (one call per Data column, against the call budget).

## How the script reads (notes for the reviewer)

- **Borders** come from the range as Excel's XML Spreadsheet (`Range.Value(11)`, one call per sheet): each cell's
  own style, with the style's parent and the row and column styles. Range reads of borders are not reliable, so they
  cannot be split like the other properties (measured on b07, 19.11.2026):
  - a multi-cell row's `Borders(edge)` reads a style while some of its cells have none;
  - `xlInsideHorizontal` reads none over rows that have borders, and a style over rows that mostly have none.
  Reading every cell on its own works, but takes too long for the 128 000 cells of b16's Lines sheet. The first full
  run used the range reads and missed the 54 cells of 55f99c2.
- **Other cell properties** are read once per uniform rectangle (a mixed range reads None), splitting columns first.
  On b02 and b07, `--exact` (every cell on its own) gave 0 differences against the split reads for every property
  except the borders, which led to the XML read above.
- The Lines sheet of b16 (A1:AS2843) reads in 7 to 8 s per side; the whole read of both workbooks takes about 7 min.
- The reads go straight to the IDispatch with cached DISPIDs and LCID 0x409, one cross-process call each. A dynamic
  pywin32 object took about 15 ms per read. Number formats therefore read in en-US ('General', 'm/d/yyyy').
- Reading activates each Q sheet of the two workbooks to get the freeze panes and gridlines. ScreenUpdating and the
  active sheet are put back after each case. Nothing is written.
- Each draw child records its workbook in `sessions.json` under its own process id and the `_last:` entry of the
  worktree folder, as bench_answers.py does. The owner's own sessions in other folders are not touched.
- The dumps (rounded numbers) live in a temp folder that the script deletes. The iteration dumps of this task were
  deleted after use.

Not compared: the Sources and _svmeta sheets, conditional formats, the fills of single highlighted points, how far a
thin step's label is lifted, legend entries, and the plot area's size.
