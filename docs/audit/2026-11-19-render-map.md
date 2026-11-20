# Render map: how an answer reaches Excel today (for the phase 3 plan)

Read-only survey of branch `main` (HEAD cea6119, same commit as `feat/light-fast-p0`), 2026-11-19. Every reference is
`file:line` under `engine/smartass/` unless the path says otherwise (tests: `engine/tests/...`). "COM call" means one
`xl.jobs.com(...)` call (xl/jobs.py:48), which is what the stage timings count (core/timing.py:83).
Colours are given as RGB hex. The code mostly holds them as Excel BGR integers (`tn.rgb` = r + g*256 + b*65536,
xl/vendor/tn_rapportstil.py:30), so `0xC00000` in the code is **blue** #0000C0, and `0x64381F` is navy #1F3864.
Where I could not verify something from the repo (mostly XlsxWriter's abilities) it says **to verify**.

## 0. The ten-line summary

- One answer = one user-Excel job (`pipeline.excel_job`, pipeline.py:539) or, for a follow-up/refresh, `redraw_job`
  (pipeline.py:558). Both run `_draw` (pipeline.py:458), which writes Data, Lines, Answer, charts, reads back, writes
  the status cell, appends to the workbook-wide Sources log, stores the spec in `_svmeta`, orders the sheets, checks.
- The only pure description is the Answer's `SheetPlan` (render/layout.py:107). The Data sheet has only a pure header
  (`datasheet.Header`, render/datasheet.py:26); the Lines sheets have **no** plan at all (xl/drill_xl.py:207 writes them
  directly). The look (styles, heights, merges, number-format codes, conditional formats, page setup) lives in
  `xl/styling.apply_report` (xl/styling.py:347), not in the plan.
- Several parts of the look are derived from what Excel computed: row heights/merges of the Key points read the
  formula **results** (xl/styling.py:374), Data and Sources column widths come from `AutoFit` (render/excel.py:214,
  xl/sources.py:179), one chart label position reads `DataLabel.Top` (render/charts.py:218).
- No `fast_excel` (screen updating/calculation/events off) around the Data, Answer, styling or chart writes; only the
  Lines sheet (xl/drill_xl.py:222), the Sources log (xl/sources.py:128) and sheet moves (xl/sheetorder.py:185) use it.
- The Sources sheet and `_svmeta` are workbook-level and depend on what is already in the workbook. Design §2 does not
  mention the Sources log at all; it names "store the spec in `_svmeta`" as a step, but the spike did not cover it.


## 1. The sequence

### 1.0 How the job is called

- `fin_analyze(spec=...)` (server.py:252) -> `_new_question` (server.py:244): `guard_chart`, `guard_key_points`, then
  `_prepare` (server.py:173; drills lines blocks first, server.py:151-166, alias `b<i>` = block index, server.py:163;
  fetches grids in a thread; the ad hoc fallback in the user's Excel when grids failed, server.py:193-198), then
  `_draw_job(lambda h: excel_job(h, pre, target), ...)` (server.py:248).
- `_draw_job` (server.py:134-142) -> `_user` (server.py:130) -> `serverkit._call_user` (core/serverkit.py:215) ->
  `ExcelHost.call(fn, user_excel=True)` (xl/host.py:117). A user-Excel job gets its own STA thread
  (xl/host.py:168-174), a `JobToken` (abandoned on timeout: its next COM call raises, xl/jobs.py:42-45), the timeout
  `config.USER_CALL_TIMEOUT_S` = 600 s (core/config.py:178), and releases the per-Excel turn locks when it ends
  (xl/host.py:163-166). `with_timings` (pipeline.py:525) wraps the job so its `Timer` summary (stage seconds +
  "com calls.<stage>" counts) travels back; `_draw_job` adds it to the run log as `excel job.*` (server.py:138-141).
  `scripts/bench_answers.py` reads exactly these keys (scripts/bench_answers.py:33, 110-112: `excel job._total`,
  `excel job.draw`, `...draw.data sheet`, `...draw.sources`, `...draw.answer`, `excel job.save`). A new writer must
  keep (or deliberately rename) these stage names, because the phase 3 gate is read from them.
- Every COM call is `jobs.com(call, limit=60, check=True)` (xl/jobs.py:48): busy-retry, abort check, a readable
  `HostError(BUSY)`.

### 1.1 What "t" is

`t` is an `xl.target.Target` (xl/target.py:21-60):

```
@dataclass
class Target:
    app: object          # the user's Excel.Application (COM)
    wb: object           # the target Workbook (COM)
    pid: int; hwnd: int
    bridge: object = None            # the svbridge_<hash>.xlam workbook (Smart View VBA bridge)
    created: bool = False            # our own analysis file (saved after each answer)
    sheets_added: list[str] = field(default_factory=list)   # what rollback deletes
    name -> wb.Name;  path -> instances.local_path(wb.FullName) (OneDrive URL -> synced local path);  full_name
    run(fn, *args) -> bridge macro;  status(text) -> Application.StatusBar
```

### 1.2 `excel_job(h, pre, target="session")` (pipeline.py:539-555), step by step

1. **stage "open target"**: `_target(h, target, pre.spec.question[:60])` (pipeline.py:211) ->
   `xl.target.resolve_target(h, target, allow_archive=False, title=...)` (xl/target.py:99):
   - always first `sso.ensure_global_watcher()` (xl/target.py:105) (SSO clicker; slated for deletion);
   - `"session"`: `session.workbook()` path -> resolved as a path, then `t.created = True` (xl/target.py:106-111); no
     session workbook yet -> `"new"`;
   - a bare name -> `_find_named` over every Excel (xl/target.py:260), `hold_turn`, `_finish_target`;
   - an absolute path: archive refusal before anything is opened (xl/target.py:123-124), then `_open_path`
     (xl/target.py:312): already open anywhere / busy holder wait / file locked / `Workbooks.Open` in the user's running
     Excel / start `EXCEL.EXE /x path`;
   - `"active"`: the first Excel's `ActiveWorkbook` that is not an .xlam (xl/target.py:128-137), else `"new"`;
   - `"new"`: `_new_target` (xl/target.py:151): a free Excel (or a spawned one on the TN template, max
     `jobs.MAX_EXCELS` = 4), `Workbooks.Add(<tn_template_<ver>.xltx>)` (template from `styling.ensure_template`,
     xl/styling.py:157), `_save_new` (xl/target.py:191): first sheet renamed **"Sources"**, sensitivity label
     (`instances._label`, xl/instances.py:272), `SaveAs(analyser\<project>\<YYYY-MM-DD_HHMM> <title>.xlsx, 51)`; with no
     Excel running: a copy of the labelled seed file (`_new_analysis_file`, xl/target.py:398). Then
     `session.set_workbook(t.path)` (xl/target.py:143);
   - `_finish_target` (xl/target.py:251-257): wait for an EPM login window to clear, archive check again,
     `_wait_ready` (`app.Ready`, edit mode), **`_attach_bridge`** (xl/target.py:79): builds/opens
     `svbridge_<hash>.xlam` in the user's Excel and calls `B_Ping` (needs the hidden Excel to build it the first time,
     xl/instances.py:236-248). The answer job itself never uses the bridge.
2. `with autosave_paused(t) as notes` (pipeline.py:260-281): reads `wb.AutoSaveOn` (exception -> treated as off), sets
   it False; on exit switches it back on only for our own file (`_is_ours`), else appends the note "AutoSave was
   switched off in <name> ...".
3. `q = _next_q(t)` (pipeline.py:216) -> `sheets.next_question` (xl/sheets.py:30): reads every sheet name; max of
   `question_of` (xl/sheets.py:21) + 1. `question_of` matches `_Q = ^Q(\d+) (Answer|FinMax|FinLite|FinPlan|Lines)\b`
   (xl/sheets.py:16), or the topic pattern `^Q(\d+) [A-Z][A-Za-z0-9 &/-]{1,24}$` (xl/sheets.py:18) **only when
   "Q<n> Answer" exists**. So "Q<n> Data" only counts through the topic rule, and a user's "Q7 Forecast" counts when a
   "Q7 Answer" exists.
4. `ws_answer = _add_sheet(t, "Q<q> Answer", q)` (pipeline.py:545, 221) -> `sheets.add_sheet(t, name, after_q=q)`
   (xl/sheets.py:114-135): anchor = last sheet of question q, else the last sheet that is not Sources/Sources SV/
   _svmeta, else before the first sheet; positional `Worksheets.Add(None, anchor)` (keyword `After=` is ignored through
   late binding, xl/sheets.py:125-127); name + optional tab colour in one call; **appended to `t.sheets_added`**.
5. `ws_data = _add_sheet(t, "Q<q> Data", q, DATA_TAB)` (pipeline.py:546); `DATA_TAB = 0xA6A6A6` (grey #A6A6A6,
   pipeline.py:208).
6. **stage "draw"**: `_draw(t, q, pre, ws_answer, ws_data)` (pipeline.py:458-497):
   1. `fetched = datetime.now().astimezone()`, `user = datasheet.windows_user()`,
      `header = header_rows(q, spec, pre, fetched, user)` (render/datasheet.py:113): one timestamp per draw.
   2. **stage "data sheet"**: `write_data(ws_data, pre.frame, q, header)` (render/excel.py:43) (redraw:
      `rewrite_data`, render/excel.py:120). Writes the table, then the header, then the look (§2.2, §4.4).
   3. For each line set (`pre.frame.lines`, keyed `b<i>`): **stage "lines"**: `_write_lines(t, q, ls,
      fetched_text)` (pipeline.py:367) - before the Answer "so its formulas resolve" (pipeline.py:470) (§6).
   4. `free, too_many = free_points(pre.spec)` (render/keypoints.py:344); warnings for > 5 points.
   5. `plan = layout(q, spec, result, stock=_stock(pre), subtitle=answer_subtitle(...), const=constant_columns(frame),
      caveats=pre.caveats, key_points=partial(key_points, stock=..., free=free))` (pipeline.py:475-478); then
      `plan.footer = answer_footer(fetched)` = `"Source: Oracle EPM Cloud · fetched YYYY-MM-DD · page &P of &N"`
      (render/excel.py:313-315).
   6. redraw only: `_make_room(ws_answer, old_owned, {"rows": plan.last_row, "cols": col_number(plan.last_col)})`
      (pipeline.py:480-481) -> `owned.make_room` (xl/owned.py:25) (§8).
   7. **stage "answer"**: `write_answer(ws_answer, plan, pre.warnings)` (render/excel.py:267-272):
      `answer_grid(plan)` (render/excel.py:255) builds ONE rectangle from the smallest used row/column (in practice B1)
      to the largest, "" in empty cells, text prefixed with an apostrophe (render/excel.py:261-262); one
      `Range.Formula2 = grid` (render/excel.py:271); then `styling.apply_report(ws, answer_roles(plan))`
      (render/excel.py:272, 275) -> returns the style lint (§4).
   8. **stage "charts"**: `write_charts(ws_answer, plan, drawn=drawn)` (render/charts.py:334) -> chart lint +
      one line per failed chart; `drawn` gets each drawn title (§5).
   9. `_calculate(t)` (pipeline.py:226-230): `app.Calculate()` (**the whole application**, every open workbook), then
      `sheets.wait_calculated(app)` (xl/sheets.py:316) polls `CalculationState == 0` every 0.3 s for up to 60 s.
   10. `mismatches = compare(plan.addr, pre.result.cells, read_block(ws_answer, plan), plan.fractions)`
       (pipeline.py:488; render/readback.py:23; `read_block` = one `Value2` read of `A1:<last_col><last_row>`,
       render/excel.py:332-335).
   11. `write_status(ws_answer, plan, mismatches)` (render/excel.py:323-329): `"OK"` or `"DIFF in <n> cells"` as
       `Value2` into `plan.status_cell`.
   12. **stage "sources"**: `logged = _write_sources(t, q, pre, ws_data, header, fetched_text)` (pipeline.py:424-446):
       reads each query's Data status cell `.Value` (OK/DIFF, calculated), each `'Q<q> Lines <alias>'!I3` tie-out,
       then `sources.write_sources(...)` (xl/sources.py:113) (§6). Never raises: a failure is a warning text.
   13. `_save_spec(t, q, pre.raw, {"rows": plan.last_row, "cols": col_number(plan.last_col)})` (pipeline.py:233-238):
       creates `_svmeta` when missing (`Worksheets.Add(None, last sheet)` + `_name_hidden`: name + `Visible = 0`,
       xl/sheets.py:181-183) - **not added to `t.sheets_added`** - then `sheets.save_spec` (xl/sheets.py:191) (§8).
   14. `arranged = _arrange(t, "Q<q> Answer")` (pipeline.py:241-250) -> `sheetorder.arrange_sheets(t.wb, "full" if
       _is_ours(t) else "active", activate=...)` (xl/sheetorder.py:163); a failure is a warning, never raised.
   15. **stage "read back"**: `found = _check(plan, ws_answer, ws_data, sources_ok, header)` (pipeline.py:452-455) ->
       `render.check.check` (render/check.py:162) (§7).
   16. `return reply(pre, q, str(t.path), ["Q<q> Answer", "Q<q> Data", *line_sheets], mismatches, lint + [logged,
       arranged], layout=found, charts=drawn)` (pipeline.py:496-497; `reply` at pipeline.py:171).
7. **stage "save"**: `_save(t)` (pipeline.py:284-287): `wb.Save()` only when `_is_ours(t)` (pipeline.py:253-257:
   `t.created` or `serverkit._is_our_analysis_file(t.path)`, core/serverkit.py:206-212 = the file sits in
   `config.ANALYSER_DIR` or one folder below it and is not in the archive).
8. On any exception inside the `try`: `_rollback(t)` (pipeline.py:290) -> `sheets.rollback(t)` (xl/sheets.py:44-68):
   DisplayAlerts off, delete every name in `t.sheets_added` in reverse, DisplayAlerts back, StatusBar cleared; then
   re-raise. Sheets it deletes: Answer, Data, any new Lines sheet (drill_xl.add_sheet), a new Sources sheet
   (xl/sources.py:62-64). Not `_svmeta`, and not the Sources rows already appended to an existing log.
9. After the AutoSave context: `remember(pre, out)` (pipeline.py:500) stores the run in answers.db (warning on
   failure), `_with_notes(out, notes)` (pipeline.py:519) adds the AutoSave note.

### 1.3 `redraw_job(h, pre, q, owned, target)` (pipeline.py:558-596)

Called by a follow-up (server.py:268-273) and by `fin_refresh` (server.py:296-307) with `target = path`, the absolute
path `read_stored` returned - so `t.created` is False and "ours" depends only on `_is_our_analysis_file`.

1. stage "open target": `_target(h, target, ...)` as above.
2. `autosave_paused(t)`.
3. `ws_data = _sheet(t, "Q<q> Data")` (pipeline.py:327: `wb.Worksheets(name)`; a missing sheet is a COM error).
4. `old_data = _read_data(ws_data, q)` -> `render.excel.read_data` (render/excel.py:82-88): `ListObjects("Data_Q<q>")
   .Range.Value2` (None without the table).
5. `b_answer = _backup(t, "Q<q> Answer")` (pipeline.py:295) -> `owned.backup` (xl/owned.py:56-64):
   `ws.Copy(None, ws)` (a same-workbook copy of the Answer sheet, charts included), the copy renamed
   `unique_sheet_name("~Q<q> Answer")` and hidden (`Visible = 0`).
6. try: `ws_answer = _sheet(t, a_name)`; `_clear(ws_answer, owned)` (pipeline.py:338) -> `owned.clear`
   (xl/owned.py:40-47): with a stored owned range `Range(A1, Cells(rows, cols)).Clear()`, else `Cells.Clear()`; then
   `drop_charts` (xl/owned.py:50-53): charts whose name starts with "smartass chart" (all charts when no owned range).
7. stage "draw": `_draw(..., redraw=True, old_owned=owned)` - same as 1.2 step 6, except `rewrite_data` instead of
   `write_data` (the existing table is moved/resized and rewritten in place, never deleted or renamed:
   render/excel.py:91-147) and `_make_room` before the Answer is written.
8. except: three independent restore steps, each in its own try (pipeline.py:575-587): `_rollback(t)`;
   `_restore(t, a_name, b_answer)` (pipeline.py:300-311: if the Answer still exists `owned.restore` deletes it and
   un-hides/renames the backup, else only un-hide + rename); if `old_data` `rewrite_grid(_sheet(t, d_name),
   old_data, q)` (render/excel.py:91). Any failing step -> `RestoreError(original; the restore was incomplete - ...)`
   (pipeline.py:30, 585-586), which makes `fin_refresh` skip the remaining questions (server.py:299-302, 311-313).
   **Not restored**: Lines sheets rewritten in place, Sources rows appended, the `_svmeta` row (if `_save_spec` ran).
9. success: `_drop_backup(t, b_answer)` (pipeline.py:319 -> `_delete_sheet` -> `sheets.drop_parked` with alerts off,
   xl/sheets.py:84-91); stage "save" `_save(t)`; `out["followup"] = True`; when `owned` was None a warning "old answer
   without a stored range: the whole Answer sheet was redrawn" (pipeline.py:592-594).
10. `remember`, `_with_notes`.

### 1.4 The small helpers, in one table

| helper | file:line | what it does in Excel |
|---|---|---|
| `_target` | pipeline.py:211 | `resolve_target(..., allow_archive=False)` (xl/target.py:99) |
| `_next_q` | pipeline.py:216 | sheet names -> next q (xl/sheets.py:30) |
| `_add_sheet` | pipeline.py:221 | `Worksheets.Add` after the q group, name, tab colour, `t.sheets_added` (xl/sheets.py:114) |
| `_calculate` | pipeline.py:226 | `Application.Calculate()` + poll `CalculationState` (xl/sheets.py:316) |
| `_save_spec` | pipeline.py:233 | create hidden `_svmeta` if missing; `sheets.save_spec` row `SPEC Q<q>` (xl/sheets.py:191) |
| `_arrange` | pipeline.py:241 | `sheetorder.arrange_sheets` full/active; returns a warning text on failure |
| `_is_ours` | pipeline.py:253 | `t.created` or file in the analyser folder (core/serverkit.py:206) |
| `autosave_paused` | pipeline.py:260 | `AutoSaveOn` off during the job; back on only for our file |
| `_save` | pipeline.py:284 | `wb.Save()` for our file only |
| `_rollback` | pipeline.py:290 | delete `t.sheets_added` (xl/sheets.py:44) |
| `_backup` | pipeline.py:295 | hidden same-workbook copy "~Q<q> Answer" (xl/owned.py:56) |
| `_restore` | pipeline.py:300 | backup becomes the Answer sheet again |
| `_delete_sheet` / `_drop_backup` | pipeline.py:314, 319 | `Worksheets(name).Delete()` with alerts off |
| `_read_data` | pipeline.py:323 | Data_Q<q> as a grid (render/excel.py:82) |
| `_sheet` | pipeline.py:327 | `wb.Worksheets(name)` |
| `_make_room` | pipeline.py:332 | insert rows/columns before a larger redraw would cover user cells (xl/owned.py:25) |
| `_clear` | pipeline.py:338 | clear the owned block + our charts (xl/owned.py:40) |
| `read_stored` | pipeline.py:348 | `(_stored_specs(t), ActiveSheet.Name, str(t.path))` - a short job of its own |
| `_write_lines` / `_lines_header` | pipeline.py:367, 403 | Lines sheet (§6) |
| `_write_sources` | pipeline.py:424 | Sources log (§6) |
| `_check` | pipeline.py:452 | layout check (§7) |


## 2. The sheets of one answer Q<n>

| sheet | name pattern | created by | tab colour | Excel Tables | names |
|---|---|---|---|---|---|
| Answer | `Q<n> Answer` | `_add_sheet` (pipeline.py:545) | navy #1F3864 (xl/styling.py:524; sheetorder `ANSWER_TAB = 0x64381F`, xl/sheetorder.py:25) | none | sheet-level hidden `smartass_pagesetup` (xl/styling.py:761); built-in Print_Area; manual HPageBreaks |
| Data | `Q<n> Data` | `_add_sheet(..., DATA_TAB)` (pipeline.py:546) | grey #A6A6A6 | `Data_Q<n>` (render/excel.py:39) | `smartass_pagesetup` (via `sheet_print`), Print_Area |
| Lines | `Q<n> Lines b<i>` (i = index of the lines block in spec.blocks, server.py:163) | `drill_xl.write_lines -> add_sheet(..., color=0xA6A6A6)` (xl/drill_xl.py:236-237) | grey #A6A6A6 | `Lines_Q<n>_b<i>` (`table_name`, xl/drill_xl.py:27-28; alias "d" -> `Lines_Q<n>` is v1 only) | `smartass_pagesetup`, Print_Area, Print_Titles `$4:$4` |
| Sources log (workbook-wide, all answers) | `Sources`, or `Sources SV` when "Sources" is a user sheet or a v1 log (xl/sheets.py:307-313) | `sources._sheet` (xl/sources.py:56-67), at the end; tracked in `sheets_added` when new | none (class "other" in sheetorder) | none (an AutoFilter on A1:N<end>) | `smartass_pagesetup` |
| Spec store (workbook-wide) | `_svmeta` (xl/sheets.py:157) | `_save_spec` (pipeline.py:235-237), after the last sheet | - (hidden, `Visible = 0` = xlSheetHidden, not VeryHidden) | none | - |

Workbook-level: hidden name `_tnStylesVer` = `="<sha1[:10] of tn.style_defs()>"` (xl/styling.py:138-149) and the 54 TN
named styles (19 fixed + 7 kinds x 5 emphases, xl/vendor/tn_rapportstil.py:34-73), installed by `ensure_styles` on the
first answer into a workbook without them (xl/styling.py:107-135: one `Styles.Merge` from `tn_template_<ver>.xltx`,
opened ReadOnly with alerts off; else ~600 property writes). A new analysis file starts from the template, so it has
them already.

A new analysis workbook's first sheet is an empty sheet renamed "Sources" (xl/target.py:199), which `log_kind` calls
"empty" and the Sources log then takes over (xl/sources.py:37-38, xl/sheets.py:311).

### 2.1 Q<n> Answer

Fixed zones in this order (render/layout.py:428-544; checked by `ZONE_ORDER`, render/check.py:20): `title` (row 1,
text in B1), `subtitle` (row 2), `key_points` (row 4 "Key points" section, then the point rows), `charts`
(`ceil(n_charts/2) * 18` empty rows, render/layout.py:454), `tables` (every non-chart block in spec order: section row
with the block title in B, header row, line rows, one blank row; render/layout.py:389-426), `checks` ("Checks", the
"Sheet = model" row whose D cell is `plan.status_cell`, then one OK/DIFF row per check line), `notes` (conventions,
chain note, bridge axis notes, Thresholds/Parameters tables, Findings, Caveats), `chart_data` (chart blocks' own data
tables), `source` (one note row "Source: Oracle EPM Cloud (<cubes>) · table Data_Q<n> (sheet Q<n> Data) · Sources
sheet", render/layout.py:544). Column A empty (width 2); B codes and every title/section/note/text
(`TEXT_COL = CODE_COL = 2`); C labels (`LABEL_COL = 3`); values from D (`FIRST_COL = 4`) (render/layout.py:24-27).
Chart helper columns sit right of every table and of the charts' right edge, under a grey "Chart helper" header
(render/layout.py:1137-1170).

Formulas on the Answer that reference a **table** (these are the ones the design leaves out of the file):

- every cell whose expression tree contains a `Pick` (model/expr.py:59-80) -> `SUMIFS(Data_Q<n>[Value],...)`, and
  every one that contains a `LinePick` (model/expr.py:294-318) -> `SUMIFS(Lines_Q<n>_b<i>[AMOUNT],...)` or
  `SUM(Lines_Q<n>_b<i>[AMOUNT])`. The layout writes them in pass 2 (render/layout.py:550-558): `"=" +
  e.to_excel(table, addr, const)`, a `Var` with a missing base guarded as `=IF(<base>=0,"",<a-b>)`
  (render/layout.py:554-556), a `Const` as a plain value (render/layout.py:552-553).
- Real strings, as the tests pin them (table `T` there stands for `Data_Q<n>`):
  - `=SUMIFS(Data_Q3[Value],Data_Q3[Account],"ac.C50S",Data_Q3[Period],"Sep 26")` (engine/tests/v2/test_v2_expr.py:19)
  - `=SUM(SUMIFS(Data_Q3[Value],Data_Q3[Account],{"ac.C50S","ac.C60"},Data_Q3[Period],"Sep 26"))` (one multi-value
    criterion: an array constant wrapped in SUM; two multi-value criteria use `,` then `;`, model/expr.py:71-80;
    more than two are split into a Sum of Picks, model/expr.py:267-283) (engine/tests/v2/test_v2_expr.py:25)
  - `=SUMIFS(Lines_Q3_b0[AMOUNT],Lines_Q3_b0[AP_CUSTOMER_NAME],"Accenture AS")` and `=SUM(Lines_Q3_b0[AMOUNT])`
    (engine/tests/v2/test_v2_lines_expr.py:30-31)
  - Criteria on a column that holds one value over the whole table are dropped (`const` = `constant_columns(frame)`,
    render/layout.py:176-183; model/expr.py:69-70), so a typical formula has no `[Cube]`, `[Query]` or `[Scenario]`
    criterion (engine/tests/v2/test_v2_short_formulas.py:17-23). The table name always comes from
    `table = f"Data_Q{q}"` (render/layout.py:328); `LinePick` derives `Lines_Q<n>_<alias>` from it
    (model/expr.py:311).
- **Not** table references, but sheet references: `CubeRef` -> `'Q3 Lines b0'!$D$3` (model/expr.py:321-332,
  engine/tests/v2/test_v2_lines_expr.py:39). It can stay in the built file only if the Lines sheet is copied in the
  same `Copy` call (references between sheets copied together follow the copies - **to verify** for the one-call copy).
- Everything else on the Answer references cells of the Answer only: `Ref` -> an address, `Sum` of Refs ->
  `SUM(D5:D9)` (model/expr.py:97-128), `Ratio` -> `IF(<d>=0,"",(<n>)/(<d>))` / `IF(ABS(<d>)<$D$40,"",...)`
  (model/expr.py:215-220), `Material` -> `IF(OR(ABS(F5)>=$D$41,AND(ABS(N(G5))>=$E$41,ABS(F5)>=$F$41)),1,0)` shape
  (model/expr.py:361-372), check rows `=IF(MAX(ABS(D20),ABS(E20))<=$D$45,"OK","DIFF")` shape (render/layout.py:560-563),
  Key point sentences `="•  <label> <period> "&SUBSTITUTE(FIXED(D12/$D$44,1),"-","−")&" MNOK; "&...`
  (render/keypoints.py:76-107), chart helpers `=MAX(MIN(T11,T12),0)` etc. (render/chartdata.py:19-45, 88-96).
  (Addresses above are illustrative shapes, not real output.) These could be written into the file - but their
  results are only right after the table formulas exist and Excel has calculated.

### 2.2 Q<n> Data

- Header block rows 1..h in A:D (render/datasheet.py:1-12, 92-125): A1 the title `"Q<n> Data — <question>"` (value
  only); row 2 Question; per query 8 rows (Query, POV, Periods, Route, Fetched, Rows, Control total, Sources); then
  "Tolerance (NOK)" = 1.0 and "Refresh" text. `tol_row = 3 + 8 * len(queries)` (render/datasheet.py:115), asserted.
  One query: rows 1-12, table header at row 14 (`Header.table_row = len(header) + 2`, render/datasheet.py:36-38).
- Formulas (all on the Data sheet, two of three reference the table):
  - Rows: `=COUNTIFS(Data_Q3[Query],"fm")` (+ `+COUNTIFS(Data_Q3[Query],"fm:*")` when the query has derived concept
    facts) (render/datasheet.py:104-105; engine/tests/v2/test_v5_datasheet.py:80)
  - Control total: B the static EPM sum (a number), C `=SUMIFS(Data_Q3[Value],Data_Q3[Query],"fm")`, D
    `=IF(ABS(B9-C9)<=$B$11,"OK","DIFF")` (render/datasheet.py:106-108; status cell `D<first+6>`, sources cell
    `B<first+7>`, render/datasheet.py:110).
  - The Sources value cell stays empty in the header and gets a **hyperlink** to `'Sources'!A<r>` with text
    `"Sources row <r>"` once the Sources row is known (xl/sources.py:159-161).
- Table `Data_Q<n>` at `A<table_row>`: columns `Cube, Query, Account, Entity, [Product,] Period, Scenario, Version,
  Balance, Value` (Product only when some fact has one, model/frame.py:9-11, 49-50); one row per fact; label columns
  are text (`NumberFormat "@"` on every column but the last, header included, render/excel.py:26-33).
- Frozen below the table header (`styling._freeze(ws, r0, 0)`, render/excel.py:222), gridlines off.

### 2.3 Q<n> Lines b<i> (details in §6)

A1 title "Ledger lines behind <cell values joined ' / '>" (bold), A2 the source text, C3 "Cube value", D3 the cube
value (`CUBE_CELL = "$D$3"`, model/lines.py:8), F1:K3 provenance + tie-out, row 4 the table header, rows 5.. the
lines, table `Lines_Q<n>_b<i>` on `A4:<last><4+n>`. Formulas: G3 `=SUM(Lines_Q3_b0[AMOUNT])` (table), I3
`=IF(NOT(ISNUMBER($D$3)),"no cube value",IF(ABS($G$3-$D$3)<=$K$3,"OK","DIFF"))` (cells only)
(render/datasheet.py:133-145).

### 2.4 Sources and _svmeta

Sources: §6.2. _svmeta: §8.3. Neither is per answer: both are shared by every answer in the workbook.


## 3. The plan object (render/layout.py)

### 3.1 `SheetPlan` (render/layout.py:107-130)

| field | type | meaning |
|---|---|---|
| `cells` | `dict[(row, col), object]` | every Answer cell: text, number, or a formula string starting with "=" (1-based row/col) |
| `rows` | `list[dict]` | one entry per styled row: `{"row": r, "kind": k, ...}`, k in title, subtitle, section, note, text, subhead, header, line, sum, grand, check; `"num": {col letter: number kind}` on line/sum/grand/check rows; `"num_cols": [letters]` on header rows (right-aligned numeric headers); `"status_col": "D"` on check rows. Rows not listed are "blank". |
| `addr` | `dict[str, str]` | model cell key `"<block>.<line>.<col>"` (`ck`, model/blocks/common.py:139) -> A1 address: the read-back compares exactly these |
| `last_col` / `last_row` | `str` / `int` | the drawn block; last_col includes the chart helper columns (render/layout.py:1169-1170) |
| `charts` | `list[ChartSpec]` | §3.2 |
| `fractions` | `set[str]` | keys of percentage cells (read-back: no near-zero rule) |
| `params` | `dict[str, str]` | threshold/parameter key -> absolute address (`$D$40`) |
| `zones` | `list[(name, first row)]` | sheet order of the zones |
| `status_cell` | `str` | the empty "Sheet = model" cell, written after the read-back |
| `widths` | `dict[str, float]` | column letter -> width in characters: A 2, B `code_width` (9..24), C `label_width` (46..80), D.. 16, gap 2, Material 9, helper 16 (render/layout.py:645-653) |
| `chart_lint` | `list[str]` | chart problems found while laying out |
| `helper_cols` | `list[str]` | chart helper column letters |
| `print_last_col` | `str` | last table column (text rows end here) |
| `print_area_col` | `str` | print area's last column (tables, or the charts' right edge) |
| `footer` | `str` | set by the pipeline after `layout()` (pipeline.py:479) |
| `literal` | `set[(row, col)]` | text cells always written as text even when they start with "=" |

Method `at(key) -> (row, col)` (render/layout.py:127-130).

Number kinds in `rows[*]["num"]` (`_num`, render/layout.py:133-142): `heltall` (amounts), `tall1` (stock metrics:
FTE), `avvik` (variance NOK), `avvikpst` (variance %), `prosent1` (share, key, threshold %), `markor` (Material
0/1). Helper cells get `heltall` or `tall1` (render/layout.py:1160).

Layout constants a writer needs (render/layout.py:24-39): `CHART_ROWS = 18`, `CHART_WIDTH = 480.0` pt,
`CHART_GAP = 8.0`, `CHART_HEIGHT = 260.0`, `ROW_PT = 15.0`, `DEFAULT_WIDTH = 8.43`, `NOTE_WIDTH = 78`, and
`width_pt(chars) = (chars * 7 + 5) * 0.75` (render/layout.py:203-205), which the chart anchors and the print area use.

### 3.2 `ChartSpec` (render/layout.py:64-104)

| field | default | meaning |
|---|---|---|
| `kind` | - | `waterfall` / `line` / `column` / `bar` / `stacked` / `stacked100` / `combo` / `variance` |
| `title` | - | final title (`chart_title`, render/layout.py:842-860) |
| `cat_range` | "" | category labels: column C rows (`'C12:C16'`, or `'C12:C14,C16'` with gaps); for stacked: the header row across the picked columns (render/layout.py:1062) |
| `series` | [] | `[(name, range, role)]`; name is literal text (a header), range like `'D12:D16'` or `'D12:D14,D16'` on the Answer; role `ac py bu fc var pct`, `fav unfav` (variance), `base up up_neg down down_neg total` (waterfall, render/chartdata.py:13-14), `part` (stacked). [] = not drawn. |
| `anchor` | "B1" | top-left cell: B or the first column 488 pt right of B, row `charts_row + (i//2)*18` (render/layout.py:1173-1186) |
| `width` / `height` | 480 / 260 pt | |
| `millions` | True | data labels in MNOK (`#,##0.0,,`) |
| `whole` | False | NOK under 1 m: `#,##0` |
| `reverse_values` | False | every plotted value <= 0: value axis reversed |
| `axis_min` | None | truncated bridge axis: MinimumScale (MaximumScale when negative) |
| `thin_steps` | [] | waterfall points (1-based) too thin for a white label |
| `sort` | "none" | bar: `desc` keeps the first row on top |
| `highlight` | [] | 1-based points in the accent colour, the rest grey |
| `source` | "" | the out key (`b<i>`) |
| `helpers` | {} | after layout: `{(row, col): formula or 0}` of the helper cells (already merged into `cells`) |
| `hues` | [] | per series: index into `charts.HUES` |
| `label_points` / `label_below` / `label_min` / `label_pos` | [] | line charts: which points get labels, labels below, the min point, per-point "above"/"below"/"left" |
| `unit` | "nok" | unit class (never mixed in one chart) |
| `auto_title` | False | engine-named chart |
| `_meta` | None | internal (rows, picked columns) |

Properties `anchor_col`, `anchor_row`, `height_pt` (render/layout.py:94-104). There is no separate chart-spec type in
`render/chartdata.py`: it only holds the helper-column formulas (`waterfall_helpers`, `variance_helpers`) and their
Python mirror (`waterfall_bars`, `stacked_spans`).

### 3.3 Key points and the header

- Key points are not a plan object: `layout()` calls the `key_points` hook twice (once with a placeholder `addr` to
  reserve rows, render/layout.py:441; once with the real one, render/layout.py:564-565); each point is a `text` row in
  column B holding either a formula (not literal) or a wrapped static line (`point_rows`, render/keypoints.py:328-341).
- The Data header is `datasheet.Header(list)` (render/datasheet.py:26-38): `[(label, value)]` in row order, value =
  text, number, formula or a tuple of them (B, C, D); `.sources: {query id: "B10"}`, `.status: {query id: "D9"}`,
  `.table_row`. Pure. The table grid is `data_grid(frame) = [frame.columns(), *frame.rows()]` (render/excel.py:22-23).

### 3.4 What a file writer needs that is NOT in the plan today

Answer sheet (all of this is computed or applied in `styling.apply_report` / `render/excel.py` / `render/charts.py`):

1. **The style of every cell**: row kind -> TN style per column (xl/styling.py:399-469), including column A and the
   helper columns; the status column style; a numeric value in a column without a `num` entry gets "TN Heltall"
   **judged from the value Excel computed** (xl/styling.py:455-456).
2. **Concrete number-format codes** for each kind (xl/styling.py:82-91, xl/vendor/tn_rapportstil.py:10-17) (§4.2).
3. **Row heights** (xl/styling.py:408-459) and **merges** (text rows merged B:`print_last_col` when the cells to the
   right are empty, xl/styling.py:425-426; a title that wraps, xl/styling.py:409-415). The text row heights use the
   **computed** Key point sentence (Value2 read at xl/styling.py:374).
4. Code column grey + left (xl/styling.py:327-330, 482-484), helper columns grey font (xl/styling.py:485-488).
5. The check conditional formats (xl/styling.py:333-344).
6. Tab colour, gridlines off, no freeze on the Answer (`answer_roles` sets no `freeze_rows`, so
   `_freeze(ws, 0, 0)`, xl/styling.py:525).
7. Page setup: header text "Intern", landscape, one page wide, margins, `PrintGridlines False`, the manual page breaks
   (`keep_spans(plan)` render/excel.py:292-310 + `keep_breaks(heights, ...)` xl/styling.py:625-644, which needs the
   row heights from item 3).
8. The text-vs-formula rule: a string is text unless it starts with "=" and is not in `literal`; text is written with a
   leading apostrophe (render/excel.py:261-262).
9. The **per-block rectangles** the design wants for one `Formula2` array and one read-back per block: `block_rows`
   (header row, line rows, column letters per out key) is a local of `layout()` (render/layout.py:350, 425) and is
   not kept on the plan. Only per-cell `addr` and `rows` survive.
10. Which cells reference a table (to leave out of the file): not marked on the plan, which holds formula strings only.
    The expression trees are still available on `pre.result.outs[*].lines[*].cells` (model/blocks/common.py:25-36,
    67-85), and `plan.addr` maps each tree's key (`ck`) to its address, so a writer can walk the trees for
    `Pick`/`LinePick` (§2.1); the layout's own extra formulas (check rows, Key points, helpers) never reference a table.
    Checking the strings for `Data_Q<n>[` / `Lines_Q<n>_` is the cruder alternative.
11. The status cell's text (after the read-back).
12. Chart rendering rules: everything in §5 (colours, labels, axes, legend) lives in `render/charts.py`; the plan
    only carries kind, ranges, roles, hues and label hints.

Other sheets: there is **no plan** for the Data sheet beyond `Header` + `data_grid` (its look is in
`render/excel.py:_write_header/_style`), for the Lines sheets (`xl/drill_xl.write_lines` + `pipeline._write_lines/
_lines_header`), or for sheet names, tab colours and sheet order (`pipeline`, `xl/sheets.add_sheet`,
`xl/sheetorder`).


## 4. The look

### 4.1 Where TN style comes from

- Tokens and named styles: `xl/vendor/tn_rapportstil.py` (pinned copy of the skill): `FONT = "Arial"`; NAVY #1F3864
  (31,56,100), INK #262626, GREY #767676, RULE_DARK #404040, INPUT_BLUE #0000FF, INPUT_FILL #FFF7DD, KEY_FILL #EDF2F9,
  BREACH_FONT #C00000, BREACH_FILL #FDE9E9 (xl/vendor/tn_rapportstil.py:5-8); `styling.GREEN = (0,112,60)` #00703C
  (xl/styling.py:80).
- The answer is styled with **named cell styles** (`Range.Style = "TN ..."`, xl/styling.py:471-474) on top of
  `ClearFormats` + `Style = "TN Tom"` for the whole block (xl/styling.py:381-386), plus **direct** overrides: some
  number formats (xl/styling.py:475-481), code-column colour/alignment, helper-column font colour, conditional
  formats, row heights, column widths, merges. The Data, Lines and Sources sheets use **direct formats only** (fonts,
  number formats) plus the built-in table style `TableStyleLight1`.
- `ensure_styles(wb)` (xl/styling.py:107) installs the named styles once per workbook version (§2). **To verify**:
  XlsxWriter writes cell formats, not named cell styles; with direct formats the copied cells will show style
  "Normal" + formats. The design allows that ("direct formats or named styles ... as long as `xl_tidy` reports no
  BLOCK finding"), and BLOCK findings come only from the formula checker: `xl_tidy` = optional `apply_report` when
  rows + last_col are given, then `fmt.check` -> `fx.check` over formulas (workbench.py:76-84,
  xl/workbench/fmt.py:223-232; severities in xl/vendor/formelsjekk.py:17-25, 213-402). No check reads a cell's `.Style`
  name; it is read only by `xl_snapshot`/`xl_revert` (to restore named styles, xl/workbench/history.py:96-115) and by a
  workbook reader that lists styles (xl/workbench/books.py:240-245).

### 4.2 The named styles the answer uses, as concrete formats

All Arial; "v" = vertical alignment; borders other than listed: none; no fill unless listed
(xl/vendor/tn_rapportstil.py:37-73).

| style | font | align | border / fill | number format |
|---|---|---|---|---|
| TN Tittel | 16 bold #1F3864 | left, v bottom | bottom medium #1F3864 | - |
| TN Undertittel | 9 italic #767676 | left, v center | - | - |
| TN Seksjon | 11 bold #1F3864 | left, v bottom | bottom thin #1F3864 | - |
| TN Underseksjon | 9 bold #1F3864 | left, v bottom | - | - |
| TN Kolonne | 9 bold #1F3864 | left, v bottom, wrap | bottom thin #1F3864 | - |
| TN Kolonne tall | 9 bold #1F3864 | right, v bottom, wrap | bottom thin #1F3864 | - |
| TN Note | 9 italic #767676 | left, v center, indent 1 | - | - |
| TN Tekst | 10 #262626 | left, v top, wrap | - | - |
| TN Tekst tallkolonne | 10 #262626 | right, v center | - | - |
| TN Tom | 10 #262626 | general, v center | - | - |
| TN Etikett | 10 #262626 | left, v center | - | - |
| TN Heltall | 10 #000000 | right, v center | - | `#,##0_);(#,##0);"–"_)` |
| TN Tall 1d | 10 #000000 | right, v center | - | `0.0_);(0.0);"–"_)` (overridden on answers, below) |
| TN Prosent 1d | 10 #000000 | right, v center | - | `0.0 %_);(0.0 %);"–"_)` |
| "<kind> · Sum" | + bold | | + top thin #404040 | |
| "<kind> · Hovedsum" | + bold | | + top thin #404040, bottom double #404040 | |

(Unused by answers but defined: Beløp, Prosent, Faktor kinds; Input/Nøkkel emphases; TN ID, TN Lenke, TN Kolonne
faktor, TN Stortall *, TN Produkt *.)

Direct number-format overrides on answers (set after the styles, xl/styling.py:82-91, 447-450, 475-481):

| plan kind | base style | format applied on top (en-US code) |
|---|---|---|
| `heltall` | TN Heltall (+emph) | - |
| `tall1` | TN Tall 1d (+emph) | `#,##0.0_);(#,##0.0);"–"_)` (TALL1_THOUSANDS) |
| `prosent1` | TN Prosent 1d (+emph) | - |
| `avvik` | TN Heltall (+emph) | `[Color10]"▲ "#,##0;[Red]"▼ ("#,##0")";"–"` |
| `avvikpst` | TN Prosent 1d (+emph) | `[Color10]"▲ "0.0 %;[Red]"▼ ("0.0 %")";"–"` |
| `markor` | TN Heltall (+emph) | `"●";"●";""` |

(`[Color10]` is palette index 10, green #008000 in Excel's default palette - from memory, **to verify**.)

### 4.3 Row kind -> styles, heights, merges on the Answer (`apply_report`, xl/styling.py:347-536)

Block = `A1:<last_col><last_row>` (own block, xl/styling.py:350, 368-372); `first_label` = column B (the first of
`label_cols = [B, C]`, render/excel.py:281); `text_last` = `print_last_col` (xl/styling.py:362).

| kind | styles per column | row height (pt) |
|---|---|---|
| title | TN Tittel on A..last_col (border across every column, helper columns included) | 28; when B's text wraps over B..text_last (`wrap_lines(text, width, 16, bold)` > 1): `n*21+4`, merged B:text_last in TN Tittel + wrap |
| subtitle | TN Undertittel A..last | 16 |
| section | TN Seksjon A..last | 22 |
| note | TN Note A..last | 15 (notes are pre-wrapped by the layout, render/layout.py:376-381) |
| blank (any unlisted row) | TN Tom | 15 |
| subhead | A: TN Tom; B..: TN Underseksjon | 15 |
| text (Key points) | A: TN Tom; B..: TN Tekst | `max(15, n*13+2)`, n = `wrap_lines(<computed text>, width B..text_last)` (4 if a word does not fit); merged B:text_last when C..text_last are empty |
| header | TN Kolonne tall on `num_cols`, TN Kolonne elsewhere (A included) | `max(26, lines*12+3)`, lines = max `wrap_lines(header, col width, 9, bold)` |
| line / sum / grand / check | A: TN Tom; `num` columns: `TN <kind>[ · Sum| · Hovedsum]`; status column: TN Tekst tallkolonne (TN Etikett · <emph> on sum rows); B, C: TN Etikett[ · emph]; other numeric value: TN Heltall; else TN Etikett | 15 |

Then: code cells (column B on table rows, when B is not a number column): font #767676, left (xl/styling.py:462,
482-484); helper columns `<first>1:<last><last_row>` font #767676 (xl/styling.py:485-488); column widths from
`answer_roles` (render/excel.py:278-279: A 2, B 9, C `label_width`, D..last 16, then `plan.widths` wins) set per run
of equal widths (xl/styling.py:490-502); row heights per run (xl/styling.py:503-512); merges `_merge_text`
(xl/styling.py:316-324: wrap, merge, style again, top alignment); check status cells: two conditional formats per run
(xl/styling.py:333-344): cell value = "OK" -> font #00703C bold; cell value <> "OK" -> font #C00000 bold, fill #FDE9E9;
tab colour navy (xl/styling.py:524); `_freeze(ws, 0, 0)` = gridlines off, scroll to A1, no split (xl/styling.py:238-276;
only when the workbook window is visible, and it activates the sheet and back, xl/styling.py:195-235).

Page setup of the Answer (`print_setup`, xl/styling.py:709-763, called at xl/styling.py:526-532): print area
`$A$1:$<print_area_col>$<last_row>`, no title rows, landscape (2), `Zoom False`, `FitToPagesWide 1`,
`FitToPagesTall False`, `PrintGridlines False`, `CenterHorizontally True`, CenterHeader `&"Arial"&8Intern`,
Left/RightFooter "", CenterFooter = `plan.footer`, Left/RightMargin = 1.5 cm (`CentimetersToPoints(1.5)`), top/bottom
left at Excel's default; manual `HPageBreaks.Add(A<r>)` at `keep_breaks(...)` rows (after `ResetAllPageBreaks`);
batched with `PrintCommunication = False`; a sheet-level hidden name `smartass_pagesetup` =
`="a1|$A$1:$X$99|Intern|<footer>[|breaks r1,r2]"` skips an identical re-setup (xl/styling.py:720-725, 759-763). Paper
size is not set (printer default). `keep_breaks` simulates pages of A4 landscape with those margins
(`PRINT_WIDTH_PT, PRINT_HEIGHT_PT`, xl/styling.py:622) from the row heights above.

### 4.4 Data, Lines and Sources look (direct formats)

- Data header (render/excel.py:161-193): `A1:D<h>` ClearFormats, then font Arial 9 #262626; A1 bold 14 #1F3864;
  A2:A<h> #767676; the Control total and Tolerance rows' B:C `#,##0;(#,##0);"–"` (`VALUE_FORMAT`,
  render/excel.py:156).
- Data table (render/excel.py:196-225): `TableStyle = "TableStyleLight1"`, `ShowTableStyleRowStripes = False`, font
  Arial 9 on the table, header row bold, Value body `VALUE_FORMAT` + font #0000C0 (`VALUE_BLUE = 0xC00000` BGR) +
  right; **each column `AutoFit` over the table rows**, then `min(width, 40)`, columns A-D at least 16
  (render/excel.py:212-220); frozen below the table header row (+ gridlines off); `sheet_print` (render/excel.py:240-252)
  -> `print_setup(ws, max(D, last table col), last table row, SHEET_FOOTER)`, `SHEET_FOOTER = "Source: Oracle EPM
  Cloud · &A · page &P of &N"` (render/excel.py:229).
- Lines and Sources: §6.

### 4.5 Locale

Through late-bound COM Excel reads number formats in the user's locale, so every numeric code is translated first
(`styling.local_format` / `tn.to_local` / `NumberFormatLocal`, xl/styling.py:782-790; enforced by the source-scanning
test engine/tests/v2/test_v5_locale_formats.py:28-48), and chart data-label codes get U+00A0 as the thousands/scaling
character under nb-NO (`chart_format`, xl/styling.py:793-807). An .xlsx stores en-US codes, so a file writer uses the
codes above as they are; whether the chart label code `#,##0.0,,` written into the file shows "3 811,4" in nb-NO Excel
is **to verify** live (the NBSP rule was found for the COM path only). The Lines date format is built from the locale's
day/month/year letters (xl/drill_xl.py:108-115) = `dd.mm.yyyy` in en-US code.

### 4.6 What depends on reading Excel back

| what | where | how a file writer could replace it |
|---|---|---|
| Key point row heights + merges + text-fit lint use the **calculated** sentence (Value2 of the block) | xl/styling.py:374-378, 419-426, 573-595 | build the sentence in Python from `res.cells`, or a fixed height per point row |
| numeric fallback style for a value cell without `num` | xl/styling.py:455-456 | treat any formula cell as numeric, or require `num` everywhere |
| Data table column widths (`AutoFit`) | render/excel.py:212-220 | `tn.text_units(text, 9)` over the grid (xl/vendor/tn_rapportstil.py:108-111) |
| Sources widths (`AutoFit`) | xl/sources.py:179 | stays COM (the log stays in the workbook) |
| `_lint_fast`: UsedRange, block `Font.Name`, vertical borders | xl/styling.py:547-571 | from the plan / the writer's own formats |
| thin waterfall label lifted 10 pt by reading `DataLabel.Top` | render/charts.py:215-220 | not expressible in a file: see §5 |
| `Application.International` for every number format | xl/styling.py:352-359, 782-807 | not needed in a file (en-US codes) |
| `_freeze` only when the window is visible; activates sheets | xl/styling.py:238-276 | `freeze_panes` in the file (no window needed) |
| check status, mismatches, Data OK/DIFF, lines tie-out | §7 | stays: these are the read-back |


## 5. Charts (render/charts.py)

### 5.1 How a chart is drawn today (`_draw`, render/charts.py:272-317)

1. `_place` (render/charts.py:90-95): `ws.Shapes.AddChart2(-1, XL_TYPE[kind], cell.Left, cell.Top, 480, 260)` at
   `spec.anchor`; shape name `"smartass chart"` (`owned.CHART_NAME`, xl/owned.py:9). `-1` = Excel's default chart
   style for the type (the Office 2013+ look).
2. `_drop_auto_series` (render/charts.py:98-103): delete whatever AddChart2 guessed.
3. `_add_series` (render/charts.py:106-117): per `(name, rng, role)` `NewSeries`, `Values = _series_ref(rng)`,
   `XValues = _series_ref(cat_range)`, `Name = name` (literal text). A range with gaps is passed as the union text
   `=('Q3 Answer'!$D$5:$D$7,'Q3 Answer'!$D$9)` (`union_ref`, render/charts.py:120-135). combo + role `pct`:
   `ChartType = xlLine (4)` and `AxisGroup = 2` before any format.
4. Formats (only after every series and type is set): chart-area font Arial (`ChartArea.Format.TextFrame2`,
   render/charts.py:278); title (`_title`, render/charts.py:261-269: text, Arial, size TN Seksjon 11 + 1 = **12**, bold,
   #1F3864); legend (render/charts.py:282-286); axes (`_axes`, render/charts.py:231-258); `GapWidth = 50` on chart
   group 1 for every non-line chart, `Overlap = 100` for variance and waterfall (render/charts.py:288-292); per series
   `_fill` + `_labels` (render/charts.py:293-297); highlight (render/charts.py:298-311).
5. Any exception: the half-drawn shape is deleted and the error re-raised; `write_charts` turns it into one lint
   line and draws the others (render/charts.py:312-317, 334-348). A spec without series is skipped (its zone stays).

### 5.2 Kinds

| kind | `XL_TYPE` (render/charts.py:26-27) | series / categories | notes |
|---|---|---|---|
| column | 51 xlColumnClustered | value columns of the source block's last column group; categories column C of the line rows | |
| bar | 57 xlBarClustered | same | category axis `ReversePlotOrder = (sort != "asc")` (first row on top) |
| line | 4 xlLine | per metric column (trend) | per-point labels (5.4) |
| stacked | 52 xlColumnStacked | one series per **row** (`_across(letters, row)`), categories = the header row across the picked columns (render/layout.py:1051-1063) | parts `PARTS` colours, legend always |
| stacked100 | 53 xlColumnStacked100 | as stacked | |
| combo | 51 xlColumnClustered + `pct` series as xlLine on the secondary axis | | secondary axis: no gridlines/ticks/labels/line (render/charts.py:243-248) |
| variance | 57 xlBarClustered | helper columns "Better" = `=MAX(var,0)`, "Worse" = `=MIN(var,0)` (render/chartdata.py:88-96), roles fav/unfav | `Overlap 100`, never a reversed value axis, no legend |
| waterfall | 52 xlColumnStacked | 8 helper columns `Run, Base, Base <0, Up, Up <0, Down, Down <0, Total`; series = all but `Run` (render/chartdata.py:13-45; render/layout.py:1164-1165) | bases invisible; never Excel's native waterfall (render/charts.py:13-14) |

Series and category ranges always point at the Answer sheet itself (the table rows or the helper columns), never at
the Data sheet.

### 5.3 Colours (RGB)

- Per-metric hues `HUES` (render/charts.py:51): navy #1F3864, teal #008080, amber #D68910, purple #7030A0, grey-blue
  #546E96; `spec.hues[k]` picks one per series.
- Bars (`_fill`, render/charts.py:138-173): `ac` solid hue; `py` solid hue tinted 55 % towards white (`tint`,
  render/charts.py:58-60: navy -> #9AA5B9, teal #8CC6C6, amber #EDCA93, purple #BFA2D4, grey-blue #B2BED0); `bu` white
  fill, 1 pt border in the hue; `fc` pattern `msoPatternWideUpwardDiagonal (26)` hue on white + border in the hue;
  `var/fav/up/up_neg` #00703C; `unfav/down/down_neg` #C00000; `total` #1F3864; `part` k-th of `PARTS` = #1F3864,
  #546E96, #8CA0BE, #BEC8D7, #A6A6A6 (render/charts.py:49); `base` no fill and no line. Bars have no border except
  bu/fc.
- Lines: weight 2.25 pt; colour = hue, `py` the tint, `pct` #404040; `bu` dash (msoLineDash 4), `fc` round dot (3).
- Highlight (column/bar only): the ac/part series grey #A6A6A6, highlighted points accent #1C16C5
  (render/charts.py:45, 223-228), comparison series white fill with a 1 pt grey border (render/charts.py:302-310).

### 5.4 Data labels (`_labels`, render/charts.py:179-220)

- Every visible series: `HasDataLabels = True`, `NumberFormatLinked = False`, `NumberFormat = label_format(role,
  millions, whole)` (render/charts.py:77-87): MNOK `#,##0.0,,`, whole NOK `#,##0`, else `#,##0.0`; `pct`
  `0.0 %;"−"0.0 %;;`; positives/negatives with the typographic minus `"−"` and zeros hidden (`unit;"−"unit;;`);
  waterfall `up`: `unit;unit;;`, `down`: `"−"unit;"−"unit;;`.
- line: series position `Below` when every value <= 0, else `Above`; then per point: not in `label_points` ->
  `HasDataLabel = False`; the last labelled point `Right`; others from `label_pos` (above / below / left).
- stacked parts: `Center`, font white on the two darkest parts, ink #262626 on the others.
- combo pct: `Above`.
- waterfall total: `InsideEnd`, step: `Center`; font white bold; on `thin_steps` points font ink #262626 and the label
  moved up 10 pt by reading and writing `DataLabel.Top` (render/charts.py:215-220).
- No label font size set: Excel's default.

### 5.5 Axes, legend, title

- Value axis: no major/minor gridlines, no tick marks, `TickLabelPosition = xlNone`, axis line invisible
  (render/charts.py:237-242); `reverse_values` -> `ReversePlotOrder = True`; `axis_min` -> `MinimumScale` (or
  `MaximumScale` when negative).
- Category axis: no tick marks; `TickLabelPosition = Low`, or `High` when the value axis is reversed
  (render/charts.py:253-256); bar/variance `ReversePlotOrder` (render/charts.py:257-258). Category axis line and label
  font size: left at AddChart2's defaults.
- Legend (render/charts.py:280-286): stacked/stacked100 always; waterfall/variance never; otherwise when two or more
  visible series, or one series whose role is not ac/part/fav/unfav/var (e.g. a BU-only line). Position top.
- Title: always, as in 5.1.

### 5.6 For XlsxWriter (every item **to verify** against the XlsxWriter docs/version used)

- Types: `column`, `bar`, `line`, `column + subtype stacked / percent_stacked` exist; combo = a column chart
  `.combine()` with a line chart whose series has `y2_axis: True`.
- Needed and believed supported: per-series `fill`/`border`/`pattern` (`wide_upward_diagonal`), `line` with
  `dash_type` (`dash`, `round_dot`) and width; per-point `points` fills (highlight); `gap` 50 and `overlap` 100;
  `data_labels` with `num_format`, `position`, `font`; `custom` data labels with `delete` (line charts) and per-point
  font colour (thin waterfall steps); axis `major_gridlines: visible False`, `major_tick_mark: none`,
  `label_position: none/low/high`, `line: none`, `reverse`, `min`/`max`; legend `position: top` / `none`;
  `set_title` with a font; `insert_chart(cell, chart, {x_scale/y_scale or set_size(width px, height px)})` (480 x 260 pt
  = 640 x 347 px at 96 dpi).
- Likely gaps: (a) the per-point label **position** of line charts (last point Right, others above/below/left) -
  XlsxWriter's custom labels may not take a position (**to verify**); (b) the 10 pt lift of thin waterfall labels
  (`DataLabel.Top`) has no file equivalent short of manual layout - accept ink colour without the lift, or another
  rule; (c) the chart/shape name "smartass chart" (only needed by `owned.drop_charts`, which the design deletes);
  (d) series names are literal strings today - fine in XlsxWriter (`name` as text); (e) non-contiguous ranges
  `=('Sheet'!$D$5:$D$7,'Sheet'!$D$9)` for categories/values (**to verify**; the layout could avoid them by writing
  contiguous helper copies).
- AddChart2 defaults that today are implicit and must become explicit in a file: the chart style (Office 2013+ look:
  white chart area without a visible border, grey category axis line, label/axis font size 9 and colour grey #595959),
  the chart-wide Arial font (XlsxWriter has no chart-wide font: set it on title, axes `num_font`, legend and data
  labels), marker-less lines, plot-area size/position. "Same type, series, colours, labels and title" is the design's
  bar; pixel parity is not required.


## 6. Sources and Lines

### 6.1 Lines sheets: `pipeline._write_lines` (pipeline.py:367-397) + `drill_xl.write_lines` (xl/drill_xl.py:207-267)

Order inside `_draw`: after the Data sheet, before the Answer (pipeline.py:469-472), one per `LineSet` (alias `b<i>`).

`drill_xl.write_lines(t, q, rows, columns, title="Ledger lines behind <cell>", source=..., alias)`:

1. **The 50 000-row rule**: `cols = columns if len(rows) <= config.DRILL_XL_FULL else [c for c in columns if c in
   KEY_LABELS]` (xl/drill_xl.py:216); `DRILL_XL_FULL = int(env SV_DRILL_XL_FULL or 50000)` (core/config.py:257);
   `KEY_LABELS` = 22 key columns (epm/drill_lines.py:17-20), which include every `group_by` column the answer sums on
   (AP_CUSTOMER_NAME, INITIATIVE, ACCOUNT, model/blocks/lines.py:11). Full detail stays in drill.db.
2. `tname = table_name(q, alias)` = `Lines_Q<q>_<alias>`; formats from the locale: amount `#,##0_);(#,##0);"–"_)`
   (`AMT_FMT`, xl/drill_xl.py:75), date `dd.mm.yyyy` (xl/drill_xl.py:108-115).
3. `find_table(t.wb, tname)` - **iterates every sheet and every ListObject of the workbook** (xl/drill_xl.py:127-133).
   - Found with the same header: reused (body cleared, later `Resize`, stale cells cleared).
   - Found with other columns: `lo.Delete()` + clear, rebuilt under the same name.
   - Not found: `name = unique_sheet_name(t, "Q<q> Lines <alias>")`, `add_sheet(t, name, after_q=q, color=0xA6A6A6)`
     (tracked for rollback), header row `hdr = 4`.
4. Inside `fast_excel(t.app)` (xl/drill_xl.py:222): A1:A2 = title, `_source_text`; A1 bold, A2 font black
   (xl/drill_xl.py:188-190); row 4 = column names; the lines in chunks of `CHUNK = 5000` rows, one `Range.Value` per chunk
   with a 180 s COM limit (xl/drill_xl.py:247-251). Cell values (`_cell`, xl/drill_xl.py:93-105): text with a leading
   apostrophe; `TEXT_LABELS` columns always text; datetimes as Excel serial numbers (tz dropped). New table:
   `ListObjects.Add(1, A4:<last><end>, None, 1)`, `Name = tname`, `TableStyle = "TableStyleLight1"` (xl/drill_xl.py:193-195,
   252-254) - row stripes left **on** here (unlike Data). Amount columns (`AMOUNT_LABELS`) and date columns
   (`DATE_LABELS`) get their formats per run of columns; widths `max(WIDTHS.get(col, 14), header_width(col))`
   (xl/drill_xl.py:80-89, 260-262); `_freeze_below(t, ws, 4)`: activates **the workbook and the sheet**, selects A5,
   `ActiveWindow.FreezePanes = True` (xl/drill_xl.py:198-204) - gridlines stay on.
5. `sheet_print(ws, last_col, end, header_row=4)`: 100 % zoom, as many pages wide as needed, row 4 repeated
   (render/excel.py:246-249).

Then `pipeline._write_lines` adds (pipeline.py:379-397): refuses when the sheet came out under another name (a
`GuardError`: CubeRef needs the exact name); clears C2:D2; A2 = `LINES_SOURCE` (pipeline.py:364); C3 "Cube value", D3
= `ls.cube_value` in the amount format (`NumberFormatLocal`); `_lines_header` (pipeline.py:403-421): F1:K3 as one
`Range.Formula` rectangle (text with apostrophes) after `ClearFormats`, Arial 9 #262626, F1:F3/H3/J3 #767676, G3
`VALUE_FORMAT`; then `sheet_print` again up to `max(table width, K)` (pipeline.py:392-396).

Formulas on the Answer that sum over the Lines table come only from the `lines` block (model/blocks/lines.py:15-42):
each group line `LinePick(alias, column, spellings)` -> `SUMIFS(Lines_Q<n>_<alias>[AMOUNT],Lines_Q<n>_<alias>[<column>],
"<name>")` (several spellings: `SUM(SUMIFS(..., {"a","b"}))`), "Lines without a ..." = the same with `""`, "Ledger lines
total" = `SUM(Lines_Q<n>_<alias>[AMOUNT])`, "Cube value" = `CubeRef` `'Q<n> Lines <alias>'!$D$3`, "Other" and "Gap" are
cell arithmetic, shares are `Ratio`s (model/expr.py:294-332).

Also in drill_xl but **not** used by the v2 answer: `drill_formula` (SUMIFS or an exact-match `SUM(FILTER(...))` form,
xl/drill_xl.py:44-71), `clear_lines`, `mark_incomplete`, `table_columns`, `lines_tables` (v1 fin_drill helpers; only
tests call `drill_formula`).

### 6.2 The Sources log: `pipeline._write_sources` (pipeline.py:424-446) + `sources.write_sources` (xl/sources.py:113-190)

Written after the calculate and the read-back (stage "sources", pipeline.py:490-491), because it records calculated
values:

- before: per query `ws_data.Range(header.status[qid]).Value` (OK/DIFF), per line set
  `'Q<q> Lines <alias>'!I3` (pipeline.py:433-439).
- sheet: `_sources_name` -> "Sources" unless that sheet's `log_kind` is "v1" or "user" (then "Sources SV";
  xl/sheets.py:307-313, xl/sources.py:31-39); created at the end when missing.
- inside `fast_excel`: Known limits in column P (`LIMITS_COL = 16`): read P2.. (up to 500), merge with this answer's
  warn/block caveats, rewrite P1:P<n> (xl/sources.py:129-135); A1:N1 = `LOG_COLUMNS` (xl/sources.py:20-21); find the last
  used row (`End(xlUp)`) and the rows already logged (xl/sources.py:90-98); **append** one row per query
  (`fetched, Q<q>, query id, question, cube, scenario/version, POV, periods, route, rows, control total, OK/DIFF,
  data sheet, "row <r>" of the latest earlier row for the same Q and id`) and one per line set (cube FinLite, route
  "drill", rows, lines total, tie-out, the Lines sheet) (xl/sources.py:139-162); per row a hyperlink in M (DATA_COL 13)
  to `'<sheet>'!A1`, and per query a hyperlink **back** in the Data header's Sources cell to `'<Sources>'!A<r>`
  (xl/sources.py:157-161); font Arial 9 on A1:N<end> and the limits, A1:N1 and P1 bold navy, K (control total)
  `VALUE_FORMAT` (xl/sources.py:165-174); AutoFilter off and on again over A1:N<end>; `Columns.AutoFit`; D capped at 60,
  P = 80 wide (xl/sources.py:175-184).
- after: `_freeze(ws, 1, 0)` (activates the sheet), `sheet_print(ws, "N", end)` (xl/sources.py:185-189).
- returns `{query id or alias: row}`; a missing id -> warning "the Sources sheet has no row for ..."
  (pipeline.py:443-446); any exception -> warning "could not log the pull on the Sources sheet (...)"; either makes
  `_check` report "the Sources rows of this answer are missing" (pipeline.py:449, 495; render/check.py:184-185).
- Cost: roughly 30-45 COM calls per answer with one query (my count from the code, not measured), several of them
  page-setup and window calls. The design's file build does not cover this sheet; it is workbook state (append-only,
  row numbers) and has to stay a COM step or be redesigned.


## 7. Read-back and checks

### 7.1 Calculate and read back

- `_calculate(t)` (pipeline.py:226-230): `t.app.Calculate()` - every open workbook in that Excel - then
  `wait_calculated` (xl/sheets.py:316-326: `CalculationState == 0`, 0.3 s poll, 60 s, then returns silently).
  (Calculation was never switched to manual for the Answer, so Excel has also been recalculating during the writes.)
- `read_block(ws, plan)` (render/excel.py:332-335): one `Range("A1:<last_col><last_row>").Value2` - the whole block from
  A1, helper columns included.
- `compare(plan.addr, pre.result.cells, sheet, plan.fractions)` (render/readback.py:23-48): for every model key in
  `plan.addr`: sheet value at that address vs `result.cells.get(key)`; Excel error ints (`EXCEL_ERRORS`,
  render/readback.py:12) always mismatch; None/"" on either side = 0; a non-number on the sheet mismatches; amounts both
  under `NEAR_ZERO = 1e-3` are equal (not for fractions); tolerance `TOL_ABS = 0.5` NOK when either |value| >= 1, else
  1e-6, and at least `TOL_REL = 1e-9 * |model|`. Returns `"C4: sheet x, model y"` lines.
- `write_status` (render/excel.py:323-329): `sheet_status(mismatches)` = "OK" or "DIFF in <n> cells"
  (render/excel.py:318-320) into `plan.status_cell` (D of the "Sheet = model" row). Written **after** the read-back,
  so the status cell itself is never compared.

### 7.2 The layout check (`render/check.py`, called as `_check` at pipeline.py:452-455, after Sources/spec/arrange)

`check(plan, ws, data_ws, sources_ok, status_cells)` (render/check.py:162-186), never raises; findings are lines with a
fix:

1. zones in `ZONE_ORDER`, no "plan" zone (render/check.py:71-81) - plan only;
2. at least one chart in the plan, and `ws.ChartObjects().Count > 0` (render/check.py:84-90) - **Excel**;
3. every numeric cell of a line/sum/grand row has a `num` kind; a General number >= 1e11 named as E+
   (render/check.py:93-115) - plan + the **Value2 grid** (`_grid`, one more read of the block, render/check.py:61-68);
4. no number wider than its column: `shown_text` (the TN format rendered in Python) measured with `tn.text_units`
   against `plan.widths` (render/check.py:41-58, 118-133) - **calculated values**;
5. no Key point cell shows an Excel error (render/check.py:136-148) - **calculated values**;
6. every Data header control-total status reads "OK" (`data_ws.Range(cell).Value`, render/check.py:151-159) - **Excel**;
7. the Sources rows were written (`sources_ok` from the `_write_sources` warning, pipeline.py:495).

### 7.3 Where the reply's fields come from (`reply`, pipeline.py:171-197)

`reply()` has no `lint` key: its `lint` argument is appended to `warnings` (pipeline.py:187).

| field | source | from Excel? |
|---|---|---|
| `q` | `next_question` over the sheet names (or the follow-up's q) | yes (sheet names) |
| `workbook` | `str(t.path)` | yes (FullName, OneDrive-mapped) |
| `sheets` | `["Q<q> Answer", "Q<q> Data", *["Q<q> Lines <alias>"]]` built from names, not read | no |
| `key_values`, `material`, `findings` | the model (`pre.result`) | no |
| `checks.sheet_equals_model` | `sheet_status(mismatches)` | **yes** (read-back) |
| `mismatches` (first 20) | `compare` | **yes** |
| `layout` | `render.check.check` | partly (rules 2-6) |
| `charts` | titles of the charts drawn without error (render/charts.py:342-344) | yes (a COM failure drops a title) |
| `warnings` | `pre.warnings` (fetch, too many points) + style lint (`_lint_fast`, xl/styling.py:539-596: font name, vertical borders, content outside the width, text that does not fit - the fit uses the calculated texts) + chart lint and chart failures + the Sources note + the arrange note; then `_with_notes`: AutoSave note, drill notes (server.py:211-219), `remember` failure, redraw "no stored range" | mixed |
| `caveats`, `routes`, `closed`, `fetch_seconds` | `Prepared` | no |
| `saved_user_workbook` | always False | no |
| `critic_brief` (standard/deep) | `critic.brief(spec, res, caveats, checks, stock)` (pipeline.py:194-196) | uses `checks` |
| `followup` | True from `redraw_job` | no |
| `text_fixes` | drill notes (server.py:213-218) | no |


## 8. Follow-ups

### 8.1 The flow (server.py + followup.py)

- `fin_analyze(patch=..., q=...)` (server.py:266-273): `target = _default_target(target)` (server.py:222-241: "session"
  means the workbook of the latest stored answer in answers.db unless that is the session file itself);
  `stored, active, path = await _user(lambda h: read_stored(h, target))` - **one short user-Excel job of its own**
  (pipeline.py:348-361: refuses "new", and "session" without a session workbook; returns `sheets._stored_specs(t)`,
  `ActiveSheet.Name`, `str(t.path)`); `q2, raw, owned = find_question(stored, q, active)` (followup.py:31-48: explicit
  q must be a v2 answer; else the active "Q<n> Answer" if it is v2; else the highest v2 q); `patched =
  restore_chain(merge_patch(raw, patch))` (followup.py:15-24: RFC 7386, lists replaced whole); `_prepare(patched,
  provider, target=path, complete=True)` (re-drills every lines block); `redraw_job(h, pre, q2, owned, path)` in a second
  user-Excel job. Between the two jobs the turn lock is released and the workbook may change.
- `fin_refresh` (server.py:289-316): one `read_stored` job, then for each `v2_questions(stored)` (followup.py:27-28)
  **sequentially**: `_prepare(stored[q]["spec"]["v2"], ...)` (no `complete`), `redraw_job(h, pre, q,
  stored[q].get("owned"), path)` as its own job; a failure is `{"q", "ok": False, "error"}` and the next question
  runs; a `RestoreError` skips all later questions.
- `fin_critic` uses `read_stored` + `find_question` only (server.py:324-343).

### 8.2 Owned ranges (`xl/owned.py`)

"Owned" = the block smartass drew on the Answer, stored with the spec as `{"rows": plan.last_row, "cols":
col_number(plan.last_col)}` (pipeline.py:492; `last_col` includes the helper columns). A redraw:

- **clears** `A1:Cells(rows, cols)` (formats and values) and deletes the charts named `"smartass chart"*`; without an
  owned range it clears the whole sheet and every chart (xl/owned.py:40-53);
- **makes room**: when the new block is taller (wider) than the old one and the band just below (right of) the old
  block holds anything (`WorksheetFunction.CountA`), it inserts whole rows (columns) so user cells move instead of
  being overwritten (`growth` / `make_room`, xl/owned.py:12-37). Never shrinks;
- **backs up** only the Answer: hidden same-workbook copy "~Q<n> Answer" (xl/owned.py:56-64); restore = delete the
  broken sheet, un-hide + rename the copy (xl/owned.py:67-72; `pipeline._restore` handles a missing original,
  pipeline.py:300-311). The Data sheet is never copied, deleted or renamed: its table is rewritten in place
  (`rewrite_data` / `rewrite_grid`, render/excel.py:91-147 - header and body as two rectangles, because one `Value2`
  over header + body deletes the table, render/excel.py:94-96), so the table name and every formula that points at it
  survive; the old grid is the restore. Lines tables are reused in place by name (§6.1). Nothing on the Data sheet
  above the table survives (`_place` inserts/deletes whole rows above it and clears them, render/excel.py:136-147).

### 8.3 `_svmeta`: exact structure

Hidden sheet `_svmeta` (xl/sheets.py:157). One row per stored spec (xl/sheets.py:191-210):

```
A: 'SPEC Q<n>          (text, apostrophe-prefixed)
B: '<json chunk 1>    C: '<json chunk 2> ...   (30 000 characters per cell; a cell holds at most 32 767)
json = {"spec": {"v2": <raw AnalysisSpec v2 dict, chain-completed>}, "ids": {}, "owned": {"rows": R, "cols": C}}
```

- Written by `_save_spec` (pipeline.py:233-238): creates the sheet when missing (the v1 `save_spec` would return
  silently, xl/sheets.py:194-195), then `save_spec(t, q, {"v2": raw}, ids={}, owned=owned)`: reads `UsedRange.Value`,
  finds the row whose A is `SPEC Q<n>` (row = `UsedRange.Row + i`), clears it, or appends below the last used row of A;
  writes A and every chunk with **one COM call per cell** (xl/sheets.py:208-210).
- Read by `_stored_specs(t)` (xl/sheets.py:213-225): `UsedRange.Value`; every row with A matching `^SPEC Q(\d+)$` ->
  `json.loads("".join(B..))` -> `{q: {"spec", "ids", "owned"}}`. `v2_questions` keeps those whose spec holds `"v2"`.
- The same sheet also holds v1 rows (a pull's sheet name in A + its layout JSON, `save_meta`, xl/sheets.py:160-178),
  still written by the ad hoc fallback (xl/pull_adhoc.py:127-131).


## 9. Tests

Counts are `def test_` functions (`rg -c`), parametrised cases counted once.

### 9.1 How Excel is faked

- `engine/tests/fakexl.py` (735 lines): an in-memory Excel object model - `Workbook` (engine/tests/fakexl.py:634), `Worksheet`
  (engine/tests/fakexl.py:505), `Range` (engine/tests/fakexl.py:173: values, formats, fonts, borders, styles, `FormatConditions`, `Insert/
  Delete/Clear/AutoFit/AutoFilter/SpecialCells`; refuses a formula on a missing table, `_check_tables` engine/tests/fakexl.py:351),
  `ListObject` (engine/tests/fakexl.py:416), `Application` (engine/tests/fakexl.py:614) with an nb-NO `INTERNATIONAL` tuple (engine/tests/fakexl.py:19),
  names, hyperlinks, windows, sheet `Move`; `snapshot(wb)` (engine/tests/fakexl.py:700) gives a comparable state; `count_com`
  (engine/tests/fakexl.py:706) counts `jobs.com` calls; `freeze_time` (engine/tests/fakexl.py:731).
- Pinned states captured from the writers before batching: `engine/tests/fixtures/excel_calls/` `arrange_ok.json`,
  `arrange_moved.json`, `report_first.json`, `report_redraw.json`, `lines.json`, `lines_redrill.json` (used by
  engine/tests/test_excel_calls.py:89-117). These are the closest thing to a "golden look" of `apply_report` and
  `write_lines` and could seed a file-writer comparison (they describe the COM end state, not a file).
- Hand-made fakes per file: `Rng/Lists/WS/LO/WSLO` (engine/tests/v2/test_v2_excel_writer.py:7-119); `Rec/Coll/FakeChart/
  FakeShapes/FakeWS` recording every chart property (engine/tests/v2/test_v5_charts.py:190-283); `Sheet`
  (engine/tests/v2/test_v5_check.py:33); `Shape/_Sheet` (engine/tests/test_owned.py:22-48); `Sheet/Sheets/WB/T`
  (engine/tests/v2/test_v2_save_spec.py:6-32); `WB/NoAutoSaveWB/T` (engine/tests/v2/test_v2_autosave.py:5-24); `FakeCells`
  (engine/tests/v2/test_v2_readback.py:5).
- Seam tests that monkeypatch every pipeline helper and assert the call order: `wire()` in
  engine/tests/v2/test_v2_pipeline_job.py:16-39 (asserts the exact `excel_job` order at :47-49) and
  engine/tests/v2/test_v2_redraw.py:17-56.
- The test set (`smartass/testset/runner.py`) only calls `prepare()` and compares `key_values`
  (testset/runner.py:40-45, 120-135): it never runs a writer, so "every test-set question drawn both ways" (design
  phase 3 proof) has no harness yet.

### 9.2 Per area

| area | test file (count) | fakes | if the COM writer is replaced |
|---|---|---|---|
| job order / seams | v2/test_v2_pipeline_job.py (9), v2/test_v2_redraw.py (5) | monkeypatched helpers | **rewrite**: the order and helper names change |
| Answer writer | v2/test_v2_excel_writer.py (11: grid, text rules, rewrite_grid), v2/test_v5_writer.py (11: status cell, print setup, widths) | hand fakes, fakexl | **rewrite** (grid/text rules can move to the file writer; rewrite_grid tests go with in-place redraw) |
| styling look | test_excel_calls.py (7, pinned fixtures + call counts), v2/test_v5_layout.py (26, 2 use `apply_report` on fakexl at :282-323), v2/test_v5_keypoints.py (20, 1 at :171-197) | fakexl | **rewrite** the apply_report ones; the pinned fixtures are a reference for the expected look |
| Data sheet | v2/test_v5_datasheet.py (16: ~5 pure `header_rows`, ~11 `write_data/rewrite_data` on fakexl) | fakexl | pure ones stay; writer ones rewrite |
| Lines sheet | v2/test_v5_lines_header.py (6), test_excel_calls.py lines cases | fakexl | rewrite |
| Lines formulas | test_drill.py (92; only ~10 touch `drill_xl`: `table_name`, `drill_formula`, `safe_for_sumifs`, :155-207, :1186), v2/test_v2_lines_expr.py (3) | none | stay |
| Charts | v2/test_v5_charts.py (33: 14 layout/ChartSpec tests at :39-178, 19 from :291 on that drive `write_charts` with `FakeWS` or check ChartSpec details), v2/test_v5_chart_block.py (27, 1 seam test), v2/test_v5_live_round2.py (21: mostly layout; :61, :85, :102 check chart COM settings, :264 print setup), v2/test_v2_layout_charts.py (6, pure), v2/test_v5_locale_formats.py (10: locale helpers, union refs, the source-scanning rule at :28) | FakeWS | pure ones stay; the COM-drawing ones must become assertions on the XlsxWriter chart objects (or the built file) |
| Sources | v2/test_v5_sources.py (19, fakexl + seams) | fakexl | stay while the log stays COM; seams at :188-226 change |
| Read-back | v2/test_v2_readback.py (6), v2/test_v2_pct_floor.py (8), v2/test_v5_material.py (11) | FakeCells | stay (`compare` is pure) |
| Layout check | v2/test_v5_check.py (13) | `Sheet` | stays if `check` keeps its signature |
| Owned / backup | test_owned.py (4), parts of v2/test_v2_redraw.py | hand fakes | **delete** with `xl/owned.py` |
| _svmeta | v2/test_v2_save_spec.py (3), v2/test_v2_followup_target.py (10), v2/test_v2_pipeline.py (10, `read_stored` parts) | hand fakes | stay (unchanged) or adapt |
| Sheet order | test_sheetorder.py (15), v2/test_v2_sheetorder_data.py (1) | fakexl | pure planning stays; `arrange_sheets` stays if still used |
| AutoSave | v2/test_v2_autosave.py (3) | hand fakes | stay |
| Pure layout | v2/test_v2_layout.py (7), v2/test_v5_layout.py (24 of 26), v2/test_v5b_presentation.py (29), v2/test_v5d_audit.py (28), v2/test_v2_short_formulas.py (2), v2/test_v5_headers.py (10), v2/test_v5_keypoints.py (19 of 20), block tests | none | stay; they are the contract the file writer is built against |
| Server paths that call the jobs | v2/test_v2_server.py (15), v2/test_v2_server_drill.py (10), v2/test_v2_critic_brief.py (8), v2/test_v2_answers.py (12), test_runlog.py (7), v2/test_v2_cli_xl.py (9) | monkeypatched `excel_job`/`redraw_job` | stay if the job signatures stay |


## 10. Risks and surprises for the file-built approach

Ordered roughly by how much they change the plan.

1. **The Sources log is not in design §2, and `_svmeta` is named there but was not in the spike.** Both are workbook-wide and append/merge into what is
   already there: the log reads its last row and earlier rows (xl/sources.py:90-98), rewrites Known limits from the
   existing lines (xl/sources.py:75-87, 129-135), and records calculated OK/DIFF values (pipeline.py:433-439); the
   Data header's Sources cell gets a hyperlink to a row number known only after that read (xl/sources.py:159-161).
   `_svmeta` is read by every follow-up (xl/sheets.py:213-225). Neither can come from the built file; they remain COM
   steps (Sources today ~30-45 calls, `_svmeta` one call per 30 000-character chunk, xl/sheets.py:208-210) and must be
   counted against the 2 s gate. The Sources sheet also has its own `AutoFit` (xl/sources.py:179) and freeze
   (xl/sources.py:186).
2. **q is needed before the file is built, but comes from the live workbook.** The built file carries q in sheet names,
   the Data title (render/datasheet.py:118), the Source note (render/layout.py:544), the `CubeRef` sheet references
   (model/expr.py:331-332) and the table names in the formulas left out. `next_question` reads the sheet names
   (xl/sheets.py:30-32). So the build must run inside the user-Excel job after reading the names (holding the turn
   while XlsxWriter writes: 2.4 s for 50 000 x 12 by the design's measure), or q must be read in an earlier job and
   re-checked (the turn is released between jobs, xl/host.py:163-166; `read_stored` already is such a separate job for
   follow-ups, server.py:268).
3. **Name collisions behave differently.** Today `Worksheets.Add` + `Name = "Q<n> Data"` fails on a taken name and the
   job rolls back; `_write_lines` refuses a renamed Lines sheet (pipeline.py:380-382). `next_question` does not count
   "Q<n> Data" or a topic sheet without its "Q<n> Answer" (xl/sheets.py:16-27), so a leftover "Q5 Data" (Answer deleted
   by hand) gives q = 5 again. `Worksheets.Copy` instead renames silently to "Q5 Data (2)" (Excel behaviour, **to
   verify**), after which every by-name step (Formula2 targets, `CubeRef` written later, Sources links) hits the wrong
   sheet. Same for table names: `ListObjects.Add` + `Name = "Data_Q<n>"` fails if any sheet still holds that table
   name (table names are workbook-unique), and `find_table` reuses a `Lines_Q<n>_<alias>` found on **any** sheet
   (xl/drill_xl.py:127-133, 220-234). The copy step needs an explicit pre-check of sheet and table names.
4. **The look reads Excel's calculated values.** Key point heights/merges/fit lint (xl/styling.py:374, 419-426,
   573-595), the numeric fallback style (xl/styling.py:455-456), Data widths by `AutoFit` (render/excel.py:212-220),
   the thin waterfall label lift via `DataLabel.Top` (render/charts.py:215-220). A file writer has to derive these from
   the model (`res.cells`) or change the rule; "same formats on the Answer sheet" (phase 3 proof) will otherwise differ
   in row heights of the Key points.
5. **Today Excel calculates during the write.** No `fast_excel` around Data, Answer, styling or charts (it is only in
   xl/drill_xl.py:222, xl/sources.py:128, xl/sheetorder.py:185); `_calculate` recalculates **every open workbook**
   (`Application.Calculate()`, pipeline.py:229) and `wait_calculated` gives up silently after 60 s
   (xl/sheets.py:316-326), after which the read-back may see stale values and report DIFF.
6. **Target resolution still drags in deleted machinery.** Every job starts the SSO watcher (xl/target.py:105) and
   loads the Smart View VBA bridge into the user's Excel (`_attach_bridge`, xl/target.py:79-92, 256), which may need the
   hidden Excel to build it (xl/instances.py:236-248). The answer never uses the bridge. Cutting `resolve_target` loose
   is a prerequisite (design "left to the builder").
7. **Rollback gaps that the move-in step must close.** `rollback` deletes only `t.sheets_added` (xl/sheets.py:52-55),
   which only `add_sheet` and the Sources creation append to (xl/sheets.py:128, xl/sources.py:62-64) - copied sheets
   would not be in it. `_svmeta` creation is untracked (pipeline.py:235-237). A failed redraw does not restore Lines
   tables rewritten in place, Sources rows already appended, or an `_svmeta` row already saved (pipeline.py:575-580).
8. **Whole-sheet replacement drops behaviour the code deliberately has.** The redraw keeps the Data table and its name
   so "every formula that points at it" survives (render/excel.py:91-98, pipeline.py:559-562), pushes user cells on the
   Answer aside instead of overwriting them (xl/owned.py:12-37), and only clears its own charts (xl/owned.py:50-53).
   The owner's 19.11 decision accepts losing references into Q sheets, and it also covers user cells typed onto a Q
   Answer ("the user's own notes and formulas live on the user's own sheets"): they are now deleted with the sheet - a
   behaviour change the decision accepts, worth one line in the release note. A follow-up that drops a lines block
   leaves its old "Q<n> Lines b<i>" sheet today (no code deletes Lines sheets); whole replacement must decide which old sheets belong to Q<n>
   (`question_of` / `_OURS` patterns, xl/sheets.py:16-27, xl/sheetorder.py:22).
9. **Where copied sheets land.** `add_sheet` puts new sheets after the q group and before Sources/_svmeta
   (xl/sheets.py:114-127). "After the workbook's last sheet" is usually after the hidden `_svmeta` and after Sources.
   `arrange_sheets` in "active" mode (the user's workbook) only reorders our sheets within the slots they occupy
   (xl/sheetorder.py:85-87) and pins hidden sheets (xl/sheetorder.py:90-96), so the result differs from today's order;
   whether `Copy(After=<hidden sheet>)` works and keeps the copies visible is **to verify**.
10. **Active sheet and window side effects.** `_freeze` needs a visible window and activates the sheet and back
    (xl/styling.py:195-276); `_freeze_below` activates the workbook and selects A5 (xl/drill_xl.py:198-204);
    `arrange_sheets` activates the Answer (xl/sheetorder.py:203-208); the design's temp file open/copy will activate
    sheets too. With the file approach freeze panes and gridlines can be set in the file, removing three of these.
11. **Read-only / archive / protected targets.** Archive paths are refused by path before opening
    (xl/target.py:123-124, 253-254; `allow_archive=False` hard-coded at pipeline.py:213). Nothing in the answer path
    checks `wb.ReadOnly`, `ProtectStructure` or sheet protection (only `xl/workbench/books.py:134` reports ReadOnly):
    today a read-only workbook gets written and only our own file's `Save` would fail. `Copy` into a
    structure-protected workbook fails like `Worksheets.Add` does.
12. **Sensitivity label and AutoSave.** New analysis files carry the label (xl/target.py:200, xl/instances.py:272-284;
    seed check xl/target.py:386-395). The XlsxWriter temp file has none; it is opened read-only and never saved, so a
    label prompt is unlikely but **to verify** on the owner's tenant. Keep the temp file outside OneDrive. AutoSave
    stays paused by `autosave_paused` (pipeline.py:260-281); for a user's workbook it is left off with a note.
13. **XlsxWriter traps** (all **to verify** against the version used). `worksheet.write()` turns a string starting
    with "=" into a formula by default (XlsxWriter's `strings_to_formulas`, from memory); use `write_string` for every text cell, matching `plan.literal` and the apostrophe rule,
    render/excel.py:261-262, xl/drill_xl.py:93-105); no named cell styles (§4.1); formulas are written without cached
    results, so the move-in must calculate before any read; non-contiguous chart ranges and per-point label positions
    (§5.6). Formulas the file keeps are legacy (non-Formula2) formulas: none of the cell-only formulas listed in §2.1
    uses arrays, but this is **to verify** against Excel 365's implicit-intersection "@".
14. **Copy and links.** Formulas or chart series that reference a sheet not in the copied set become external links to
    the temp file (the design's "break the link"); `CubeRef` (`'Q<n> Lines b<i>'!$D$3`) is the only cross-sheet
    reference on the Answer, so it is safe only if the Lines sheets are copied in the same call. Sheet-level names
    (`smartass_pagesetup`, Print_Area, Print_Titles) travel with the sheets.
15. **Unformatted cells take the destination's Normal style.** The style lint flags a block whose `Font.Name` is not
    uniformly Arial (xl/styling.py:562-565); every cell of `A1:<last_col><last_row>` (blank ones too, today "TN Tom")
    needs an Arial format in the file.
16. **Things that only look like problems.** The v1 helpers in xl/drill_xl.py (`drill_formula`, FILTER forms) are not
    used by v2 answers; formula checks (BLOCK) do not look at formats (§4.1); the test set never touches Excel
    (testset/runner.py:40-45); the Sources hyperlinks hold sheet names as text (`'Q<n> Data'!A1`, xl/sources.py:158), so
    a replaced sheet with the same name keeps them working (Excel does not rewrite a hyperlink's SubAddress on rename,
    as far as I know - **to verify**).
17. **Stage names are part of the gate.** `scripts/bench_answers.py:33-35, 110-112` reads `excel job.draw.data sheet`,
    `...draw.answer`, `...draw.sources`, `excel job.save`; new stages (build, move in, tables, formulas) need columns there
    or the before/after comparison breaks.
