# Open items after Stage 2+3, sub-units and the speed pass (2026-10-26)

Parked by the user on 2026-10-26: "we will deal with them another day". State of `main` at the time: f2b13d3,
723 tests pass, ruff clean.

1. **Full live matrix not rerun.** Only S3, S11, S12, `warehouse_live.py` and `test_styling_live.py` ran after
   today's changes. The block write in `answer._Cells` (d15df6a) is used by every answer sheet, so rerun
   `PY tests\live\matrix.py` (S0-S12, ~30-45 min, its own Excel): above all S2 `sv_ws_build`, S4 concepts,
   S5 drill reports, S7 refresh.
2. **Template answers SLOW (no FAIL).** S3 on 2026-10-26: 42-121 s against a 60 s budget (earlier runs 17-45 s).
   The time is a cold Excel start (~67 s on the first scenario) and Smart View pulls (~7-10 s per sheet: "open
   default grid" ~2.5 s, retrieve ~2 s, style ~1 s). Skipping the default grid crashed Excel on 2026-10-14
   (`userexcel.py`, "open default grid"), so look elsewhere: fewer pull sheets per answer, "G page setup" (~2-4 s),
   "H lint" (~1 s).
3. **MCP tool list.** `sv_warehouse` and `fin_pnl` for sub-units appear in a client only after a new session or
   an MCP reconnect.
4. **Freeze panes is not checked live any more.** `test_styling_live.py` prints SKIP when its own DispatchEx Excel
   has no grid area (VisibleRange one cell), which is every run from a console so far. Real answers do freeze
   (the "Svar" sheets checked with openpyxl on 2026-10-26). A check on a real answer workbook (e.g. in S11: read
   `freeze_panes` of the saved file) would close this.
5. **Untracked folder** `engine\Microsoft\` in the main checkout. Not inspected; find out what writes it, then
   delete it or add it to `.gitignore`.
6. **Planned for later (see the Stage 2+3 plan, "Not in this plan"):** the shared warehouse and the policy for
   cube data on local disk (confirm with IT), the balance sheet, capex, price/volume/mix bridges, cost-centre
   anomalies, `numfmt.py`, the other `sv_analyse` templates on the engine, Stage 4-6.
