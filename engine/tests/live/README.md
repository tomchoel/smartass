# Live checks (run before every release)

These need desktop Excel, Smart View, an EPM login and the network, so CI cannot run them. Run them on a
maintainer PC with the MCP connected (Copilot CLI or Claude Code), as tool calls. The scripts in this folder
(`live_smoke.py`, `live_build.py`, `live_drill.py`, `live_relay.py`, ...) do the same over stdio.

1. **`sv_query` on each cube** - a small grid comes back with numbers.
   - FinMax:
     `sv_query(connection="FinMax", rows=[{"dimension":"All Accounts","members":["ac.C30","ac.EBITDA"]}],
     columns=[{"dimension":"All Time Periods","members":["Aug 26","Sep 26"]}],
     pov={"All Entities":"un.SY.2WHS","All Scenarios":"Actual","All Currencies":"NOK","All Balance Types":"Accumulated"})`
   - FinLite: the same call with `connection="FinLite"`.
   - FinPlan:
     `sv_query(connection="FinPlan", rows=[{"dimension":"Entity","members":["BANORWAY"]}],
     columns=[{"dimension":"Years","members":["FY25","FY26"]},{"dimension":"Period","members":["TotalYear"]}],
     pov={"Scenario":"Target 2025","Version":"Final"})`

   Expected: no error, numeric values in the table.
2. **`sv_ws_build`** - one small analysis, e.g. the FinMax pull above as `pulls=[{"id":"m", ...}]` with a short
   spec (see `live_build.py`), `target="session"`. Expected: a live Smart View sheet and a `Q1 Svar` sheet, saved in
   your `analyser` folder (`paths.analyser`).
3. **`sv_drill`** on FinLite ac.12881 x un.SY x Sep 26:
   `sv_drill(question="live check", cell={"All Accounts":"ac.12881","All Entities":"un.SY","All Time Periods":"Sep 26"})`.
   Expected: **95 lines, sum equal to the cell, difference exactly 0.**
4. **Relay reload** - save `engine\smartview_mcp\config.py` without changes, make one tool call (e.g. `sv_status`),
   and check `%LOCALAPPDATA%\smartass\logs\proxy.log`. Expected: exactly one restart and no error.
5. **Doctor bundle** - run `smartass doctor --bundle` and open the zip. Expected: no `sso_token=` value and not the
   `http.secret` value from `config.toml` anywhere in it.
