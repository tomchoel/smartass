# Troubleshooting

Start with `smartass doctor`. It prints one line per check (`OK`, `WARN` or `FAIL`), and under every line that is
not OK a `fix:` line. The sections below follow the checks in the order the doctor runs them. The fix wording is
the same as in the doctor (`engine\smartview_mcp\doctor.py`).

User settings live in `%LOCALAPPDATA%\smartass\config.toml`; defaults in `engine\defaults.toml`.

`smartass` not found? Open a new terminal after install (the installer adds it to your user PATH). Otherwise run it
by its full path: `%LOCALAPPDATA%\smartass\app\engine\.venv\Scripts\smartass.exe`.

## config

- **Symptom:** `FAIL  config  <TOML error>` or `WARN  config  no ...\smartass\config.toml`.
- **Cause:** `config.toml` has a syntax error, or the installer has not run for this user.
- **Fix:** FAIL: Fix the TOML syntax, or delete config.toml and re-run installer\install.ps1.
  WARN: Run installer\install.ps1.

## excel

- **Symptom:** `FAIL  excel  not found: <path>`.
- **Cause:** desktop Excel is not installed, or it is installed in a non-standard place.
- **Fix:** Install desktop Excel (Microsoft 365), or set paths.excel_exe in config.toml.

## smartview

- **Symptom:** `FAIL  smartview  HsAddin.dll not found at <path>`.
- **Cause:** the Oracle Smart View add-in is not installed, or is installed in a non-standard place.
- **Fix:** Install Oracle Smart View (Company Portal), or set paths.smartview_dll in config.toml.

## epm

- **Symptom:** `FAIL  epm  https://... unreachable: <error>`.
- **Cause:** no network route to Oracle EPM Cloud (off the Telenor network, VPN down, or a proxy problem).
  An OK here only means the server answers; the login itself is checked on the first query.
- **Fix:** Connect to the Telenor network (or VPN) and check the proxy.

## drill_browser

- **Symptom:** `FAIL  drill_browser  msedge: ...; chrome: ...`.
- **Cause:** `sv_drill` needs a hidden Edge or Chrome (`drill.channels`). None is installed, or a policy blocks it
  (Chrome's UserDataDir policy blocks a hidden Chrome, which is why Edge is tried first).
- **Fix:** Install Microsoft Edge, or ask IT why the Edge UserDataDir policy is set.

## atlas

- **Symptom:** `FAIL  atlas  missing: ...\atlas.db` or `WARN  atlas  ...\atlas.db (N days old)` (older than 35 days).
- **Cause:** the member atlas (used to find members by name) was never copied, or is out of date.
- **Fix:** FAIL: Re-run installer\install.ps1 (copies the monthly atlas), or call sv_atlas_build.
  WARN: Re-run the installer after the monthly release, or call sv_atlas_build.

## output

- **Symptom:** `FAIL  output  cannot write <folder>: <error>`.
- **Cause:** the folder for answer workbooks (`paths.analyser`) does not exist and cannot be created, or is
  read-only (for example OneDrive not signed in).
- **Fix:** Set paths.analyser in config.toml to a folder you can write.

## http_api

- **Symptom:** `WARN  http_api  no http.secret in config.toml` or `WARN  http_api  not answering on 127.0.0.1:<port>`.
- **Cause:** the installer has not written the per-user secret, or the background engine (`smartass serve`,
  started by the Startup entry `smartass-engine.lnk`) is not running.
- **Fix:** no secret: Run installer\install.ps1.
  Not answering: Log off and on (Startup entry), or run `smartass serve`.

## update

- **Symptom:** `WARN  update  <error>` or `WARN  update  <current> -> <target>`.
  (`OK  update  release folder not configured` is normal until `release.folder` is set.)
- **Cause:** the release folder cannot be read (not synced yet), or a newer release is available.
- **Fix:** error: Check that the release folder is synced (OneDrive).
  Newer release: Run `smartass update` (or log off and on).

## Any check: `FAIL  <check>  <ExceptionType>: <message>`

The check itself crashed. Fix: Send the file from `smartass doctor --bundle` to the maintainer.

## Smart View asks me to log in all the time

The SSO watcher clicks "Telenor SSO" in Smart View's login dialog for you. It starts hidden at logon from the
Startup entry `smartass-sso-watcher.lnk` and logs to `%LOCALAPPDATA%\smartass\logs\sso_watcher_ps.log`.

- Check the log: no recent lines means the watcher is not running.
- Restart it: double-click `smartass-sso-watcher.lnk` in the Startup folder (`Win+R`, `shell:startup`), or log off
  and on. If the entry is missing, re-run `installer\install.ps1`.

## "Smart View showed its sign-in window (WebView2 failed …)"

Smart View's login page runs in WebView2. When WebView2 cannot start, Smart View says so ("WebView2 Runtime failed …
switching to IE") and opens the login in Internet Explorer, where the Oracle page fails ("Script Error: Function
expected"). That window cannot be closed, so the engine kills its own hidden Excel, starts a new one and retries the
call once. If the second try fails the same way, you get this error.

- Smart View keeps one WebView2 folder for all Office apps (`%APPDATA%\Oracle\SmartView\WebView`). If the error names
  an app that holds it (e.g. `POWERPNT.EXE`), close that app and retry.
- In your own (visible) Excel the engine recovers by itself since 2026-10-23: when the login has fallen back to
  Internet Explorer, it ends Smart View's stale WebView2 of that Excel (or an orphan whose app is gone; never another
  app's) and closes the broken IE windows, and Smart View reopens the login in WebView2 (at most twice per call;
  `server.log`: "fell back to Internet Explorer; ended stale WebView2").
- Sign in once in your own Excel (Smart View ribbon → Panel → connect), then retry.
- Still failing: run `smartass doctor --bundle`.

## Known limits (accepted by the owner, 2026-10-23)

These were investigated in the reliability audit (`docs\audit\2026-10-23-reliability-audit.md`) and accepted as they
are. If one of these symptoms shows up, the cause below is already known; no new investigation is needed.

| Symptom | Known cause | What to do |
|---|---|---|
| Logins are clicked late (18–200 s), pulls are 2–3× slower than usual (`open default grid` 7–167 s), "The Oracle EPM login dialog is still open" | Another Office app, usually PowerPoint, holds Smart View's WebView2 folder. The SSO watcher's UI Automation then blocks for 47–118 s (`sso_watcher_ps.log`: "loop stalled"). | Close PowerPoint (and restart a long-running Excel), then retry. Measured: S2.build median 39.6 s → 25.8 s. |
| The first FinMax call after a server start takes 15–30 s | It starts the hidden Excel (about 13–19 s) and does the hourly SSO login. | None; later calls are fast. |
| A build in a newly started Excel window takes 35–50 s in `open target.new workbook` | A cold Excel start plus SSO for a new interactive Excel (a parallel agent, or no Excel open). | None; the next build in that Excel takes about 2 s. |
| Progress notes (`still working: …`) keep arriving after a call to your Excel timed out | The abandoned job's COM call is still hung; its ticker runs until that call returns. The client drops notes for a finished request. | Harmless. It stops when Excel answers or closes. |
| Many `msedgewebview2.exe --type=crashpad-handler` processes for the Smart View WebView folder, with their parent gone | Each forced restart of the engine's hidden Excel leaves Smart View's WebView2 crash reporter behind. | Harmless (they do not lock the folder). They disappear at log off; ending them by hand is safe. |
| Matrix: `S9.lookup` is LIMIT | A fresh-profile lookup needs a real signed-in Excel; the COM Excel inherits the simulated environment. | Verify by hand on a new machine (colleague rollout). |
| Matrix: `S10.5.textfixes` is LIMIT | The user override of a shipped text fix is verified in-process; no current drill line hits a shipped key. | None. |
| Matrix: `S10.6.pack` is LIMIT | The `produktregnskap` pack has no runtime consumer yet; only its wiring is checked. | Roadmap: give it a consumer or remove it. |
| Matrix: `S10.8.coldexcel` is LIMIT | A cold start needs no Excel running, and the owner's Excel is open during runs. | Run it with all Excels closed if needed. |
| Matrix: `S5.drill.report` FAILs with "needs S#.build" under `--only drill` | It builds on S2.build's workbook from the same run. | Include `S2.build` in `--only`, or run the full matrix. |

## Still stuck

Run `smartass doctor --bundle` and send the zip it prints to the maintainer. It contains the doctor result,
versions, your `config.toml` and the tail of the logs. Tokens and passwords are redacted, but it does contain
file paths (including your user name).
