# Troubleshooting

Start with `smartass doctor`. It prints one line per check (`OK`, `WARN` or `FAIL`), and under every line that is
not OK a `fix:` line. The sections below follow the checks in the order the doctor runs them. The fix wording is
the same as in the doctor (`engine\smartass\ops\doctor.py`).

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

## smartview_startup

- **Symptom:** `WARN  smartview_startup  Smart View checks for extension updates at every Excel start ...` (an EPM
  login window ~20 s after any Excel opens; smartass then reports Excel as busy until it is clicked).
- **Cause:** Smart View's extension update check at Office start (`HKCU\Software\Oracle\SmartView\extensions`
  `CheckOnStartup=1`) needs an EPM session, so it opens the login (found 2026-11-03 in Smart View's own traffic log).
- **Fix:** close Excel; Smart View > Options > Extensions: untick "Check for updates when Microsoft Office starts".
  Smart View then signs in only when you use it; extension alerts on connect and the manual update check still work.
- **Symptom:** `WARN  smartview_startup  Smart View diagnostics capture is on ...`
- **Fix:** Smart View > Options > Advanced: turn diagnostics off (`Profile=0`), then delete
  `%APPDATA%\Oracle\SmartView\Diagnostics` (it holds sign-in cookies in plain text).

## vba_access

- **Symptom:** `FAIL  vba_access  Excel blocks access to the VBA project object model ...`.
- **Cause:** the hidden Excel loads its Smart View bridge through the VBA project, which Excel refuses unless
  `HKCU\Software\Microsoft\Office\16.0\Excel\Security` `AccessVBOM=1` (a value under `Software\Policies` wins).
- **Fix:** close Excel; File > Options > Trust Center > Trust Center Settings > Macro Settings: tick "Trust access to
  the VBA project object model". A policy value needs IT.

## epm

- **Symptom:** `FAIL  epm  https://... unreachable: <error>`.
- **Cause:** no network route to Oracle EPM Cloud (off the Telenor network, VPN down, or a proxy problem).
  An OK here only means the server answers; the login itself is checked on the first query.
- **Fix:** Connect to the Telenor network (or VPN) and check the proxy.

## oauth

The bearer route: on a pod that has a scope, grids are fetched over plain HTTPS with your own OAuth access token, so
no hidden browser has to sign in there first. It is optional; without it, and whenever it cannot authenticate, grids
sign in through the hidden browser as before. Drill-through uses the browser (over the token it is untested).

Measured 19.11.2026 with scopes for plan1 (FinMax, FinLite) and stratmodel (FinPlan): from a cold start a FinMax
grid took 0.7 s, a FinPlan grid 0.7 s and the closed month 0.1 s, with no browser started; through the browser's
sign-in the first grid took about 20 s. Every grid compared was the browser route's cell for cell. A drill link
(GetDrillThroughReports) also goes over the token; the drill lines themselves still load in the hidden browser
(a 2 839-line drill: 28 s including the browser's start, reconciled exactly).

- **Set-up (per user, `config.toml`):** `[oauth]` `idcs` (the identity domain's base URL) and `client_id` (a public
  client an administrator registered there, with the device-code and refresh-token grants), and one scope per pod
  under `[oauth.scopes]`: `"https://<pod host>" = "urn:opc:serviceInstanceID=<id>urn:opc:resource:consumer::all"`.
  A token covers one environment, so a pod without a scope keeps using the browser
  (`oauth  ...: no scope (hidden browser)`). The administrator adds an environment to the client as a resource and
  reads its scope there.
- **Symptom:** `WARN  oauth  <pod>: not signed in` or `<pod>: sign-in ended`.
- **Cause:** no sign-in was made on this PC, or it was not used for 7 days (Oracle's lifetime for a refresh token;
  the MCP server and `smartass serve` renew it daily while they run), or the identity domain withdrew it.
- **Fix:** Run `smartass login`: a page opens in your browser and the code is on the clipboard; paste it and
  approve with Telenor SSO. `smartass login --status` shows the sign-ins, `--forget` deletes them from this PC.
- The sign-in is stored per pod in `%LOCALAPPDATA%\smartass\oauth\`, DPAPI-protected for your Windows user; smartass
  never sees a password. The token is only ever sent with the read-only Smart View requests.
- Log line `svhttp: no bearer route to <pod> (...); using the browser route`: the reason is in the brackets; the
  bearer route is tried again after two minutes.

## drill_browser

- **Symptom:** `FAIL  drill_browser  msedge: ...; chrome: ...`.
- **Cause:** `fin_drill` (and lines blocks) need a hidden Edge or Chrome (`drill.channels`). None is installed, or a policy blocks it
  (Chrome's UserDataDir policy blocks a hidden Chrome, which is why Edge is tried first).
- **Fix:** Install Microsoft Edge, or ask IT why the Edge UserDataDir policy is set.

## atlas

- **Symptom:** `FAIL  atlas  missing: ...\atlas.db` or `WARN  atlas  ...\atlas.db (N days old)` (older than 35 days).
- **Cause:** the member atlas (used to find members by name) was never copied, or is out of date.
- **Fix:** FAIL: Re-run installer\install.ps1 (copies the monthly atlas), or run `smartass atlas build`.
  WARN: Re-run the installer after the monthly release, or run `smartass atlas build`.

## output

- **Symptom:** `FAIL  output  cannot write <folder>: <error>`.
- **Cause:** the folder for answer workbooks (`paths.analyser`) does not exist and cannot be created, or is
  read-only (for example OneDrive not signed in).
- **Fix:** Set paths.analyser in config.toml to a folder you can write.

## http_api

- **Symptom:** `WARN  http_api  no HTTP API secret (secret.bin)` or `WARN  http_api  not answering on 127.0.0.1:<port>`.
- **Cause:** the installer has not written the per-user secret (`%LOCALAPPDATA%\smartass\secret.bin`, DPAPI), or
  the background engine (`smartass serve`, started by the Startup entry `smartass-engine.lnk`) is not running.
- **Fix:** no secret: run `smartass setup` (it also moves an old plain `http.secret` out of config.toml).
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

## Sign-in messages (login hardening, 2026-11)

The engine tells you what it is doing within seconds and never closes or restarts your Excel.

| Message | Meaning | What to do |
|---|---|---|
| "Smart View's Login is open in your Excel; waiting for the automatic SSO click" (note, after about 5 s) | The Login dialog in your Excel is waiting for the SSO watcher. | Usually nothing; if it stays, click "Telenor SSO" yourself or check the watcher (above). |
| "The Smart View Login window in your Excel is not responding; nothing was closed" (after about 10 s) | Windows marks the Login window "Not responding". | Close that Login window (or that Excel) yourself, then ask again. |
| "Smart View's sign-in window stopped responding (Windows marks it 'Not responding')" | The hidden Excel's Login hung; it was restarted once. | Retry; if it repeats, `smartass doctor --bundle`. |
| "EPM gave an empty response ...; retrying" (note) | A transient Smart View connect error. The hidden Excel retries, then restarts itself once; your Excel retries the pull once. | Nothing, unless the call then fails: ask again; if every connect fails see the cookie row under Known limits. |
| "EPM kept answering 'Empty response from provider' after a restart of the engine's Excel" | The hidden Excel's Smart View session stayed broken after one restart. | Wait a minute and ask again; if it persists, sign in once in your Excel and run `smartass doctor --bundle`. |
| "Smart View in this Excel is not signed in correctly" | Your Excel answers every connect with an empty response and shows no Login. | Sign in once in your Excel (Smart View ribbon → Panel → connect), then ask again. |
| HTTP API `409 busy` | Another `/gate` or `/answer` is running. | Wait and resend. |

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

These were investigated in the reliability audit (`docs\archive\audit\2026-10-23-reliability-audit.md`) and accepted as they
are. If one of these symptoms shows up, the cause below is already known; no new investigation is needed.

| Symptom | Known cause | What to do |
|---|---|---|
| Logins are clicked late (18–200 s), pulls are 2–3× slower than usual (`open default grid` 7–167 s), "The Oracle EPM login dialog is still open" | Another Office app, usually PowerPoint, holds Smart View's WebView2 folder. The SSO watcher's UI Automation then blocks for 47–118 s (`sso_watcher_ps.log`: "loop stalled"). | Close PowerPoint (and restart a long-running Excel), then retry. Measured: S2.build median 39.6 s → 25.8 s. |
| The first FinMax call after a server start takes 15–30 s | It starts the hidden Excel (about 13–19 s) and does the hourly SSO login. | None; later calls are fast. |
| A build in a newly started Excel window takes 35–50 s in `open target.new workbook` | A cold Excel start plus SSO for a new interactive Excel (a parallel agent, or no Excel open). | None; the next build in that Excel takes about 2 s. |
| A new analysis workbook has no TN cell styles (`TN Heltall`, ...) in its style gallery | Since phase 3c the first answer's built file is the workbook; the styles are installed when something needs them. | Run `xl_tidy` on a sheet (it merges them from the TN template), or ignore: the answer's own cells carry their formats. |
| A new analysis file that a failed answer (or a probe) deleted is back in the analyser folder minutes later, online-only | OneDrive restores a file Excel had just saved to the cloud when the local delete races that upload (seen with probe files Excel saved and closed, 20.11.2026). | Delete it again. |
| Progress notes (`still working: …`) keep arriving after a call to your Excel timed out | The abandoned job's COM call is still hung; its ticker runs until that call returns. The client drops notes for a finished request. | Harmless. It stops when Excel answers or closes. |
| Many `msedgewebview2.exe --type=crashpad-handler` processes for the Smart View WebView folder, with their parent gone | Each forced restart of the engine's hidden Excel leaves Smart View's WebView2 crash reporter behind. | Harmless (they do not lock the folder). They disappear at log off; ending them by hand is safe. |
| "Empty response from provider" on every connect, in every Excel (also the engine's hidden one), and **no Login window** appears (not even after Disconnect All) | Smart View's shared WebView2 cookie store holds a broken session, e.g. after a login was interrupted (seen 2026-10-23 after stuck IE-fallback logins). Smart View reuses it silently and EPM rejects it. Not an Oracle block. | Close every Excel, rename `%APPDATA%\Oracle\SmartView\WebView\EBWebView\Default\Network\Cookies` and `Cookies-journal` (e.g. to `*.bak`), then connect: the Login appears and the watcher clicks Telenor SSO. This fixed it at once on 2026-10-23. Since `ca0accd` the engine fails fast with this advice instead of hanging. |
| Matrix: `S9.lookup` is LIMIT | A fresh-profile lookup needs a real signed-in Excel; the COM Excel inherits the simulated environment. | Verify by hand on a new machine (colleague rollout). |
| Matrix: `S10.5.textfixes` is LIMIT | The user override of a shipped text fix is verified in-process; no current drill line hits a shipped key. | None. |
| Matrix: `S10.6.pack` is LIMIT | The `produktregnskap` pack has no runtime consumer yet; only its wiring is checked. | Roadmap: give it a consumer or remove it. |
| Matrix: `S10.8.coldexcel` is LIMIT | A cold start needs no Excel running, and the owner's Excel is open during runs. | Run it with all Excels closed if needed. |
| Matrix: `S5.drill.report` FAILs with "needs S#.build" under `--only drill` | It builds on S2.build's workbook from the same run. | Include `S2.build` in `--only`, or run the full matrix. |

## An Excel step behaves differently since 2026-11-19 (the fast COM wrapper)

- **Symptom:** an Excel step fails or gives an odd result it did not give before 2026-11-19, e.g. an
  `AttributeError` that names a COM object, or a method where a value was expected.
- **Cause to rule out:** since 2026-11-19 the engine drives Excel through `engine\smartass\xl\fastcom.py`: pywin32's
  own late-bound wrapper, except that what it learns about each interface is shared by all objects of it (about
  x1.7 faster Excel jobs).
- **Check:** set the environment variable `SV_FASTCOM=0` for the smartass server, restart it and repeat the step: the
  standard pywin32 wrapper is used. If the problem goes away, report the step; `scripts\check_fastcom.py` compares
  both wrappers idiom by idiom in a hidden Excel of its own.

## The reply says the save after the last answer failed (since phase 3c)

- **Symptom:** a warning `the save of <file> after the last answer failed (<Excel's message>); it is saved now`, or
  `... could not be saved now either (...): save it yourself`, or `<file> was closed before smartass saved it after
  the last answer`.
- **Cause:** with the add-in, Excel saves the answer's own file right after the reply (`engine\smartass\xl\inproc\
  later.py`), and writes the outcome to `%LOCALAPPDATA%\smartass\inproc\saved\<Excel's pid>.json`. The next answer
  in that Excel reads it; on a failure it says so and saves in the job. `fin_status` shows the last outcome under
  `save_after_reply`.
- **Check:** if it repeats, set `SV_SAVE=sync` for the smartass server and restart it: every answer then saves its
  file before it replies, as before (0.4-1.1 s slower). If Excel is closed before the save ran, Excel asks whether to
  save the workbook (its AutoSave is still off at that moment).

## Still stuck

Run `smartass doctor --bundle` and send the zip it prints to the maintainer. It contains the doctor result,
versions, your `config.toml` and the tail of the logs. Tokens and passwords are redacted, but it does contain
file paths (including your user name).
