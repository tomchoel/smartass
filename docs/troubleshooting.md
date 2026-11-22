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

## smartview_startup

smartass uses no Smart View (since 0.7.0), but Smart View stays installed in your Excel and can still open its own
Login there. While that Login is open, Excel does not answer smartass (see "Smart View's Login is open in your Excel"
below).

- **Symptom:** `WARN  smartview_startup  Smart View checks for extension updates at every Excel start (its Login opens
  ~20 s after Excel opens, and smartass does not touch that Excel until it is signed in or closed)`.
- **Cause:** Smart View's extension update check at Office start (`HKCU\Software\Oracle\SmartView\extensions`
  `CheckOnStartup=1`) needs an EPM session, so it opens the login (found 2026-11-03 in Smart View's own traffic log).
- **Fix:** close Excel; Smart View > Options > Extensions: untick "Check for updates when Microsoft Office starts".
  Smart View then signs in only when you use it; extension alerts on connect and the manual update check still work.
- **Symptom:** `WARN  smartview_startup  Smart View diagnostics capture is on ...`
- **Fix:** Smart View > Options > Advanced: turn diagnostics off (`Profile=0`), then delete
  `%APPDATA%\Oracle\SmartView\Diagnostics` (it holds sign-in cookies in plain text).

## epm

- **Symptom:** `FAIL  epm  https://... unreachable: <error>`.
- **Cause:** no network route to Oracle EPM Cloud (off the Telenor network, VPN down, or a proxy problem).
  An OK here only means the server answers; the login itself is checked on the first query.
- **Fix:** Connect to the Telenor network (or VPN) and check the proxy.

## oauth

Grids and drill lines go over plain HTTPS with your own OAuth access token; smartass starts no browser (since 0.6.0
there is no hidden browser at all). Without the set-up below every tool call that reads EPM answers with the set-up
text instead.

Measured 19.11.2026 with scopes for plan1 (FinMax, FinLite) and stratmodel (FinPlan): from a cold start a FinMax
grid took 0.7 s, a FinPlan grid 0.7 s and the closed month 0.1 s; through the old browser's sign-in the first grid
took about 20 s. Every grid compared was the browser route's cell for cell. Measured 21.11.2026 on four cells: the
drill lines over the token are the browser route's row for row, with exact ties; a 2 839-line drill took 8.4 s
(26.1 s through the browser, its start and sign-in included). Details:
`docs\audit\2026-11-19-drill-token-vs-browser.md`.

- **Set-up (per user, `config.toml`):** `[oauth]` `idcs` (the identity domain's base URL, https) and `client_id` (a public
  client an administrator registered there, with the device-code and refresh-token grants), and one scope per pod
  under `[oauth.scopes]`: `"https://<pod host>" = "urn:opc:serviceInstanceID=<id>urn:opc:resource:consumer::all"`.
  A token covers one environment. The administrator adds an environment to the client as a resource and reads its
  scope there.
- **Symptom:** `FAIL  oauth  not set up: config.toml has no [oauth] idcs, client_id and scopes`; a tool reply says
  "EPM access is not set up on this PC".
- **Fix:** add the three settings above to config.toml, then run `smartass login`.
- **Symptom:** `FAIL  oauth  <pod>: no scope (epm.base: no answers)`; every tool reply that reads EPM says "no OAuth
  scope is configured for https://<pod>".
- **Cause:** the pod of `epm.base` has no scope under `[oauth.scopes]`. Every answer reads the closed month and the
  list of data sources from that pod first, and those have no other route: no answer works, no drill either.
  Another pod without a scope loses only what is asked of that pod: its grids cannot be fetched (Planning REST needs
  the same scope), and the reply's warnings say so.
- **Fix:** Add the pod's scope under [oauth.scopes] in config.toml (the identity domain's administrator reads it
  where that environment was added to the OAuth client as a resource), then run `smartass login`.
- **Symptom:** `WARN  oauth  <pod>: not signed in` or `<pod>: sign-in ended`; a tool reply gives a page and a code.
- **Cause:** no sign-in was made on this PC, or it was not used for 7 days (Oracle's lifetime for a refresh token;
  the MCP server renews it daily while it runs), or the identity domain withdrew it.
- **Fix:** Run `smartass login`: a page opens in your browser and the code is on the clipboard; paste it and
  approve with Telenor SSO. Or approve the code the tool reply shows and call the tool again.
  `smartass login --status` shows the sign-ins, `--forget` deletes them from this PC.
- The sign-in is stored per pod in `%LOCALAPPDATA%\smartass\oauth\`, DPAPI-protected for your Windows user; smartass
  never sees a password. The token is only ever sent with read-only requests: the Smart View provider's, the
  drill's own (its query, view and getData) and Planning REST's reads (a grid's export, a plan type's dimensions,
  a dimension's members).
- **Symptom:** a reply says EPM refused the access token twice (HTTP 401/403).
- **Fix:** Run `smartass login --force`: plain `smartass login` skips a pod that still looks signed in. If it
  repeats, ask the EPM administrator to check your role and the OAuth client's scope for that pod.
- **Left over from before 0.6.0, safe to delete by hand:** the old hidden browser's profile folders
  `%LOCALAPPDATA%\smartass\drill_edge*` and failed drill pages `logs\drill_fail_*`. smartass no longer uses them and
  never deletes them itself. A `[drill] channels` line in config.toml is ignored.

## How a grid is fetched (since 0.7.0: the Excel fallback is gone)

Every grid goes first over the Smart View provider route (HTTP, your token). A grid that route cannot fetch is
fetched once over Oracle's Planning REST `exportdataslice` with the same token, the second engine; no Excel is ever
asked for a grid. The reply's `routes` counts the grids per engine (`http`, `rest`). What a warning means:

| Warning or error in the reply | Meaning | What to do |
|---|---|---|
| "N grid(s) were fetched over Planning REST, the second engine: the Smart View provider route could not fetch them (...)" | The numbers are complete and the same as the provider route's (live gate: docs\audit\2026-11-21-rest-vs-provider.md). Over REST a cell you have no access to cannot be told apart from an empty one: it reads as missing, not #NoAccess. | Nothing, unless you expected #NoAccess cells. |
| "FinPlan: EPM refused Entity member X over Planning REST ... so it was left out: its figures are missing from this answer and every total that includes it (metrics) is INCOMPLETE - check the member or rebuild the atlas (`smartass atlas build`)" | EPM no longer has that member there (renamed or removed since the atlas or the question), or you may not read it. The rest of the grid is fetched without it (at most 3 members are left out of one request), so a total that includes it shows only the others. The same text is under Caveats on the Answer, and in Known limits on the Sources sheet after the question and the day ("Q3 (2026-11-22): ..."), until that question is answered again without it. | Check the member with fin_lookup; rebuild the atlas; ask again with today's member. |
| "... could not fetch it either (query q): Planning REST refused the ... grid: ... member X ... cannot be left out - check the member or rebuild the atlas" | The refused member is a POV member or the only member on its axis (often the question's one entity): leaving it out would change what every number means, so that grid is not fetched. The answer is still drawn, without that grid's numbers, and this warning names the member. The same when more than 3 members of one request are refused ("Planning REST refused more than 3 members of the ... grid"). | As above: correct the member (fin_lookup) and ask again. |
| "N grid(s) could not be fetched (the Smart View provider route, then Planning REST): ..." | Neither engine could fetch them; the reasons follow. Those numbers are missing from the answer. | Ask again in a minute; if it repeats, `smartass doctor`. |
| A drill: "the drill link of ... could not be fetched over the Smart View provider route (...); Planning REST has no drill-report request, so there is no other route. Next step: ..." | Only the provider route can give a cell's drill link. | The reply's next step, by cause: the scope to add, `smartass atlas build`, `smartass login --force` (a token EPM refused), `smartass login --status` (no access token), the members by their codes (a grid EPM refused for an alias clash), else try again in a minute. |
| "... EPM refused the access token twice (HTTP 401/403)" | EPM's gate refused a token this PC still holds as valid, also after one new token. Usually a stale sign-in; it can also be a missing EPM role or a missing OAuth scope for that pod. | Run `smartass login --force` (plain `smartass login` skips a pod that still looks signed in), then ask again. If it repeats, ask the EPM administrator to check your role and the OAuth client's scope for that pod. |

## atlas

- **Symptom:** `FAIL  atlas  missing: ...\atlas.db` or `WARN  atlas  ...\atlas.db (N days old)` (older than 35 days).
- **Cause:** the member atlas (used to find members by name) was never copied, or is out of date.
- **Fix:** FAIL: Re-run installer\install.ps1 (copies the monthly atlas), or run `smartass atlas build`.
  WARN: Re-run the installer after the monthly release, or run `smartass atlas build`.

`smartass atlas build` (since 0.7.0) reads every dimension over Planning REST with your sign-in: read-only, no Excel,
the member trees of both pods in about 25 s (docs\audit\2026-11-22-atlas-rest.md); a full build then derives the
behaviour cards from data, as before: about two minutes more. It builds a new file and replaces the atlas only when
every dimension was built. `smartass atlas status` shows the last attempt as `last_build`.

- **Symptom:** the build's last line is `atlas: NOT replaced (<why>); ...` and it exits 1. `smartass knowledge
  refresh` prints "atlas not replaced: <why> - this refresh used the atlas as it was", and its report file says
  "Atlas: NOT replaced (<why>); the refresh used the atlas as it was".
- **Cause and fix, by the reason given:**
  - "not signed in to EPM - run `smartass login`": run it, then build again.
  - "1 dimension or cube failed" (or more): the line above it names each failed dimension and EPM's reason. Build
    again later; if it repeats, `smartass doctor`.
  - "another atlas build is running": wait for it to end (another terminal, or a knowledge refresh).
  - "nothing was built (no such dimension?)": a `--dimension` name that no cube has; check the spelling.
  - "atlas.db could not be read / written (... in use by another process ...)": a server held the atlas for over
    30 s; build again.

  Your atlas is unchanged in every case.
- **Symptom:** the build's last line contains `; FinMax: dimension X is new in EPM; this release cannot read it -
  update smartass` (also in `atlas status`).
- **Cause:** EPM has a dimension this release has no read rule for (each dimension is allowed by name). It is
  skipped and its old rows are kept (none, when it is new); the rest of the atlas is replaced.
- **Fix:** Update smartass (`git pull`, then `uv sync --frozen` in `engine`) when a release that knows it is out;
  nothing else is needed meanwhile.
- **An older smartass on the same atlas:** an installed release before 0.7.0 can still rebuild this atlas file. This
  release then ignores the real member names of the dimensions it rebuilt until its own next `atlas build`.
- **Symptom:** `WARN  atlas  ...\atlas.db (N days old); no dimension has EPM's real member names` (or "K of N
  dimensions have no ...").
- **Cause:** the atlas was built before 0.7.0 (or an older release rebuilt those dimensions since). It works, but
  without the strict member check over Planning REST: when a grid comes over the second route, smartass cannot check
  that EPM answered for the member it asked for, in the rare case where one member's key is another member's name.
- **Fix:** Rebuild once with `smartass atlas build`. `smartass setup` never replaces an atlas that has the real
  names with a release copy that has none; it says "atlas: local copy kept" and why.

## output

- **Symptom:** `FAIL  output  cannot write <folder>: <error>`.
- **Cause:** the folder for answer workbooks (`paths.analyser`) does not exist and cannot be created, or is
  read-only (for example OneDrive not signed in).
- **Fix:** Set paths.analyser in config.toml to a folder you can write.

## leftovers

- **Symptom:** `WARN  leftovers  an older smartass crashed with Smart View's options changed; the originals recorded in
  ...\smartview_options_original.json: ...` and/or `N VBA bridge add-in file(s) of an older release in ... (unused)`.
- **Cause:** a release before 0.7.0 ran a hidden Excel of its own with Smart View, changed three of Smart View's global
  options while it ran, and put them back when it stopped. The file is left when such a run crashed: nothing can put
  those options back now. The `svbridge_*.xlam` files are add-ins it built; nothing uses them now.
- **Fix:** Set the recorded values back by hand in Excel: Smart View ribbon > Options (close the other Excels
  first). Then run `smartass doctor --fix`: it names each file and removes it. A bridge file still loaded in an Excel
  stays (it says so): close that Excel and run it again. A copy in the old svkit folder is reported, never changed.
- At server start smartass also ends an Excel such a release started hidden and left running, but only one its
  `pids.json` proves was its own; then that file goes. `smartass setup` removes that release's sign-in watcher from
  the Startup folder and stops it when it runs from this install.

**The HTTP API and the Excel task pane (up to 0.7.0)**

- **Symptom:** `WARN  leftovers` naming `smartass-engine.lnk`, a running `smartass serve`, `secret.bin`, a plain
  `[http] secret` in config.toml, the task pane's Startup entry, its add-in catalogue (`...\TNClaudeCatalog`) or its
  classic add-in (`TNClaude64.xll`, loaded by an `OPEN<n>` value).
- **Cause:** a release up to 0.7.0 ran a localhost HTTP API (`smartass serve`, started at logon by the Startup entry
  `smartass-engine.lnk`, with its secret in `secret.bin`) for an Excel task pane. Both are gone: the agent starts the
  MCP server itself, and nothing runs at logon.
- **Fix:** run `smartass setup` or `smartass doctor --fix`. They remove the Startup entry and stop that `serve` when
  its command line names this install's `.venv` (another install's is left running, and said); `doctor --fix` also
  removes `secret.bin`, the plain secret line (the value is never shown) and a Startup entry of the pane, each named
  first.
- The pane's registration lives in Office's own settings, which smartass never changes. Remove it in Excel:
  - **The add-in catalogue:** File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs: select
    the catalogue ending in `TNClaudeCatalog`, Remove, tick "Next time Office starts, clear all previously-started
    web add-ins cache", OK, and restart Excel. The folder on the file share and its `manifest.xml` can then be
    deleted by hand.
  - **The classic add-in:** File > Options > Add-ins > Manage: Excel Add-ins > Go: untick `TNClaude` (or the entry
    naming `TNClaude64.xll`), OK, and restart Excel. Use the dialog: do not delete the `OPEN<n>` value in the
    registry by hand, Excel renumbers those values and rewrites them when it closes. Then delete the add-in's folder
    (`%LOCALAPPDATA%\svkit\addin`) by hand; smartass never changes the old svkit folder.
  - If the pane's local server still runs (`node server.mjs` on port 3000), end it in Task Manager, or log off.

**The updater (up to 0.7.0)**

- **Symptom:** `WARN  leftovers` naming `release.folder`, `release.version` or `http.port` ("no longer used") in
  config.toml, or the updater's record `update_attempt.json` in the home.
- **Cause:** a release up to 0.7.0 updated itself from a shared release folder (`smartass update`) and copied the
  monthly atlas from there. Both are gone: smartass is updated from git, and `smartass login` builds the atlas over
  EPM when there is none (`smartass atlas build` rebuilds it at any time).
- **Fix:** run `smartass doctor --fix`: it names the settings (never their values) and takes the `[release]` and
  `[http]` sections out of config.toml, and removes the record. A setting written another way (a dotted key, an
  inline table) is said to be left: delete it by hand.

## Any check: `FAIL  <check>  <ExceptionType>: <message>`

The check itself crashed. Fix: Send the file from `smartass doctor --bundle` to the maintainer.

## Smart View's Login is open in your Excel (since 0.7.0)

- **Symptom:** a reply says "Smart View's Login is open in the Excel showing '<window title>' (pid N); smartass does
  not touch an Excel while it is ...", with the next step.
- **Cause:** Smart View (installed in your Excel; smartass does not use it) opened its EPM Login there, often at Excel's
  start (see `smartview_startup`). While that window is open, Excel does not answer calls from outside, so smartass
  refuses at once instead of waiting. It no longer clicks "Telenor SSO" for you; that automatic click is gone, also
  when you use Smart View by hand.
- **Fix:** sign in in that Login, or close it, then ask again. If it shows at every Excel start, untick the start-up
  check (`smartview_startup`).

## Known limits (accepted by the owner, 2026-10-23)

These were investigated in the reliability audit (`docs\archive\audit\2026-10-23-reliability-audit.md`) and accepted as they
are. If one of these symptoms shows up, the cause below is already known; no new investigation is needed.

| Symptom | Known cause | What to do |
|---|---|---|
| A build in a newly started Excel window is slow in `open target.new workbook` | A cold Excel start for a new interactive Excel (a parallel agent, or no Excel open), and the 12 s from when the workbook shows that smartass watches a new Excel for Smart View's Login before the answer's first step there (until 25 s after Excel's start where Smart View's start-up check is on). | None; the next build in that Excel is fast. |
| A new analysis workbook has no TN cell styles (`TN Heltall`, ...) in its style gallery | Since phase 3c the first answer's built file is the workbook; the styles are installed when something needs them. | Run `xl_tidy` on a sheet (it merges them from the TN template), or ignore: the answer's own cells carry their formats. |
| An empty (0-byte) analysis file, or one with an answer's sheets but no formulas, tables, Sources rows or stored spec, in the analyser folder | The server was killed (or Excel died) while the first answer of a new workbook was made: the name is reserved as an empty file, then the built file is written and Excel opens it; only the job's clean-up deletes it, and the session points at it only after the job's save. | Close it in Excel without saving if it is open, and delete it; ask the question again. |
| A new analysis file that a failed answer (or a probe) deleted is back in the analyser folder minutes later, online-only | OneDrive restores a file Excel had just saved to the cloud when the local delete races that upload (seen with probe files Excel saved and closed, 20.11.2026). | Delete it again. |
| Progress notes (`still working: …`) keep arriving after a call to your Excel timed out | The abandoned job's COM call is still hung; its ticker runs until that call returns. The client drops notes for a finished request. | Harmless. It stops when Excel answers or closes. |

## Open items since 0.7.0 (known, not decided yet)

| Symptom | Known cause | What to do |
|---|---|---|
| An answer in an Excel smartass has just started waits 10 minutes and fails, with Smart View's Login open there | That Login opened later than the watch (25 s after Excel's start, on a slow PC), after the answer's first step in that Excel. | Sign in in the Login or close it, then ask again. |
| An Excel smartass started closes without saving when something ends the smartass server's processes from outside | An Excel smartass starts (no Excel open, or a parallel agent) is a child of the smartass server process; ending the server's whole process tree ends that Excel with it. smartass's own stop of a server that does not exit never ends an Excel. | Save your work in such an Excel before you end the server. An open item. |
| Every first workbook made with no Excel open has default column widths and fonts, not the TN template's, and none is kept for the next time | The stored look of a new workbook has a column width that is not a standard one, so it is not used; and as a look is stored, smartass does not make that workbook from the template, which would store a new one. | An open item: tell the maintainer. The answers in it are right. |

## An Excel step behaves differently since 2026-11-19 (the fast COM wrapper)

- **Symptom:** an Excel step fails or gives an odd result it did not give before 2026-11-19, e.g. an
  `AttributeError` that names a COM object, or a method where a value was expected.
- **Cause to rule out:** since 2026-11-19 the engine drives Excel through `engine\smartass\xl\fastcom.py`: pywin32's
  own late-bound wrapper, except that what it learns about each interface is shared by all objects of it (about
  x1.7 faster Excel jobs).
- **Check:** set the environment variable `SV_FASTCOM=0` for the smartass server, restart it and repeat the step: the
  standard pywin32 wrapper is used. If the problem goes away, report the step; `scripts\check_fastcom.py` compares
  both wrappers idiom by idiom in a script's own hidden Excel.

## The reply says the save after the last answer failed (since phase 3c)

- **Symptom:** a warning `the save of <file> after the last answer failed (<Excel's message>); it is saved now`, or
  `... could not be saved now either (...): save it yourself`, or `<file> was closed before smartass saved it after
  the last answer`.
- **Cause:** with the add-in, Excel saves the answer's own file right after the reply (`engine\smartass\xl\inproc\
  later.py`), and writes the outcome to `%LOCALAPPDATA%\smartass\inproc\saved\<Excel's pid>.json`. The next answer
  in that Excel reads it; on a failure it says so and saves in the job. `fin_status` shows the last outcome under
  `save_after_reply`. (The first answer of a new workbook is always saved in the job.)
- **Check:** if it repeats, set `SV_SAVE=sync` for the smartass server and restart it: every answer then saves its
  file before it replies, as before (0.4-1.1 s slower). If Excel is closed before the save ran, Excel asks whether to
  save the workbook (its AutoSave is still off at that moment).

## The reply says an earlier smartass job did not finish (since 0.5.0)

- **Symptom:** a warning `an earlier smartass job in this Excel did not finish and left Calculation manual, ...: put
  back now`, or `AutoSave could not be switched back on in <file> ...`.
- **Cause:** every answer's job records the Excel settings it changes (ScreenUpdating, EnableEvents, DisplayAlerts,
  Calculation, and AutoSave of the answer's own file; no figures) in the state file
  `%LOCALAPPDATA%\smartass\jobstate\<Excel's pid>.json` before it changes them, and removes it at its end. A job that
  died (a killed server) or whose put-back Excel refused leaves the file, and the next job in that same Excel process
  puts back each setting that still reads as the dead job left it - one you changed since is left alone - and says
  so. What Excel refuses again stays in the file for the job after; a file of an Excel process that is gone is
  removed. The file is never used in another Excel process that got the same pid.
- **Check:** nothing to do. If the same warning comes with every answer, Excel keeps refusing a setting: restart that
  Excel. Deleting the folder is safe when no answer is running.

## Still stuck

Run `smartass doctor --bundle` and send the zip it prints to the maintainer. It contains the doctor result,
versions, your `config.toml` and the tail of the logs. Tokens and passwords are redacted, but it does contain
file paths (including your user name).
