# Troubleshooting

Start with `smartass doctor`. It prints one line per check (`OK`, `WARN` or `FAIL`), and under every line that is
not OK a `fix:` line. The sections below follow the checks in the order the doctor runs them. The fix wording is
the same as in the doctor (`engine\smartview_mcp\doctor.py`).

User settings live in `%LOCALAPPDATA%\smartass\config.toml`; defaults in `engine\defaults.toml`.

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

## Still stuck

Run `smartass doctor --bundle` and send the zip it prints to the maintainer. It contains the doctor result,
versions, your `config.toml` and the tail of the logs. Tokens and passwords are redacted, but it does contain
file paths (including your user name).
