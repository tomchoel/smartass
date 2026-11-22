# Light and fast, phase 4: slim down - the pane, the HTTP API and the updater are gone

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan
> task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** smartass is one thing: an MCP server (grids, lines and the member catalogue over HTTPS; answers in the
user's Excel) with a small command line (`doctor`, `setup`, `login`, `atlas`, `knowledge`). The Excel task pane and
its server, the localhost HTTP API (`smartass serve`) with its secret and Startup entry, the pane's `smartass xl`
sub-commands, and the updater with its release-folder settings are deleted. Install is `uv sync`, `smartass setup`,
`smartass login`.

**Spec:** `docs/specs/2026-11-19-light-fast-design.md`, section "3. Slimmer product" and the "Order of work" row for
phase 4. **Proof:** `smartass doctor` clean on a fresh checkout with a fresh home (apart from "not signed in" /
"no atlas", which say which command to run); the whole suite and the offline test set green.

**What exists (0.7.0):** `addin/tn-claude/` (a web task pane, `server.mjs`, start scripts, `manifest.xml`),
`engine/smartass/ops/httpapi.py` and `cli.py`'s `serve`, `config.HTTP_SECRET` / `SECRET_FILE` / `load_http_secret`,
`setup_user.secret_step` and the `smartass-engine.lnk` Startup entry, `cli.py`'s `xl` sub-commands (changes, revert,
check), `ops/update.py` and `cli.py`'s `update`, `config.RELEASE_FOLDER`, `setup_user.copy_atlas`, `doctor`'s
`http_api` and `update` checks, `installer/install.ps1`, `docs/release.md`. The MCP relay (`engine/run_server.py`,
`child_server.py`) is NOT part of this: it stays.

## Global constraints

- No tool name, tool parameter or reply field of the MCP server changes. The `xl_*` MCP tools stay (only the
  pane's CLI twins go).
- Nothing on the path of a tool call changes behaviour: the suite's existing tests of `fin_*` / `xl_*` pass
  unchanged; the 0.5.0 job-runner tests and Task 5's characterisation tests are not touched.
- Standing rules: never close, kill or save the user's own workbooks or Excel; read-only towards EPM; no token,
  session key, cookie or URL query in a log, error, reply or fixture; tests reach no EPM, no network, no Excel
  (`tests/conftest.py` enforces it).
- An upgrading PC is handled, never surprised: what an older release left (a Startup entry, a running `serve`, the
  secret file, the pane's registration in Excel, release-folder settings in a user's config) is found by
  `smartass doctor` and removed only by `doctor --fix` / `setup`, each thing said first; a process is stopped only
  when its command line proves it is this install's. Nothing in the user's Excel is changed without the owner's
  word (controller only).
- No EPM, no Excel, no network for builders. Line length 120; ruff and the suite green before every commit;
  commits end with the session's `Co-Authored-By` trailer.

---

### Task 1: The pane and the HTTP API are deleted

**Delete:** `addin/tn-claude/` (and `addin/` if nothing else is in it), `ops/httpapi.py`, `cli.py`'s `serve` and `xl`
sub-commands with the code only they used (find it: `ops/`, `xl/workbench` entry points used ONLY by the CLI - the
MCP tools' code stays), `config.HTTP_SECRET` / `SECRET_FILE` / `read_secret` / `load_http_secret` and the
`logsetup.add_secret` calls for it, `setup_user.secret_step` and the making of `smartass-engine.lnk`, `doctor`'s
`http_api` check, the sentence in `oauth.py` about `serve` keeping the sign-in fresh (the MCP server does), the
tests and docs of all of it, the pane's open item in DESIGN / README / troubleshooting / the guard list
(`server.mjs`'s tree kill is gone with it).
- **Leftovers of an older release (tests for each):** `setup` removes `smartass-engine.lnk` and stops a running
  `serve` whose command line names THIS install's venv and `smartass.ops.cli serve` (another install's is left and
  said); `doctor` lists `secret.bin`, a plain `http.secret` still in the user's `config.toml` and the pane's
  registration - find out from `installer/install.ps1`, `addin/tn-claude/README.md` and the start scripts HOW the
  pane was registered with Excel (a sideloaded manifest: a registry value, a shared-folder catalogue, a file in
  the user's profile) - and `doctor --fix` removes exactly those, each said first. Registry or Excel-side removal
  is written but NOT run by the builder.
- [ ] Failing tests first for the leftovers; then the deletion; grep every deleted name through engine, scripts,
  skills, installer, docs, the MCP instructions; a guard test (as Task 6 of phase 2 added for the hidden Excel):
  the engine opens no listening socket and holds no HTTP server. Suite, ruff. Commits by step.

### Task 2: The updater and the release folder are deleted; install as the design says

**Delete:** `ops/update.py`, `cli.py`'s `update`, `config.RELEASE_FOLDER` and the `release.*` settings,
`setup_user.copy_atlas`, `doctor`'s `update` check, `docs/release.md` (or rewrite it to what a release is now: a
git tag and the CHANGELOG), the `-WithAtlas` publishing, their tests.
- `smartass login` builds the atlas when there is none (about half a minute, plus the cards; it says so and can be
  skipped with a flag); a tool call without an atlas already says which command to run - pin it with a test for
  every `fin_*` entry that needs the atlas.
- `installer/install.ps1`: what is left is `uv sync`, `smartass setup`, `smartass login` - rewrite it to that or
  delete it and put the three lines in the README (choose the smaller; say which). Read it, do not run it.
- A user's `config.toml` that still has `release.folder` or `http.secret` is read without error and `doctor` says
  the setting is no longer used.
- [ ] Tests first; the deletion; grep; suite, ruff.

### Task 3: Doctor and setup as the design lists them

- `doctor`'s checks: config, excel, smartview_startup (kept by the ruling of 22.11), EPM reachable, oauth
  sign-ins, atlas, knowledge, output folder, leftovers. Fix the **leftovers false alarm** found on the owner's
  PC on 22.11: a `smartview_options_original.json` in the OLD svkit home is a record an older engine always kept
  there, not proof of a crash; doctor warns only when the file is in THIS home (a crashed run), and for the
  legacy home says nothing unless Smart View's stored options
  (`HKCU\Software\Hyperion Solutions\HyperionSmartView\Options` `CAOptionsXML`: `messageDisplay`,
  `bDisplayHomePanelOnStartup`, `bCommentDlgEnable`) differ from the recorded originals - read without COM; when
  they cannot be read, no warning. Tests with fake registry readers.
- `setup`: user config, MCP entries for Claude Code and Copilot, the skill link, the leftovers of Tasks 1-2. It
  writes no Startup entry any more.
- **The proof:** a test (and a script the controller runs) that builds a fresh home in a temp folder
  (`SMARTASS_HOME`), runs `setup`'s steps with fakes for the registry / MCP files, and `doctor`: every check OK or
  a "run this command" message for the sign-in and the atlas - nothing else.
- [ ] Tests first; suite, ruff.

### Task 4: Docs, version 0.8.0

`docs/DESIGN.md`, `docs/troubleshooting.md`, `README.md` (install in three commands), the skill's `reference/`, the
MCP instructions (budgets), `CHANGELOG.md` 0.8.0 (what is gone, what an upgrading PC should run: `smartass setup`,
`smartass doctor --fix`; what a new PC needs), the plan's "As built". Version 0.8.0 (`pyproject`, the lock's
version line only, `VERSION`). A docs guard as in 0.7.0: the docs name no deleted part as a current thing.

## Controller steps (with the owner's word)

On the owner's PC after the merge: `smartass setup`'s leftover steps only (the Startup shortcut), `doctor` then
`doctor --fix` (the secret file, the pane's registration outside Excel); anything that changes what Excel loads
(the pane's manifest in Excel) only on his explicit answer.
