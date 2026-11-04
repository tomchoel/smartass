# smartass v2 · Plan 1G: lift, harden, retire v1

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** v2 stands alone.
- **Lift.** The v1 modules v2 needs move into `engine/smartass/` with real imports. The legacy monkeypatches
  become native settings.
- **Harden.** The runtime fixes go in.
- **Keep the CLI and API.** The CLI (`doctor`, `setup`, `update`) and `serve` (`/health`, `/analyze`) are ported.
- **Retire v1.** `engine/smartview_mcp`, its tools, tests and skill text are deleted. One MCP entry, `smartass`,
  runs v2.
- **Gates:**
  - `uv run pytest -q`, `uv run ruff check .` and `python -m smartass.testset run` (8/8) are all green.
  - A live smoke check on Tom's go.

**Decisions (Tom, 2026-11-04):**
- Generic open-workbook work uses Excel MCP plus `xl_snapshot` / `xl_revert` / `xl_tidy`. v1's other `xl_*` tools
  are not lifted.
- Everything outside the repo that points at v1 is updated: the global CLAUDE.md, the skills `smartass`,
  `tn-formelstil`, `excel-mcp` and `produktregnskapet`, and memory. Tom sees the CLAUDE.md and skill diffs before
  they are saved.
- One MCP entry named `smartass` runs v2. The `smartass2` entry is removed.
- `serve` and the CLI are kept and ported. `serve` exposes `/health` and `/analyze` (a v2 spec).
- Answer files go back to the plain analyser folder (no `v2` subfolder).

**Conventions:**
- Branch `v2/1g` from `main`. Git Bash from `/c/dev/smartass/engine`.
- Use `git mv` for moved files, so history follows.
- Never run python through a stdin heredoc. Bare python only with `< /dev/null`.
- Stage files explicitly.
- No live runs before Task 9.
- Commit after each task. Messages end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- After every task, the full `uv run pytest -q` and `uv run ruff check .` must be green before committing. A task
  that cannot get there stops and reports.

## Target layout (v1 → v2)

| v1 (`smartview_mcp/…`) | v2 (`smartass/…`) |
|---|---|
| `config.py`, `logsetup.py`, `timing.py`, `session.py` | `core/config.py`, `core/logsetup.py`, `core/timing.py`, `core/session.py` |
| `server.py`: `lifespan`, `guard`, `_err`, `with_next_step`, `explain`, `READ`/`LOCAL`/`WRITE`/`ADMIN`, `_call_user`, `_with_progress`, `host()`, `_ctx`, `progress_secs`, `_is_our_analysis_file`, `_warm_up` | `core/serverkit.py` (only these; nothing that imports `tools_*`) |
| `svhttp.py`, `svxml.py`, `grid.py`, `ops.py` (trimmed to what the closure uses), `sso.py`, `drill_browser.py`, `drill_direct.py`, `drill_tabs.py`, `drill_lines.py`, `drill_split.py`, `drill_derive.py`, `core/drillstep.py`, `core/drilltime.py` | `epm/…` (same file names) |
| `atlas.py` | `knowledge/atlas.py` |
| `host.py`, `hostproc.py`, `dialogs.py`, `excel/instances.py`, `excel/login.py`, `excel/target.py`, `excel/jobs.py`, `excel/sheets.py`, `core/owned.py`, `sheetorder.py`, `drill_xl.py`, `styling.py`, `formulas.py`, `vendor/` | `xl/…` (same names; `vendor/` → `xl/vendor/`) |
| `excel/adhoc.py` | `xl/pull_adhoc.py` |
| `workbench/books.py`, `workbench/history.py`, `workbench/fmt.py` | `xl/workbench/…` |
| `cli.py`, `doctor.py`, `setup_user.py`, `update.py`, `httpapi.py` | `ops/…` |
| `core/recipes/lines.py`: `add_initiative` + its constants | `model/lines.py` |
| everything else (recipes, gate, contract, answer*, fin_*, analytics, specs, templates, catalog, concepts, factstore, extract, pnl_facts, tools_*, workbench/changes, atlas_history, drill_store, qlog, …) | deleted |

The package is `xl/`, not `excel/`, to avoid confusion with `render/excel.py`. `vba/`, `defaults.toml` and the
relay files stay in `engine/`.

**Path constants.** `config.py` computes `HERE` / `REPO_DIR` / `BAS_FILES` / defaults paths from `__file__`. Moving
it one level deeper changes `.parent` chains. Fix them, and add a test that `config.REPO_DIR / "knowledge"`,
`config.BAS_FILES` and `defaults.toml` exist.

---

### Task 0: Branch and baseline

- [ ] `cd /c/dev/smartass && git switch main && git switch -c v2/1g`
- [ ] Record the baseline: `uv run pytest -q | tail -1`, `uv run ruff check . | tail -1`, and
  `uv run python -m smartass.testset run < /dev/null` (8/8 ok). Put the results in the Task 0 commit message of
  an empty commit: `git commit --allow-empty -m "chore(v2): 1G baseline ..."`.

### Task 1: `core/serverkit.py`

Extract the helpers listed in the table from v1 `server.py` into `smartass/core/serverkit.py`. Keep their code
unchanged; for now they import the v1 modules they need. `smartass/server.py` then uses
`from .core import serverkit as kit` instead of `from smartview_mcp import server as v1` (`kit.guard`,
`kit.lifespan`, `kit._call_user`, `kit.READ` …).

- `tests/v2/test_v2_save_spec.py` patches `smartview_mcp.server`; repoint it.
- Write one test: importing `smartass.server` does not import `smartview_mcp.tools_core`. Check
  `sys.modules` after a fresh import in a subprocess run with `< /dev/null`.
- Green, then commit: `refactor(v2): server helpers in core/serverkit, no import of v1 tools`.

### Task 2: Move the modules (mechanical)

For each row of the table, in dependency order (core, then knowledge/atlas, then epm, then xl, then ops):
1. `git mv` the file to its new path, and create the `__init__.py` files.
2. Rewrite imports in the moved files: relative imports inside v1 (`from . import config`,
   `from ..excel import jobs`) become the new relative paths.
3. Rewrite imports in `smartass/**` and `tests/v2/**` from `smartview_mcp.X` to the new path. Include the
   monkeypatch strings in tests, such as `"smartview_mcp.excel.jobs.com"`.
4. Leave nothing behind. There are no shims in `smartview_mcp`. Any v1 module still importing a moved module is in
   the "deleted" row, or it is a v1 test that Task 4 deletes or migrates.
5. While moving, delete the parts of `ops.py` that the closure does not use (keep the names the drill and ad hoc
   paths use). Then copy `add_initiative` and its constants into `model/lines.py`, and repoint `server.py`.

Do it in small commits (one group per commit). Run the v2 suite after each:
`uv run pytest tests/v2 -q`, plus the v1 tests of the moved modules once Task 4 migrates them. The full suite may
be red between commits of this task only. It must be green at the end of Task 4.

- **Path constants.** Fix them in `core/config.py` and add the test from the header.
- **ruff config.** `extend-exclude` changes from `smartview_mcp/vendor` to `smartass/xl/vendor`. Update
  `.gitattributes` `-text` rules for `vendor/*.py` to the new path.

Commit messages: `refactor(v2): lift <group> from v1`.

### Task 3: Native settings, legacy retired

- `core/session.py`:
  - Make v2's session behaviour native: the state file `sessions.json` (v2 owns it now), relay detection on
    `run_server.py`, and the `_last:<cwd>` fallback and v2-style pruning that `legacy.py` patched in.
  - Make every state-file write atomic: write `*.tmp`, then `os.replace`.
- `core/config.py`: `ANALYSER_DIR` is the plain `paths.analyser` (no `v2` subfolder).
- `xl/sheetorder.py`: `Data` is in `_OURS` / `_RANK` natively.
- Delete `smartass/legacy.py` and the `configure_v2()` calls. Port its tests to the native behaviour:
  - session reconnect
  - sheet order of Data
  - v2 sessions surviving each other
- Commit: `refactor(v2): session, analyser folder and sheet order native; legacy seam removed`.

### Task 4: Delete v1, migrate the tests that matter

- `git rm -r engine/smartview_mcp`.
- **Delete** the v1 tests whose subject is deleted:
  - recipes, gate, contract, answer, fin_*, analytics, specs, templates, catalog, concepts, factstore, extract;
  - every `tools_*` test;
  - `test_skill_docs`, `test_smoke_stdio` (replaced in Task 7);
  - `tests/live/*`, whose bench is replaced by `smartass.testset`.
- **Migrate** (repoint imports, keep assertions) the v1 tests of lifted modules:
  - `test_svhttp`, `test_svxml`, `test_grid`, `test_timing`, `test_logsetup`, `test_owned`, `test_sheetorder`,
    `test_config`, `test_global_options`;
  - `test_excel_*`, `test_host_dialogs`, `test_dialogs`, `test_drill*`, `test_vendor`, `test_formulas`,
    `test_ops_transient`;
  - `test_reload_fingerprint` (rewritten in Task 5), `test_setup_user`, `test_doctor`, `test_httpapi`,
    `test_update`.
- **Keep the helpers** `svxml_util`, `fakexl`, `stdio_util` and `conftest` where they are used, and delete the
  rest. `conftest` must still point `SMARTASS_HOME` at a temp folder.
- `pyproject.toml`:
  - wheel `packages = ["smartass"]`;
  - scripts `smartass = "smartass.ops.cli:main"`;
  - remove `smartview-mcp`.
- Run the full suite and ruff, both green.
- Commit: `refactor(v2): v1 package deleted; tests of lifted modules migrated`.

### Task 5: Runtime fixes (TDD each)

1. **Atomic registry.** `xl/hostproc` `PidRegistry.save` writes atomically under a lock file (`msvcrt` byte lock,
   like v1's other locks). Test: a concurrent save from two threads never leaves invalid JSON.
2. **Reload on a version file only.** The relay `run_server.py` reloads when `engine/VERSION` (a new file holding
   the package version) changes, or on an explicit `fin_status(restart=True)`. A `.py` edit no longer restarts
   the child. Test the fingerprint function. Document in the docstring how to roll a code change:
   `scripts/bump.ps1`, which bumps `VERSION`, or the restart tool.
3. **Secret with DPAPI.** `ops/setup_user` writes the HTTP API secret protected with DPAPI
   (`win32crypt.CryptProtectData`, already in pywin32) to `secret.bin`. `config` reads it via
   `CryptUnprotectData`. A plain-text `http.secret` in an existing `config.toml` is migrated once and removed from
   the file. Tests use a fake DPAPI.
4. **Smart View options restored at startup.** `xl/host` start: if `smartview_options_original.json` exists from a
   crashed run, the originals are restored first (then re-applied as now). Test with fakes.
5. **Doctor checks VBA project access.** `ops/doctor` checks
   `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM == 1` and names the fix: Trust Center → Trust
   access to the VBA project object model. Test with a fake registry reader.

Commit per fix: `fix(v2): <what>`.

### Task 6: CLI and `serve` ported

- `ops/httpapi.py`:
  - keeps `GET /health` (the doctor checks);
  - replaces `/gate` and `/answer` with `POST /analyze`, body `{"spec": <AnalysisSpec>, "target": "session"}`,
    which calls the same internal `_new_question` as `fin_analyze` and returns its reply;
  - keeps the secret check (now via `config`), the 1 MB limit, 409 while busy, and 400/401/413/503 as before.
- `ops/cli.py`: `doctor`, `setup`, `update` and `serve` as before, on the new modules.
- `installer/install.ps1`: `-m smartview_mcp.cli` becomes `-m smartass.ops.cli`.
- `scripts/publish_release.ps1`: reads `smartass.__version__`. Add `__version__` to `smartass/__init__.py`,
  synchronised with `engine/VERSION`.
- `scripts/probe_sv_http.py`, `gen_atlas_md.py` and `knowledge/build_accounts_map.py`: repoint the imports
  (`config`, `drill_browser`). Run each with `--help`, or import it, to prove it loads.
- Tests: adapt `test_httpapi` to `/analyze`, mocking `_new_question`.
- Commit: `feat(v2): CLI and serve on v2 (/health, /analyze)`.

### Task 7: One relay, one server, one skill

- `child_server.py` → `from smartass.server import main`.
- Delete `child_server_v2.py` and `run_server_v2.py`.
- `run_server.py` keeps its relay with the Task 5 version-file reload. Its log is `proxy.log`.
- `smartass/server.py`: `MCPServer("smartass", …)`.
- New stdio smoke test (replacing `test_smoke_stdio`): start `child_server.py` with
  `SV_DRILL_KEEP_WARM=0` and `SMARTASS_HOME=tmp`, send `initialize` and `tools/list`, and assert the 9 tool names:
  - `fin_analyze`, `fin_lookup`, `fin_status`, `fin_refresh`, `fin_drill`
  - `xl_snapshot`, `xl_revert`, `xl_tidy`
  - plus any `fin_status` restart parameter
  Reuse `stdio_util` if it fits.
- **Skill rewrite.**
  - `skills/smartass/SKILL.md`, under 2,000 tokens: the plan → build → critic loop, AnalysisSpec v2 essentials,
    the blocks, follow-up / refresh / drill, the Excel MCP + `xl_snapshot` / `xl_revert` / `xl_tidy` rule, quoting
    only `key_values`, finance format, and `/mcp` after a reload.
  - `reference/`: keep `cubes.md` (the dimension tops), and rewrite `operations.md` (errors, files, logs) and
    `drill.md` for v2. Delete `recipes.md`.
  - Add a test that the skill names exactly the registered tools.
- `docs/DESIGN.md`: rewritten as built (v2). The v1 design goes to `docs/archive/DESIGN-2026-11-v1.md`.
- `README.md`: the tools table and layout.
- Commit: `feat(v2): one relay, one server named smartass, skill and docs for v2`.

### Task 8: Final gates

- `uv run pytest -q` (green), `uv run ruff check .` (clean).
- `uv run python -m smartass.testset run < /dev/null`: 8/8 ok. The goldens are unchanged.
- `uv run python scripts/check_staged.py --all` (CI guard).
- CHANGELOG: under `## Unreleased`, add `### Removed` (v1, with the tool rename table `sv_answer` → `fin_analyze`,
  `sv_drill` → `fin_drill`, `sv_find_member` / `sv_members` / `sv_member_info` / `sv_dimensions` →
  `fin_lookup`, `sv_status` / `sv_restart` / `sv_atlas_*` → `fin_status`, `xl_*` → Excel MCP + `xl_snapshot` /
  `xl_revert` / `xl_tidy`) and `### Changed` (one `smartass` server, runtime fixes).
- Commit: `docs(v2): changelog for retiring v1`.

### Task 9: Outside the repo and live (controller, with Tom)

1. **MCP.** `claude mcp remove smartass -s user`, `claude mcp remove smartass2 -s user`, then:

   ```
   claude mcp add smartass -s user -e SV_DRILL_KEEP_WARM=0 -e SV_DRILL_PROFILE_SLOTS=4 -- "C:\dev\smartass\engine\.venv\Scripts\python.exe" "C:\dev\smartass\engine\run_server.py"
   ```

   `SV_DRILL_KEEP_WARM` is kept at 0 so the browser starts on first use.
2. **Global `~/.claude/CLAUDE.md` Excel section.** Open workbooks: Excel MCP (attach) for read, write and format;
   `mcp__smartass__xl_snapshot` before edits, `xl_revert` to undo, `xl_tidy` for TN style and the formula check.
   Show Tom the diff before saving.
3. **Skills outside the repo.**
   - `tn-formelstil` (the check is `xl_tidy`)
   - `excel-mcp` (the routing)
   - `produktregnskapet` (`sv_drill` → `fin_drill`)
   - `smartass` is a symlink to the repo skill, so it is already done.
   Show the diffs before saving.
4. **Memory.** Rewrite or delete the 14 memory files that describe v1 tools; keep the facts that are still true.
   Update `MEMORY.md`.
5. **Live smoke check (Tom's go)**, after `/mcp`:
   - `fin_status`;
   - `fin_analyze` with the B02 spec (3 496 062 985);
   - `fin_drill` ac.C60 / un.SY.2TEC / Sep 26 (gap 0);
   - `xl_snapshot` → Excel MCP edit → `xl_revert`;
   - `smartass doctor`;
   - `testset run` 8/8.
6. **Merge** `v2/1g` into `main` on Tom's OK.
