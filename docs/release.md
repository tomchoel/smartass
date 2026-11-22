# Release checklist

A release is an unpacked snapshot `<release folder>\<version>\` plus `<release folder>\latest.json` and (optionally)
`<release folder>\atlas\atlas.db`. Users get it with
`smartass update`. `latest.json` is written last, so clients never see a half-copied release.

Users point at the release folder with `release.folder` in `config.toml` (or `SMARTASS_RELEASE`). Environment
variables and `~` are allowed, for example `folder = '%OneDriveCommercial%\<synced library>\smartass-release'`. An empty
value means "not configured".

1. Bump the version with `scripts\bump.ps1` (`-Part patch|minor|major` or `-Set x.y.z`): it writes `engine\VERSION`
   (what the engine reports, and what makes running servers reload) and the same version in `engine\pyproject.toml`.
2. Add an entry to `CHANGELOG.md`.
3. Run `uv lock` and the tests: `cd engine; uv lock; uv run pytest -q; uv run ruff check .`
4. Run the test set against EPM on this PC (`uv run python -m smartass.testset run`, golden numbers outside git) and,
   when Excel code changed, `scripts\check_inproc.py` and `scripts\check_insert.py` (each starts an Excel of its
   own).
5. Commit and push; CI must be green.
6. Publish: `scripts\publish_release.ps1 -Folder "<release folder>" [-WithAtlas]`.
   It refuses a dirty working tree, an existing version, or an engine that reports a different version than
   `pyproject.toml`, copies HEAD (`git archive`), optionally snapshots your
   `%LOCALAPPDATA%\smartass\atlas.db`, tags and pushes `v<version>`, and only then writes `latest.json`. Any git
   failure stops it before `latest.json`. Publish `-WithAtlas` only from an atlas that 0.7.0 or later built
   (`smartass doctor` has no `atlas` warning about EPM's real member names): an older one turns the strict member
   check over Planning REST off on every PC that copies it (`smartass setup` keeps a local atlas that has them).
7. Monthly: after `smartass atlas build` (0.7.0 or later), bump the patch version (`scripts\bump.ps1`) and publish
   again with `-WithAtlas`.
8. Roll back: a user pins a version with `release.version` in their `config.toml` (pinning down is allowed), or
   you republish the older code under a new, higher version number.
