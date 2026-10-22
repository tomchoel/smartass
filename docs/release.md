# Release checklist

A release is an unpacked snapshot `<release folder>\<version>\` plus `<release folder>\latest.json` and (optionally)
`<release folder>\atlas\atlas.db`. Users get it through the update check at login (`smartass serve`) or
`smartass update`. `latest.json` is written last, so clients never see a half-copied release.

Users point at the release folder with `release.folder` in `config.toml` (or `SMARTASS_RELEASE`). Environment
variables and `~` are allowed, for example `folder = '%OneDriveCommercial%\<synced library>\smartass-release'`. An empty
value means "not configured".

1. Bump `version` in `engine\pyproject.toml` (the only place; the engine reads its version from the package metadata).
2. Add an entry to `CHANGELOG.md`.
3. Run `uv lock` and the tests: `cd engine; uv lock; uv run pytest -q; uv run ruff check .`
4. Run the live checks in `engine\tests\live\README.md` on this PC.
5. Commit and push; CI must be green.
6. Publish: `scripts\publish_release.ps1 -Folder "<release folder>" [-WithAtlas]`.
   It refuses a dirty working tree, an existing version, or an engine that reports a different version than
   `pyproject.toml`, copies HEAD (`git archive`), optionally snapshots your
   `%LOCALAPPDATA%\smartass\atlas.db`, tags and pushes `v<version>`, and only then writes `latest.json`. Any git
   failure stops it before `latest.json`.
7. Monthly: after `sv_atlas_build`, bump the patch version and publish again with `-WithAtlas`.
8. Roll back: a user pins a version with `release.version` in their `config.toml` (pinning down is allowed), or
   you republish the older code under a new, higher version number.
