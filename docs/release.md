# Release checklist

A release is a commit on `main` tagged `v<version>`, with its entry in `CHANGELOG.md`. There is no release folder:
a PC runs a git checkout and is updated with `git pull`, then `uv sync --frozen` in `engine` (and `smartass setup`
when the release says so). Servers that are running reload when `engine\VERSION` changes. The atlas is not shipped:
`smartass login` builds it on a PC that has none, and `smartass atlas build` rebuilds it at any time.

1. Bump the version with `scripts\bump.ps1` (`-Part patch|minor|major` or `-Set x.y.z`): it writes `engine\VERSION`
   (what the engine reports, and what makes running servers reload) and the same version in `engine\pyproject.toml`.
2. Add an entry to `CHANGELOG.md`: what changed, and what an upgrading PC should run.
3. Run `uv lock` and the tests: `cd engine; uv lock; uv run pytest -q; uv run ruff check .`
4. Run the test set against EPM on this PC (`uv run python -m smartass.testset run`, golden numbers outside git) and,
   when Excel code changed, `scripts\check_inproc.py` and `scripts\check_insert.py` (each starts an Excel of its
   own).
5. Commit and push; CI must be green.
6. Tag and push the tag: `git tag v<version>; git push origin v<version>`.
7. Roll back: check out the older tag (`git checkout v<version>`, then `uv sync --frozen`), or release the older code
   again under a new, higher version number.
