import re
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
# Built from pieces so this file does not match itself.
BANNED = re.compile("|".join(["t99" + "2115", "prodreg" + "2026", "tom_" + "chamkrai", r"OneDrive - Telenor\\Pro" + "jects"]),
                    re.I)
GLOBS = ["engine/**/*.py", "engine/**/*.ps1", "engine/**/*.toml", "engine/**/*.bas", "installer/**/*",
         "skills/**/*.md", "scripts/**/*", "knowledge/*.py"]


def test_no_user_specific_paths_in_code_or_skills():
    hits = []
    for g in GLOBS:
        for f in REPO.glob(g):
            if not f.is_file() or ".venv" in f.parts or f.name == "test_paths.py":
                continue
            for n, line in enumerate(f.read_text(encoding="utf-8", errors="replace").splitlines(), 1):
                if BANNED.search(line):
                    hits.append(f"{f.relative_to(REPO)}:{n}: {line.strip()[:100]}")
    assert not hits, "\n".join(hits)
