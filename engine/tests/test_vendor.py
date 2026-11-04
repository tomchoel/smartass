import os
import subprocess
import sys
from pathlib import Path

ENGINE = Path(__file__).resolve().parents[1]

SCRIPT = """
import smartass.xl.styling as s
import smartass.xl.formulas as f
r = f.check([("S", "C5", "=C1+C2+C3")])
assert r["ok"] is not None, r
assert r["checked"] == 1, r
print("TN", s.tn.__name__)
"""


def test_style_helpers_do_not_depend_on_the_user_profile(tmp_path):
    env = dict(os.environ)
    for k in ("USERPROFILE", "HOME"):
        env[k] = str(tmp_path)
    env.pop("TN_RAPPORTSTIL_DIR", None)
    env["PYTHONUTF8"] = "1"
    r = subprocess.run([sys.executable, "-c", SCRIPT], cwd=ENGINE, env=env, capture_output=True, text=True,
                       timeout=120)
    assert r.returncode == 0, r.stdout + r.stderr
    assert "smartass.xl.vendor.tn_rapportstil" in r.stdout


# sha256 (LF line ends) of the vendored advisor texts as of 2b6c237: their wording is the vendor's, never edited here
# (NOTICE.md as of Plan 2A task 7: findings.md and gate.md, v1-only, were deleted and their lines dropped from it)
VENDORED = {"variance-commentary.md": "85d197d36cd57cebfb047ddebf0c94c17441c6066530fa827a1b4d2b38e4218f",
            "NOTICE.md": "ec103aa63bc174fb4dc4f00afdf70ee40815dc64ca29ad51830139296fd78836"}


def test_v1_only_advisor_files_are_gone():
    advisor = ENGINE.parent / "knowledge" / "advisor"
    assert sorted(p.name for p in advisor.iterdir()) == ["NOTICE.md", "variance-commentary.md"]
    notice = (advisor / "NOTICE.md").read_text(encoding="utf-8")
    assert "findings.md" not in notice and "gate.md" not in notice and "variance-commentary.md" in notice


def test_vendored_advisor_text_unchanged():
    import hashlib
    advisor = ENGINE.parent / "knowledge" / "advisor"
    for name, sha in VENDORED.items():
        text = (advisor / name).read_bytes().replace(b"\r\n", b"\n")
        assert hashlib.sha256(text).hexdigest() == sha, f"{name} was edited"
