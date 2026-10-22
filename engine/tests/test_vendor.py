import os
import subprocess
import sys
from pathlib import Path

ENGINE = Path(__file__).resolve().parents[1]

SCRIPT = """
import smartview_mcp.styling as s
import smartview_mcp.formulas as f
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
    assert "smartview_mcp.vendor.tn_rapportstil" in r.stdout
