# engine/tests/v2/test_v2_serverkit.py
import subprocess
import sys

CODE = ("import sys, smartass.server; "
        "print('LEAK' if any(m.startswith('smartview_mcp') for m in sys.modules) else 'CLEAN')")


def test_importing_the_v2_server_does_not_import_v1():
    out = subprocess.run([sys.executable, "-c", CODE], capture_output=True, text=True, timeout=120,
                         stdin=subprocess.DEVNULL)
    assert out.returncode == 0, out.stderr[-2000:]
    assert out.stdout.strip().splitlines()[-1] == "CLEAN"


def test_serverkit_names_the_v2_server_uses():
    from smartass.core import serverkit as kit
    for name in ("lifespan", "guard", "_err", "with_next_step", "explain", "READ", "LOCAL", "WRITE", "ADMIN",
                 "_call_user", "_with_progress", "_ctx", "progress_secs", "_is_our_analysis_file"):
        assert hasattr(kit, name), name
