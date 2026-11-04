"""Plan 2D task 3: the CLI entry points (smartass CLI, python -m smartass.testset) take their own browser
profile slots and no warm browser, so they never collide with the running servers."""
import os
import subprocess
import sys

import pytest

READ = "from smartass.core import config; print(config.DRILL_PROFILE_SLOTS, config.DRILL_KEEP_WARM)"


def _env(**extra) -> dict:
    env = {k: v for k, v in os.environ.items()
           if not k.startswith("SV_DRILL_") and k not in ("PYTHONIOENCODING", "PYTHONUTF8")}
    env.update(extra)
    return env


def _run(code: str, **extra) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, "-B", "-c", code], env=_env(**extra), capture_output=True,
                          stdin=subprocess.DEVNULL, timeout=120)


@pytest.mark.parametrize("module", ["smartass.ops.cli", "smartass.testset.__main__"])
def test_the_cli_takes_its_own_browser_slots(module):
    out = _run(f"import {module}; {READ}")
    assert out.returncode == 0, out.stderr.decode(errors="replace")
    assert out.stdout.decode().split() == ["6", "False"]


def test_an_explicit_setting_wins():
    out = _run(f"import smartass.ops.cli; {READ}", SV_DRILL_PROFILE_SLOTS="2", SV_DRILL_KEEP_WARM="1")
    assert out.stdout.decode().split() == ["2", "True"]


def test_a_server_import_keeps_the_server_defaults():
    out = _run(f"import smartass.testset.runner; {READ}")
    assert out.stdout.decode().split() == ["3", "True"]
