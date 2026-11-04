"""Plan 2D task 4: the CLI entry points (smartass CLI, python -m smartass.testset) print UTF-8 on a pipe, whatever
the Windows code page."""
import pytest
from test_v2_cli_env import _run


@pytest.mark.parametrize("call", [
    "from smartass.ops import cli; cli.main(['--version-that-does-not-exist'])",
    "from smartass.testset.__main__ import main; main(['nonsense'])",
])
def test_the_cli_prints_utf8(call):
    """The entry point reconfigures stdout and stderr to UTF-8 before printing: '·' arrives as UTF-8 on a pipe."""
    code = f"import sys\ntry:\n    {call}\nexcept SystemExit:\n    pass\nprint('\u00b7'); print('\u00b7', file=sys.stderr)"
    out = _run(code)
    assert out.stdout.decode("utf-8").strip() == "·"
    assert out.stderr.decode("utf-8").strip().endswith("·")
