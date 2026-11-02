import importlib.util
from pathlib import Path

_p = Path(__file__).resolve().parents[2] / "scripts" / "check_staged.py"
_spec = importlib.util.spec_from_file_location("check_staged", _p)
cs = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cs)


def test_blocks_databases_and_stray_workbooks():
    assert cs.problems("engine/atlas.db", b"") and cs.problems("x/drill.sqlite", b"")
    assert cs.problems("docs/analyse.xlsx", b"")
    assert not cs.problems("packs/produktregnskap/knowledge/source_x.xlsx", b"")
    assert not cs.problems("engine/smartview_mcp/knowledge/map.xlsx", b"")


def test_blocks_tokens_unless_marked():
    tok = "sso_token=" + "a" * 30
    assert cs.problems("a.py", tok.encode())
    assert cs.problems("a.py", ("x = 'eyJ" + "a" * 20 + "." + "b" * 20 + "'").encode())
    assert cs.problems("a.toml", ('secret = "' + "z" * 24 + '"').encode())
    assert not cs.problems("a.py", (tok + "  # allow-secret").encode())
    assert not cs.problems("a.toml", b'secret = "{secret}"')


