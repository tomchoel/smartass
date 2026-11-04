"""`smartass atlas build`: dimension trees from Smart View into the atlas, errors per dimension, CLI wiring."""
import asyncio
import contextlib

from smartass.epm import ops
from smartass.knowledge import atlas
from smartass.ops import atlas_build, cli
from smartass.xl.host import HostError


async def _call(fn, timeout=None, what=""):
    return fn(None)


def _fake_cube(monkeypatch, fail=()):
    trees = {"All Entities": [{"label": "SY Telenor Norge AS (un)", "name": "un.SY", "depth": 0, "parent": None,
                               "order": 0}],
             "All Accounts": [{"label": "C30 Revenue (ac)", "name": "ac.C30", "depth": 0, "parent": None,
                               "order": 0}]}
    monkeypatch.setattr(ops, "dimensions", lambda h, conn: list(trees))
    monkeypatch.setattr(ops, "canon_dim", lambda conn, d: d)

    def dump(h, conn, dim):
        if dim in fail:
            raise HostError(f"{dim} timed out")
        return trees[dim]
    monkeypatch.setattr(ops, "dump_dimension", dump)


def test_build_stores_every_dimension_and_records_failures(tmp_path, monkeypatch):
    _fake_cube(monkeypatch, fail=("All Accounts",))
    with contextlib.closing(atlas.connect(tmp_path / "atlas.db")) as db:
        done = asyncio.run(atlas_build.build(_call, ["FinMax"], db=db, progress=lambda m: None))
        assert [d.get("members") for d in done] == [1, None] and "timed out" in done[1]["error"]
        assert atlas.has_member(db, "FinMax", "All Entities", "un.SY")
        st = {(b["connection"], b["dimension"]): b for b in atlas.status(db)["built"]}
        assert st[("FinMax", "All Accounts")]["error"]


def test_build_of_some_dimensions_does_not_purge_the_cube(tmp_path, monkeypatch):
    _fake_cube(monkeypatch)
    with contextlib.closing(atlas.connect(tmp_path / "atlas.db")) as db:
        asyncio.run(atlas_build.build(_call, ["FinMax"], db=db, progress=lambda m: None))
        done = asyncio.run(atlas_build.build(_call, ["FinMax"], ["All Accounts"], db=db, progress=lambda m: None))
        assert [d["dimension"] for d in done] == ["All Accounts"]
        assert atlas.has_member(db, "FinMax", "All Entities", "un.SY")      # kept


def test_cli_atlas_build_reports_and_fails_on_errors(monkeypatch, capsys):
    monkeypatch.setattr(atlas_build, "run_build", lambda cubes, dims: [{"connection": "FinMax", "dimension": "x",
                                                                        "members": 3},
                                                                       {"connection": "FinPlan", "error": "down"}])
    assert cli.main(["atlas", "build", "--cube", "FinMax"]) == 1
    assert "1 built, 1 failed" in capsys.readouterr().out


def test_tree_from_labels_bottom_up():
    # ancestors-at-bottom ordering as Smart View returns it with the user's 'ancestor position' setting
    labels = ["          3WHS_W Wholesale (un.SY)", "     2WHS Wholesale (un.SY)", "     2BUS Business (un.SY)",
              "SY Telenor Norge AS (un)"]
    by = {x["name"]: x for x in ops.tree_from_labels(labels)}
    assert by["un.SY"]["depth"] == 0 and by["un.SY"]["parent"] is None
    assert by["un.SY.2WHS"]["parent"] == "un.SY"
    assert by["un.SY.3WHS_W"]["parent"] == "un.SY.2WHS"
    assert by["un.SY.2BUS"]["parent"] == "un.SY"
