"""`smartass atlas build`'s CLI wiring and the indented-label reader the atlas's rows come from (grid.tree_from_labels).
The Excel build is gone (Task 6): the build over REST is tested in test_atlas_rest.py."""
from smartass.epm.grid import tree_from_labels
from smartass.ops import atlas_build, cli


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
    by = {x["name"]: x for x in tree_from_labels(labels)}
    assert by["un.SY"]["depth"] == 0 and by["un.SY"]["parent"] is None
    assert by["un.SY.2WHS"]["parent"] == "un.SY"
    assert by["un.SY.3WHS_W"]["parent"] == "un.SY.2WHS"
    assert by["un.SY.2BUS"]["parent"] == "un.SY"
