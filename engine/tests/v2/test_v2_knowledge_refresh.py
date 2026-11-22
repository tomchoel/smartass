"""Plan 2A task 6: the monthly knowledge refresh (atlas diff, drafts, impact report) and approval. The atlas rebuild
and the test set are fakes: nothing here reaches EPM."""
import json
import sqlite3
import tomllib

import pytest
from kutil import ATLAS_ROWS, knowledge

from smartass.knowledge import refresh
from smartass.knowledge.build import summary
from smartass.knowledge.semantic import load_semantic


def snap(rows):
    out: dict = {}
    for cube, dim, name, label, parent in rows:
        out.setdefault(cube, {}).setdefault(dim, {})[name] = [parent, label]
    return out


OLD = [("FinMax", "All Entities", "un.SY.G910", "G910 Talkmore", "un.SY.2BUS"),
       ("FinMax", "All Entities", "un.SY.B1", "B1 One", "un.SY.4BUS_B"),
       ("FinMax", "All Entities", "un.SY.B2", "B2 Two", "un.SY.4BUS_B"),
       ("FinMax", "All Accounts", "ac.C60", "C60", "ac.C70")]
NEW = [("FinMax", "All Entities", "un.SY.G910", "G910 Talkmore", "un.SY.2CON"),       # moved
       ("FinMax", "All Entities", "un.SY.B1", "B1 First", "un.SY.4BUS_B"),            # relabelled
       ("FinMax", "All Entities", "un.SY.B3", "B3 Three", "un.SY.4BUS_B"),            # added (B2 removed)
       ("FinMax", "All Accounts", "ac.C60", "C60", "ac.C70")]


def test_diff_finds_added_removed_moved_and_relabelled():
    d = refresh.diff(snap(OLD), snap(NEW))
    assert [(x["member"], x["from"], x["to"]) for x in d.moved] == [("un.SY.G910", "un.SY.2BUS", "un.SY.2CON")]
    assert [x["member"] for x in d.removed] == ["un.SY.B2"]
    assert [x["member"] for x in d.added] == ["un.SY.B3"]
    assert [(x["member"], x["from"], x["to"]) for x in d.relabelled] == [("un.SY.B1", "B1 One", "B1 First")]
    assert d.counts() == {"added": 1, "removed": 1, "moved": 1, "relabelled": 1}


def test_snapshot_reads_entities_and_accounts_from_the_atlas(tmp_path):
    a, _, _ = knowledge(tmp_path)
    with sqlite3.connect(a) as db:
        s = refresh.snapshot(db)
    assert s["FinMax"]["All Entities"]["un.SY.G910"] == ["un.SY.2BUS", "G910 Talkmore Business"]
    assert "ac.C60" in s["FinMax"]["All Accounts"]


def test_drafts_for_referenced_moved_and_removed_members(tmp_path):
    _, sem, _ = knowledge(tmp_path)
    metrics, facts = load_semantic(sem)
    old = snap(OLD + [("FinMax", "All Entities", "un.SY.X9", "X9", "un.SY.2BUS")])
    new = snap(NEW + [("FinMax", "All Entities", "un.SY.X9", "X9", "un.SY.2TEC"),     # moved, referenced by nothing
                      ("FinMax", "All Accounts", "ac.C61", "C61", "ac.C60")])        # added child of a metric member
    new["FinMax"]["All Accounts"].pop("ac.C60")                                      # a metric member removed
    drafts, notes = refresh.draft_facts(refresh.diff(old, new), metrics, facts, (2026, 10))
    by = {f.id: f for f in drafts}
    g = by["refresh_2026_10_un_sy_g910"]                 # referenced by fact g910_moved
    assert g.status == "draft" and g.valid_from == "2026-10" and g.scope_members == {"FinMax": ["un.SY.G910"]}
    assert g.text == "un.SY.G910 moved from un.SY.2BUS to un.SY.2CON in Oct 26."
    c = by["refresh_2026_10_ac_c60"]
    assert c.severity == "warn" and c.scope_metrics == ["opex"] and "removed" in c.text
    assert not any("x9" in i or "c61" in i or "b2" in i for i in by)    # not referenced / added child / unreferenced
    assert any("opex" in n and "ac.C60" in n for n in notes)            # the build will fail: in the report


def _wire(tmp_path, monkeypatch, atlas_rows=None):
    """A home with an old snapshot (G910 under 2BUS... as the atlas has it, but B1 under a different parent)."""
    a, sem, k = knowledge(tmp_path, monkeypatch)
    home = tmp_path / "home"
    (home / "knowledge").mkdir(parents=True, exist_ok=True)
    old = {}
    for cube, dim, name, label, parent, *_ in ATLAS_ROWS:
        old.setdefault(cube, {}).setdefault(dim, {})[name] = [parent, label]
    old["FinMax"]["All Entities"]["un.SY.G910"] = ["un.SY.2CON", "G910 Talkmore Business"]   # it moved since
    (home / "knowledge" / "atlas_snapshot.json").write_text(json.dumps(old), encoding="utf-8")
    return a, sem, k, home


def test_refresh_end_to_end_with_fakes(tmp_path, monkeypatch):
    a, sem, k, home = _wire(tmp_path, monkeypatch)
    calls = []
    out = refresh.refresh(sem, a, k, home, (2026, 10), rebuild_atlas=lambda: calls.append("atlas") or [],
                          run_testset=lambda: calls.append("testset") or [{"id": "b01", "ok": True, "diffs": []}])
    assert calls == ["atlas", "testset"]
    assert [f.id for f in out.drafts] == ["refresh_2026_10_un_sy_g910"]
    assert (sem / "facts" / "refresh_2026_10_un_sy_g910.toml").exists()
    assert out.build.ok and summary(k)["drafts"] == 1 and summary(k)["refresh_month"] == "2026-10"
    report = (home / "knowledge" / "reports" / "2026-10.md").read_text(encoding="utf-8")
    assert out.report == home / "knowledge" / "reports" / "2026-10.md"
    for text in ("moved: 1", "refresh_2026_10_un_sy_g910", "1/1", "draft"):
        assert text in report, text
    snap_now = json.loads((home / "knowledge" / "atlas_snapshot.json").read_text(encoding="utf-8"))
    assert snap_now["FinMax"]["All Entities"]["un.SY.G910"][0] == "un.SY.2BUS"         # the new snapshot is saved


def test_a_refresh_whose_atlas_was_not_replaced_says_so_and_is_not_ok(tmp_path, monkeypatch):
    """Task 4 review, Minor 2: an atlas build that replaced nothing is not "rebuilt": the report says not replaced and
    why, and the refresh is not ok (the CLI exits non-zero); the rest of the refresh runs on the unchanged atlas."""
    a, sem, k, home = _wire(tmp_path, monkeypatch)
    not_replaced = [{"connection": "FinMax", "dimension": "All Accounts", "error": "HTTP 500"},
                    {"connection": "-", "error": "atlas not replaced: 1 dimension or cube failed"}]
    out = refresh.refresh(sem, a, k, home, (2026, 10), rebuild_atlas=lambda: not_replaced,
                          run_testset=lambda: [{"id": "b01", "ok": True, "diffs": []}])
    assert out.build.ok and not out.ok
    report = (home / "knowledge" / "reports" / "2026-10.md").read_text(encoding="utf-8")
    assert "Atlas: NOT replaced (atlas not replaced: 1 dimension or cube failed)" in report and "rebuilt" not in report


def test_a_refresh_without_a_sign_in_says_the_atlas_was_not_replaced(tmp_path, monkeypatch):
    """Task 4b review, Important 2: the monthly refresh's atlas build without a sign-in (the likeliest reason) is not
    "rebuilt": the report says NOT replaced and why, and the refresh is not ok."""
    from smartass.epm import members, oauth
    from smartass.ops import atlas_build
    a, sem, k, home = _wire(tmp_path, monkeypatch)
    monkeypatch.setattr(atlas_build.config, "ATLAS_DB", tmp_path / "rebuilt" / "atlas.db")

    class NoSignIn:
        def __getattr__(self, name):
            raise oauth.NotSignedIn("sign in")
    monkeypatch.setattr(members, "Source", NoSignIn)
    out = refresh.refresh(sem, a, k, home, (2026, 10), rebuild_atlas=atlas_build.run_build,
                          run_testset=lambda: [{"id": "b01", "ok": True, "diffs": []}])
    assert not out.ok and out.atlas_not_replaced
    report = (home / "knowledge" / "reports" / "2026-10.md").read_text(encoding="utf-8")
    assert "Atlas: NOT replaced (" in report and "not signed in to EPM" in report and "rebuilt" not in report


def test_refresh_without_atlas_rebuild_and_first_snapshot(tmp_path, monkeypatch):
    a, sem, k = knowledge(tmp_path, monkeypatch)
    home = tmp_path / "home"
    out = refresh.refresh(sem, a, k, home, (2026, 10), rebuild_atlas=None, run_testset=lambda: [])
    assert out.drafts == [] and out.first
    assert (home / "knowledge" / "atlas_snapshot.json").exists()
    assert "first refresh" in out.report.read_text(encoding="utf-8")


def test_approve_flips_the_status_and_rebuilds(tmp_path, monkeypatch):
    a, sem, k, home = _wire(tmp_path, monkeypatch)
    refresh.refresh(sem, a, k, home, (2026, 10), rebuild_atlas=None, run_testset=lambda: [])
    assert summary(k)["drafts"] == 1
    done, report = refresh.approve(sem, a, k, ["all"], by="tom", on="2026-11-04")
    assert done == ["refresh_2026_10_un_sy_g910"] and report.ok
    data = tomllib.loads((sem / "facts" / "refresh_2026_10_un_sy_g910.toml").read_text(encoding="utf-8"))
    assert data["status"] == "approved" and data["approved"] == {"by": "tom", "on": "2026-11-04"}
    assert summary(k)["drafts"] == 0
    assert summary(k)["refresh_month"] == "2026-10"                     # a rebuild keeps the refresh month
    with pytest.raises(KeyError, match="nope"):
        refresh.approve(sem, a, k, ["nope"], by="tom")


def test_approve_a_draft_metric_by_id(tmp_path, monkeypatch):
    a, sem, k = knowledge(tmp_path, monkeypatch)
    (sem / "metrics" / "cogs.toml").write_text('name = "COGS"\ndescription = "x"\nstatus = "draft"\n[finmax]\n'
                                               'members = ["ac.C60"]\n', encoding="utf-8")
    done, report = refresh.approve(sem, a, k, ["cogs"], by="tom", on="2026-11-04")
    assert done == ["cogs"] and report.ok
    m = {x.id: x for x in load_semantic(sem)[0]}["cogs"]
    assert m.status == "approved" and m.approved == {"by": "tom", "on": "2026-11-04"}


def test_due_flag(tmp_path, monkeypatch):
    a, sem, k, home = _wire(tmp_path, monkeypatch)
    assert refresh.due(home, k) == {"refresh_month": None, "closed_seen": None, "due": False}
    refresh.note_closed(home, (2026, 10))
    assert refresh.due(home, k)["due"] is True                          # never refreshed, a closed month is known
    refresh.refresh(sem, a, k, home, (2026, 10), rebuild_atlas=None, run_testset=lambda: [])
    assert refresh.due(home, k) == {"refresh_month": "2026-10", "closed_seen": "2026-10", "due": False}
    refresh.note_closed(home, (2026, 9))                                # an older month never moves it back
    refresh.note_closed(home, (2026, 11))
    assert refresh.due(home, k) == {"refresh_month": "2026-10", "closed_seen": "2026-11", "due": True}


def test_fin_status_says_refresh_due(tmp_path, monkeypatch):
    import asyncio

    from smartass import server
    from smartass.core import config
    from smartass.ops import doctor
    a, sem, k, home = _wire(tmp_path, monkeypatch)
    monkeypatch.setattr(config, "HOME", home)
    monkeypatch.setattr(server.atlas, "status", lambda db: {"stale": False, "oldest_age_days": 3.0, "built": [1]})
    monkeypatch.setattr(doctor, "run", lambda: [])
    refresh.note_closed(home, (2026, 10))
    out = asyncio.run(server.fin_status())
    assert out["knowledge"]["refresh_due"] is True
    assert "knowledge refresh due" in out["knowledge"]["note"] and "smartass knowledge refresh" in out["knowledge"]["note"]


def test_run_prepare_notes_the_closed_month(tmp_path, monkeypatch):
    from test_v2_pipeline import FakeProvider
    from v2util import facts, spec_dict

    from smartass import pipeline
    from smartass.core import config
    _, _, _, home = _wire(tmp_path, monkeypatch)
    monkeypatch.setattr(config, "HOME", home)
    pipeline.run_prepare(spec_dict(), FakeProvider(facts()))
    assert refresh.due(home, config.KNOWLEDGE_DB)["closed_seen"] == "2026-09"


def test_cli_status_and_approve(tmp_path, monkeypatch, capsys):
    from smartass.core import config
    from smartass.ops import cli
    a, sem, k, home = _wire(tmp_path, monkeypatch)
    monkeypatch.setattr(config, "HOME", home)
    refresh.refresh(sem, a, k, home, (2026, 10), rebuild_atlas=None, run_testset=lambda: [])
    assert cli.main(["knowledge", "status"]) == 0
    out = capsys.readouterr().out
    assert "refresh_2026_10_un_sy_g910" in out and "2026-10.md" in out
    assert cli.main(["knowledge", "approve", "all", "--by", "tom"]) == 0
    assert "approved 1" in capsys.readouterr().out
    assert summary(k)["drafts"] == 0


def test_cli_refresh_no_atlas_runs_the_offline_testset(tmp_path, monkeypatch, capsys):
    from smartass.core import config
    from smartass.ops import cli
    from smartass.testset import runner
    a, sem, k, home = _wire(tmp_path, monkeypatch)
    monkeypatch.setattr(config, "HOME", home)
    monkeypatch.setattr(runner, "run_offline", lambda cases: [{"id": c["id"], "ok": True, "diffs": []} for c in cases])
    from smartass.ops import atlas_build
    monkeypatch.setattr(atlas_build, "run_build", lambda *a, **k: pytest.fail("--no-atlas must not rebuild"))
    assert cli.main(["knowledge", "refresh", "--no-atlas", "--month", "2026-10"]) == 0
    out = capsys.readouterr().out
    assert "2026-10.md" in out and "9/9" in out          # nine test-set cases since b09 (Plan 2C)
