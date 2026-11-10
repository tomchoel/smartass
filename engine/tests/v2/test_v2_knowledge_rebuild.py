"""Plan 2C task 1: knowledge.db is rebuilt when the semantic files changed (the stored hash differs), not only when
it is missing."""
import pytest
from kutil import knowledge
from test_v2_pipeline import FakeProvider
from v2util import facts, spec_dict

from smartass import pipeline
from smartass.knowledge.build import summary
from smartass.knowledge.resolver import ensure_built


def test_unchanged_semantic_files_do_not_rebuild(tmp_path, monkeypatch):
    _, _, k = knowledge(tmp_path, monkeypatch)
    before = summary(k)["built_at"], k.stat().st_mtime_ns
    notes: list[str] = []
    ensure_built(notes)
    assert (summary(k)["built_at"], k.stat().st_mtime_ns) == before
    assert notes == []


def test_a_changed_semantic_file_rebuilds_with_a_warning(tmp_path, monkeypatch):
    _, sem, k = knowledge(tmp_path, monkeypatch)
    old = summary(k)["semantic_hash"]
    p = sem / "facts" / "opex_note.toml"
    p.write_text(p.read_text(encoding="utf-8").replace("Opex note.", "Opex note, changed."), encoding="utf-8")
    pre = pipeline.run_prepare(spec_dict(), FakeProvider(facts()))
    assert summary(k)["semantic_hash"] != old
    assert "knowledge.db rebuilt: semantic files or the atlas changed" in pre.warnings
    assert any(c["text"] == "Opex note, changed." for c in pre.caveats)


def test_a_new_semantic_file_rebuilds(tmp_path, monkeypatch):
    _, sem, k = knowledge(tmp_path, monkeypatch)
    old = summary(k)["semantic_hash"]
    (sem / "facts" / "extra.toml").write_text('text = "Extra."\nscope_metrics = ["opex"]\nsource = "t"\n'
                                              'status = "approved"\n', encoding="utf-8")
    notes: list[str] = []
    ensure_built(notes)
    assert summary(k)["semantic_hash"] != old
    assert notes == ["knowledge.db rebuilt: semantic files or the atlas changed"]


def test_a_failing_rebuild_stops_the_answer_naming_the_error(tmp_path, monkeypatch):
    from smartass.model.validate import GuardError
    _, sem, k = knowledge(tmp_path, monkeypatch)
    old = summary(k)["semantic_hash"]
    (sem / "metrics" / "bad.toml").write_text('name = "Bad"\ndescription = "x"\n[finmax]\nmembers = ["ac.NOPE"]\n',
                                              encoding="utf-8")
    with pytest.raises(GuardError, match="ac.NOPE"):
        pipeline.run_prepare(spec_dict(), FakeProvider(facts()))
    assert summary(k)["semantic_hash"] == old                 # the previous knowledge.db is left untouched
