"""An atlas built before 0.7.0 has no trusted side table of EPM's real names, and the REST engine's strict member
check is off on it (final review A, I1): `smartass setup` never replaces a local atlas that has one with a release
copy that has none, and `smartass doctor` (and with it fin_status's problems) says to rebuild once. Every member here
is invented."""
import os
import time

from smartass.core import config
from smartass.knowledge import atlas
from smartass.ops import doctor, setup_user


def _member(name, real=True):
    m = {"name": name, "label": name, "parent": "All Accounts", "depth": 1, "order": 1}
    return {**m, "real_name": name, "alias": None, "storage": "Store Data", "account_type": None,
            "time_balance": None} if real else m


def _atlas(path, real=True, dims=("All Accounts", "All Entities"), mtime=None):
    """An atlas whose dimensions were built over REST (real: a trusted side table) or by an older release."""
    db = atlas.connect(path)
    try:
        for d in dims:
            atlas.store(db, "FinMax", d, [_member("ac.X", real)], 1.0)
    finally:
        db.close()
    if mtime is not None:
        os.utime(path, (mtime, mtime))
    return path


def test_trust_counts_the_dimensions_whose_side_rows_are_this_builds(tmp_path):
    assert atlas.trust(_atlas(tmp_path / "new.db")) == (2, 2)
    assert atlas.trust(_atlas(tmp_path / "old.db", real=False)) == (0, 2)
    mixed = _atlas(tmp_path / "mixed.db")
    db = atlas.connect(mixed)
    atlas.store(db, "FinMax", "All Entities", [_member("ac.X", real=False)], 1.0)   # an older release rebuilt one
    db.close()
    assert atlas.trust(mixed) == (1, 2)


def test_trust_reads_without_writing_and_an_unreadable_file_is_none(tmp_path):
    p = _atlas(tmp_path / "a.db", real=False, mtime=time.time() - 3600)
    before = (p.stat().st_mtime, p.read_bytes())
    assert atlas.trust(p) == (0, 2) and (p.stat().st_mtime, p.read_bytes()) == before
    (tmp_path / "junk.db").write_bytes(b"not a database")
    assert atlas.trust(tmp_path / "junk.db") is None and atlas.trust(tmp_path / "missing.db") is None
    assert not (tmp_path / "missing.db").exists()


def _release(tmp_path, monkeypatch, local_real, release_real):
    rel = tmp_path / "release"
    (rel / "atlas").mkdir(parents=True)
    now = time.time()
    local = _atlas(tmp_path / "home" / "atlas.db", real=local_real, mtime=now - 7200) if local_real is not None \
        else tmp_path / "home" / "atlas.db"
    _atlas(rel / "atlas" / "atlas.db", real=release_real, mtime=now - 60)            # the newer copy
    monkeypatch.setattr(config, "RELEASE_FOLDER", str(rel))
    monkeypatch.setattr(config, "ATLAS_DB", local)
    return local


def test_setup_keeps_a_local_atlas_with_real_names_over_a_newer_release_copy_without(tmp_path, monkeypatch):
    local = _release(tmp_path, monkeypatch, local_real=True, release_real=False)
    said = setup_user.copy_atlas()
    assert said.startswith("atlas: local copy kept") and "real member names" in said and "smartass atlas build" in said
    assert atlas.trust(local) == (2, 2)


def test_setup_copies_a_newer_release_atlas_as_before_otherwise(tmp_path, monkeypatch):
    for local_real, release_real in ((False, False), (True, True), (False, True), (None, False)):
        sub = tmp_path / f"{local_real}-{release_real}"
        local = _release(sub, monkeypatch, local_real, release_real)
        assert setup_user.copy_atlas().startswith("atlas: copied from")
        assert atlas.trust(local) == ((2, 2) if release_real else (0, 2))


def test_doctor_warns_on_an_atlas_without_real_names_and_says_to_rebuild_once(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "ATLAS_DB", _atlas(tmp_path / "old.db", real=False))
    c = doctor._atlas()
    assert c.status == "WARN" and "no dimension has EPM's real member names" in c.detail
    assert c.fix == ("Rebuild once with `smartass atlas build` - until then the strict member check over Planning REST "
                     "is off.")
    monkeypatch.setattr(config, "ATLAS_DB", _atlas(tmp_path / "new.db"))
    assert doctor._atlas().status == "OK"
