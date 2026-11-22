"""An atlas built before 0.7.0 has no trusted side table of EPM's real names, and the REST engine's strict member
check is off on it (final review A, I1): `smartass doctor` (and with it fin_status's problems) says to rebuild once.
Every member here is invented."""
import os
import time

from smartass.core import config
from smartass.knowledge import atlas
from smartass.ops import doctor


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


def test_the_read_only_uri_keeps_a_drive_and_puts_a_unc_host_in_the_path():
    """SQLite takes no URI authority: 'file://host/share/...' fails to open (probed on a UNC path to this PC), and
    resolve() turns a mapped drive into such a path - an atlas can be read from a network share."""
    assert atlas._read_only_uri(r"\\server\share\rel\atlas.db") == "file:////server/share/rel/atlas.db?mode=ro"
    assert atlas._read_only_uri(r"Z:\rel\atlas.db") == "file:///Z:/rel/atlas.db?mode=ro"


def test_doctor_warns_on_an_atlas_without_real_names_and_says_to_rebuild_once(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "ATLAS_DB", _atlas(tmp_path / "old.db", real=False))
    c = doctor._atlas()
    assert c.status == "WARN" and "no dimension has EPM's real member names" in c.detail
    assert c.fix == ("Rebuild once with `smartass atlas build` - until then the strict member check over Planning REST "
                     "is off.")
    monkeypatch.setattr(config, "ATLAS_DB", _atlas(tmp_path / "new.db"))
    assert doctor._atlas().status == "OK"
