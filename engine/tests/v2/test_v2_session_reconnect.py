# engine/tests/v2/test_v2_session_reconnect.py
import psutil
from smartass.core import config as v1config
from smartass.core import session as v1session

from smartass import legacy


def configure(monkeypatch, tmp_path, rid, cwd):
    """Simulate a relay with this id and cwd: configure_v2 patches the v1 session module globals directly, so every
    one of them is registered with monkeypatch first (restored after the test)."""
    for name in ("_info", "_alive_relay", "set_workbook", "workbook", "_STATE"):
        monkeypatch.setattr(v1session, name, getattr(v1session, name))
    monkeypatch.setattr(v1config, "ANALYSER_DIR", tmp_path / "analyser")
    monkeypatch.setattr(v1config, "SVKIT", tmp_path)
    monkeypatch.setattr(psutil, "process_iter", lambda: [])
    monkeypatch.setattr(legacy, "_configured", False)
    monkeypatch.setattr(legacy, "_session_info", lambda: {"id": rid, "cwd": cwd, "project": "p"})
    legacy.configure_v2()


def test_reconnect_finds_the_last_workbook_of_the_same_folder(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"x")
    configure(monkeypatch, tmp_path, "v2-1-1", "C:/proj")
    v1session.set_workbook(str(book))
    assert v1session.workbook() == str(book)
    configure(monkeypatch, tmp_path, "v2-2-2", "C:/proj")            # new relay after /mcp
    assert v1session.workbook() == str(book)
    assert v1session._load()["v2-2-2"]["workbook"] == str(book)      # re-bound under the new id


def test_other_folder_and_deleted_file_do_not_match(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"x")
    configure(monkeypatch, tmp_path, "v2-1-1", "C:/proj")
    v1session.set_workbook(str(book))
    configure(monkeypatch, tmp_path, "v2-3-3", "C:/other")
    assert v1session.workbook() is None
    configure(monkeypatch, tmp_path, "v2-4-4", "C:/proj")
    book.unlink()
    assert v1session.workbook() is None


def test_last_entry_is_never_pruned(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"x")
    configure(monkeypatch, tmp_path, "v2-1-1", "C:/proj")
    v1session.set_workbook(str(book))
    configure(monkeypatch, tmp_path, "v2-2-2", "C:/elsewhere")
    v1session.set_workbook(str(book))                                # prunes dead ids
    saved = v1session._load()
    assert "_last:C:/proj" in saved and "v2-1-1" not in saved
