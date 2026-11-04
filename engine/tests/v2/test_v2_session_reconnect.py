# engine/tests/v2/test_v2_session_reconnect.py
import psutil

from smartass.core import config, session


def configure(monkeypatch, tmp_path, rid, cwd):
    """Simulate a relay with this id and cwd (a new id = a new relay after /mcp)."""
    monkeypatch.setattr(config, "ANALYSER_DIR", tmp_path / "analyser")
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(session, "_STATE", tmp_path / "sessions.json")
    monkeypatch.setattr(psutil, "process_iter", lambda: [])
    monkeypatch.setattr(session, "_info", {"id": rid, "cwd": cwd, "project": "p"})


def test_reconnect_finds_the_last_workbook_of_the_same_folder(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"x")
    configure(monkeypatch, tmp_path, "1-1", "C:/proj")
    session.set_workbook(str(book))
    assert session.workbook() == str(book)
    configure(monkeypatch, tmp_path, "2-2", "C:/proj")               # new relay after /mcp
    assert session.workbook() == str(book)
    assert session._load()["2-2"]["workbook"] == str(book)           # re-bound under the new id


def test_other_folder_and_deleted_file_do_not_match(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"x")
    configure(monkeypatch, tmp_path, "1-1", "C:/proj")
    session.set_workbook(str(book))
    configure(monkeypatch, tmp_path, "3-3", "C:/other")
    assert session.workbook() is None
    configure(monkeypatch, tmp_path, "4-4", "C:/proj")
    book.unlink()
    assert session.workbook() is None


def test_last_entry_is_never_pruned(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"x")
    configure(monkeypatch, tmp_path, "1-1", "C:/proj")
    session.set_workbook(str(book))
    configure(monkeypatch, tmp_path, "2-2", "C:/elsewhere")
    session.set_workbook(str(book))                                  # prunes dead ids
    saved = session._load()
    assert "_last:C:/proj" in saved and "1-1" not in saved


def test_analysis_files_go_to_the_plain_analyser_folder(monkeypatch, tmp_path):
    configure(monkeypatch, tmp_path, "1-1", "C:/proj")
    assert session.project_dir() == tmp_path / "analyser" / "p"
