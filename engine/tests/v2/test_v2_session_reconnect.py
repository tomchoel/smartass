# engine/tests/v2/test_v2_session_reconnect.py
import psutil

from smartass.core import config, session


def configure(monkeypatch, tmp_path, rid, cwd, client="claude-1"):
    """Simulate a relay with this id, cwd and client (a new id = a new relay after /mcp; a new client = a new
    session, e.g. Copilot after Claude Code)."""
    monkeypatch.setattr(config, "ANALYSER_DIR", tmp_path / "analyser")
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(session, "_STATE", tmp_path / "sessions.json")
    monkeypatch.setattr(psutil, "process_iter", lambda: [])
    monkeypatch.setattr(session, "_info", {"id": rid, "cwd": cwd, "project": "p", "client": client})


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


def test_a_new_session_in_the_same_folder_gets_a_new_workbook(monkeypatch, tmp_path):
    """10.11: a Copilot session asked about Wholesale and got the folder's old Streamix workbook."""
    book = tmp_path / "streamix.xlsx"
    book.write_bytes(b"x")
    configure(monkeypatch, tmp_path, "1-1", "C:/proj", client="claude-1")
    session.set_workbook(str(book))
    configure(monkeypatch, tmp_path, "2-2", "C:/proj", client="vscode-9")       # another agent, same folder
    assert session.workbook() is None
    configure(monkeypatch, tmp_path, "3-3", "C:/proj", client="")               # unknown client: never a takeover
    assert session.workbook() is None


def test_an_old_last_entry_without_a_client_is_not_taken_over(monkeypatch, tmp_path):
    book = tmp_path / "old.xlsx"
    book.write_bytes(b"x")
    (tmp_path / "sessions.json").write_text(
        '{"_last:C:/proj": {"workbook": "%s", "ts": 1}}' % str(book).replace("\\", "/"), encoding="utf-8")
    configure(monkeypatch, tmp_path, "5-5", "C:/proj")
    assert session.workbook() is None
