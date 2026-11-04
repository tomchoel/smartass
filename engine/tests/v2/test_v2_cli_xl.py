"""Plan 3A: `smartass xl changes|revert|check` - the Excel task pane's deterministic actions. Each prints ONE JSON
object on stdout: {"ok": true, "result": ...} or {"ok": false, "error": "..."} (the pane's api.ts reads that)."""
import json

import pytest

from smartass.ops import cli


@pytest.fixture
def fakes(monkeypatch):
    calls = []

    def user_job(fn):
        calls.append("job")
        return fn()

    monkeypatch.setattr(cli, "_user_job", user_job)
    monkeypatch.setattr(cli, "_xl_changes", lambda wb, limit: calls.append(("changes", wb, limit)) or
                        [{"id": "c1", "ts": "2026-11-04T10:00:00", "workbook": "B.xlsx", "sheet": "S", "op": "write",
                          "address": "A1:B2", "label": "xl_snapshot", "reverted": False}])
    monkeypatch.setattr(cli, "_xl_revert", lambda cid, wb, force: calls.append(("revert", cid, wb, force)) or
                        {"reverted": cid or "last", "workbook": "B.xlsx", "note": "Not saved."})
    monkeypatch.setattr(cli, "_xl_check", lambda wb, sheet: calls.append(("check", wb, sheet)) or
                        {"ok": True, "checked": 3, "counts": {}, "block": [], "warn": [], "block_cells": [],
                         "workbook": "B.xlsx", "sheets": ["S"]})
    return calls


def _out(capsys) -> dict:
    lines = capsys.readouterr().out.strip().splitlines()
    assert len(lines) == 1, lines                     # exactly one JSON object
    return json.loads(lines[0])


def test_changes_defaults(fakes, capsys):
    assert cli.main(["xl", "changes"]) == 0
    out = _out(capsys)
    assert out["ok"] is True
    assert out["result"][0]["id"] == "c1" and set(out["result"][0]) == {
        "id", "ts", "workbook", "sheet", "op", "address", "label", "reverted"}
    assert ("changes", None, 30) in fakes
    assert "job" not in fakes                         # the change log is files only: no Excel turn


def test_changes_workbook_and_limit(fakes, capsys):
    assert cli.main(["xl", "changes", "--workbook", "Bok æøå.xlsx", "--limit", "5"]) == 0
    assert _out(capsys)["ok"] is True
    assert ("changes", "Bok æøå.xlsx", 5) in fakes


def test_revert_args_run_in_a_user_excel_job(fakes, capsys):
    assert cli.main(["xl", "revert", "--id", "c9", "--workbook", "B", "--force"]) == 0
    assert _out(capsys)["result"]["reverted"] == "c9"
    assert fakes == ["job", ("revert", "c9", "B", True)]


def test_revert_defaults_latest_without_force(fakes, capsys):
    assert cli.main(["xl", "revert"]) == 0
    _out(capsys)
    assert ("revert", None, None, False) in fakes


def test_check_args(fakes, capsys):
    assert cli.main(["xl", "check", "--workbook", "B", "--sheet", "Q1 Answer"]) == 0
    out = _out(capsys)
    assert out["result"]["checked"] == 3 and out["result"]["ok"] is True
    assert fakes == ["job", ("check", "B", "Q1 Answer")]


def test_an_error_is_one_json_object_with_exit_1(monkeypatch, capsys):
    from smartass.xl.host import InputError

    def boom(cid, wb, force):
        raise InputError("S!A1 was edited after change c1; pass force=true to revert anyway.")
    monkeypatch.setattr(cli, "_user_job", lambda fn: fn())
    monkeypatch.setattr(cli, "_xl_revert", boom)
    assert cli.main(["xl", "revert", "--id", "c1"]) == 1
    out = _out(capsys)
    assert out == {"ok": False, "error": "S!A1 was edited after change c1; pass force=true to revert anyway."}


def test_non_ascii_survives(fakes, capsys, monkeypatch):
    monkeypatch.setattr(cli, "_xl_check", lambda wb, sheet: {"ok": False, "workbook": "Bøk.xlsx"})
    cli.main(["xl", "check"])
    assert _out(capsys)["result"]["workbook"] == "Bøk.xlsx"


def test_user_job_binds_and_releases_the_turns(monkeypatch):
    """The COM work runs as v2's user-Excel jobs do: a job token bound, per-Excel turns released afterwards."""
    from smartass.xl import jobs
    seen = []
    monkeypatch.setattr(jobs, "bind_job", lambda token=None: seen.append("bind"))
    monkeypatch.setattr(jobs, "release_turns", lambda: seen.append("release"))
    assert cli._user_job(lambda: seen.append("fn") or 42) == 42
    assert seen == ["bind", "fn", "release"]
    with pytest.raises(RuntimeError):
        cli._user_job(lambda: (_ for _ in ()).throw(RuntimeError("x")))
    assert seen[-1] == "release"


def test_the_real_functions_are_wired(monkeypatch):
    from smartass import workbench
    from smartass.xl.workbench import history
    monkeypatch.setattr(history, "changes", lambda wb, limit: ("hist", wb, limit))
    monkeypatch.setattr(workbench, "revert", lambda cid, wb, force: ("rev", cid, wb, force))
    monkeypatch.setattr(workbench, "_check", lambda wb, sheet: ("chk", wb, sheet))
    assert cli._xl_changes("B", 7) == ("hist", "B", 7)
    assert cli._xl_revert("c1", "B", True) == ("rev", "c1", "B", True)
    assert cli._xl_check("B", "S") == ("chk", "B", "S")
