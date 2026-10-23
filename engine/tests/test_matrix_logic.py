from smartview_mcp.auditlib import error_class, is_engine_excel, merge_before, poll_kill, stale_suspended, verdict


def test_error_class_normalises_digits_and_keeps_first_line():
    a = error_class("HostError: drill page never asked for its data within 45s\nnext: ...")
    b = error_class("HostError: drill page never asked for its data within 47s")
    assert a == b == "HostError: drill page never asked for its data within #s"
    assert error_class(" \n ") == ""


def test_verdict_rules():
    assert verdict(ok=True, secs=10, budget=30) == "PASS"
    assert verdict(ok=True, secs=31, budget=30) == "SLOW"
    assert verdict(ok=False, secs=5, budget=30) == "FAIL"
    assert verdict(ok=False, secs=5, budget=30, limit="no FinLite in test env") == "LIMIT"


def test_merge_before_adds_previous_result():
    now = [{"id": "S1.query", "result": "PASS"}]
    before = [{"id": "S1.query", "result": "FAIL"}, {"id": "gone", "result": "PASS"}]
    assert merge_before(now, before)[0]["before"] == "FAIL"


SCRATCH = r"C:\Users\me\OneDrive - Telenor\smartass\analyser\_matrix"
EMBED = r'"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" /automation -Embedding'
OWNERS = {500: 1000.0, 600: 2000.0}                       # the matrix's own child_server.py: pid -> create time
REG = [{"pid": 42, "ctime": 3000.0, "owner": 500, "owner_ctime": 1000.0},
       {"pid": 43, "ctime": 3100.0, "owner": 700, "owner_ctime": 1500.0}]    # 43: another session's server


def _own(cmdline, pid=1, ctime=0.0, registry=(), owners=None):
    return is_engine_excel(cmdline, SCRATCH, pid, ctime, list(registry), OWNERS if owners is None else owners)


def test_is_engine_excel_embedding_registered_to_a_matrix_server():
    assert _own(EMBED, 42, 3000.4, REG)


def test_is_engine_excel_rejects_unregistered_or_foreign_embedding():
    assert not _own(EMBED, 99, 3000.0, REG)              # not in pids.json: Excel MCP, a DispatchEx script, ...
    assert not _own(EMBED, 43, 3100.0, REG)              # registered by another smartass server (other owner)
    assert not _own(EMBED, 42, 3005.0, REG)              # pid reused: create time differs
    assert not _own(EMBED, 42, 3000.0, [{**REG[0], "owner_ctime": 999.0}])   # owner pid reused
    assert not _own(EMBED, 42, 3000.0, REG, owners={})   # no matrix client running
    assert not _own('"EXCEL.EXE"', 42, 3000.0, REG)      # registered pid, but not an automation Excel (user's)
    assert not _own(EMBED, 42, 3000.0, [{"pid": 42}])    # malformed entry


def test_is_engine_excel_x_in_scratch_only():
    assert _own(f'/x "{SCRATCH}\\a.xlsx"')
    assert _own(f'"EXCEL.EXE" /x "{SCRATCH.lower()}\\a.xlsx"')
    assert not _own(r'/x "C:\other\b.xlsx"')
    assert not _own(f'"EXCEL.EXE" /x "{SCRATCH}_other\\b.xlsx"')


def test_is_engine_excel_rejects_user_excel_and_other_apps():
    assert not _own('"EXCEL.EXE"')
    assert not _own(f'"EXCEL.EXE" "{SCRATCH}\\a.xlsx"')
    assert not _own('"POWERPNT.EXE"')
    assert not _own('"POWERPNT.EXE" /automation -Embedding', 42, 3000.0, REG)


def test_stale_suspended_resumes_only_same_process():
    marker = [{"pid": 10, "ctime": 100.0}, {"pid": 11, "ctime": 200.0}, {"pid": 12, "ctime": 300.0}, {"pid": "x"}]
    running = {10: 100.3, 11: 250.0}                     # 11 is a new process on a reused pid; 12 is gone
    assert stale_suspended(marker, running) == [10]
    assert stale_suspended([], running) == []


class _Clock:
    def __init__(self):
        self.t = 0.0

    def __call__(self):
        return self.t

    def sleep(self, s):
        self.t += s


def test_poll_kill_kills_as_soon_as_it_appears():
    clock, tries = _Clock(), []

    def kill():
        tries.append(clock.t)
        return [7] if clock.t >= 10 else []
    killed, running = poll_kill(kill, lambda: False, 60, 1, clock.sleep, clock)
    assert (killed, running) == ([7], True)
    assert tries[-1] == 10


def test_poll_kill_misses_when_call_finishes_or_times_out():
    clock = _Clock()
    killed, running = poll_kill(lambda: [], lambda: clock.t >= 5, 60, 1, clock.sleep, clock)
    assert (killed, running) == ([], False)
    clock = _Clock()
    killed, running = poll_kill(lambda: [], lambda: False, 60, 1, clock.sleep, clock)
    assert killed == [] and running and 60 <= clock.t < 62


def test_before_file_compares_like_with_like(tmp_path, monkeypatch):
    import json
    import os
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "live"))
    import matrix
    monkeypatch.setattr(matrix, "AUDIT", tmp_path)
    for i, name in enumerate(("2026-10-23-matrix-baseline", "2026-10-23-matrix-baseline-fresh")):
        p = tmp_path / f"{name}.json"
        p.write_text(json.dumps({"results": [{"id": name}]}))
        os.utime(p, (1000 + i, 1000 + i))
    assert matrix.previous(fresh=False)[0].name == "2026-10-23-matrix-baseline.json"
    assert matrix.previous(fresh=True)[0].name == "2026-10-23-matrix-baseline-fresh.json"
    assert matrix.previous(fresh=True, name="2026-10-23-matrix-baseline")[0].name == "2026-10-23-matrix-baseline.json"
