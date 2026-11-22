"""Light-fast phase 2 Task 5 (and Task 6, which deleted it): nothing on a tool call's path starts an Excel of its own.
Every tool that works in the user's Excel - fin_analyze (a new question and a follow-up), fin_drill, fin_refresh,
fin_critic, xl_snapshot / xl_revert / xl_tidy (the workbench) - runs its job through the real runner
(serverkit._call_user -> xl/userjobs.UserJobs, a thread per job) and resolves its workbook through the real
xl/target.resolve_target (the user's Excel found, its turn taken, made ready): the only code of a job
that uses the object the runner hands it (h). The Excel is a small fake; what a job draws after its workbook is found
is replaced (it never uses h). fin_status, with and without restart, starts nothing either. conftest makes a real
Excel start impossible; this pins that the code does not even try: a launch of Excel fails the test (the user's
Excel is launched only when none is running, which these tests do not set up - except the cold path at the end).

Task 6 replaced the two paths that still handed a job's work to the hidden Excel: the bridge add-in is gone, and a new
analysis file with no Excel running and no labelled seed is built without Excel (both pinned at the end)."""
from __future__ import annotations

import asyncio
import threading
from pathlib import Path

import pytest
import pywintypes
from test_v2_pipeline import FakeProvider
from v2util import facts, spec_dict

from smartass import exceljob, server
from smartass.core import config, serverkit
from smartass.core.errors import HostError
from smartass.xl import instances, sheets
from smartass.xl import target as xltarget

PID = 4_190_303                                   # no such process: no window, no dialog, a turn file of its own
NOT_IN_COLLECTION = -2147352567


class FakeBook:
    def __init__(self, path: Path):
        self.Name, self.FullName, self.ReadOnly = path.name, str(path), False
        self.ActiveSheet = type("Sheet", (), {"Name": "Q1 Answer"})()

    def Activate(self):
        pass


class FakeWorkbooks:
    def __init__(self, app):
        self.app = app

    def __call__(self, name):
        for b in self.app.books:
            if b.Name == name:
                return b
        raise pywintypes.com_error(NOT_IN_COLLECTION, "not in collection", None, None)

    def __iter__(self):
        return iter(list(self.app.books))

    def Open(self, path):
        book = FakeBook(Path(path))
        self.app.books.append(book)
        return book


class FakeApp:
    """The user's Excel: one workbook open and active; a macro it is asked to run is recorded (none may be)."""

    def __init__(self, book: FakeBook):
        self.books, self.ActiveWorkbook, self.Ready, self.StatusBar = [book], book, True, False
        self.Workbooks = FakeWorkbooks(self)
        self.runs: list[str] = []

    def Run(self, macro, *args):
        self.runs.append(macro)
        return "OK"


class Reached:
    """Every way a job could start an Excel (a launch: the only one left), recorded and failed."""

    def __init__(self, monkeypatch):
        self.calls: list[str] = []
        monkeypatch.setattr(xltarget, "_launch_excel", self._fail("target._launch_excel"))
        monkeypatch.setattr(serverkit, "_users", None)

    def _fail(self, what):
        def f(*a, **k):
            self.calls.append(what)
            raise AssertionError(f"a tool call started an Excel: {what}")
        return f


@pytest.fixture
def excel(monkeypatch, tmp_path):
    """The user's Excel with the answer workbook open, and the jobs' threads recorded."""
    monkeypatch.setattr(config, "SVKIT", tmp_path / "svkit")
    config.SVKIT.mkdir()
    path = tmp_path / "Answer.xlsx"
    path.write_bytes(b"")
    app = FakeApp(FakeBook(path))
    monkeypatch.setattr(instances, "excel_instances", lambda *a, **k: [instances.XlInstance(PID, 1, [path.name], app)])
    reached = Reached(monkeypatch)
    threads: list[str] = []
    targets: list[str] = []

    real_target = exceljob._target

    def found(h, target, title, build_new=False):                # the job's workbook: the real resolve, recorded
        threads.append(threading.current_thread().name)
        t = real_target(h, target, title, build_new)
        targets.append(t.path)
        return t

    def excel_job(h, pre, target="session"):
        t = exceljob._target(h, target, "Q")
        return {"q": 1, "workbook": t.path, "sheets": ["Q1 Answer"]}

    def redraw_job(h, pre, q, target="session"):
        t = exceljob._target(h, target, "Follow-up")
        return {"q": q, "workbook": t.path, "closed": {}}
    monkeypatch.setattr(exceljob, "_target", found)
    stored = {1: {"spec": {"v2": spec_dict()}, "ids": {}, "owned": {"rows": 3, "cols": 2}}}
    monkeypatch.setattr(server, "excel_job", excel_job)
    monkeypatch.setattr(server, "redraw_job", redraw_job)
    monkeypatch.setattr(sheets, "_stored_specs", lambda t: stored)
    monkeypatch.setattr(server, "_provider", lambda: FakeProvider(facts()))
    monkeypatch.setattr(server, "run_prepare", lambda raw, prov, lines=None, **kw: raw)
    return type("Excel", (), {"app": app, "path": path, "reached": reached, "threads": threads,
                              "targets": targets})


def _ran_in_the_user_excel(excel, n: int, what: str) -> None:
    assert excel.reached.calls == []
    assert excel.threads == [f"user-job-{what}"] * n
    assert excel.targets == [str(excel.path)] * n
    assert excel.app.runs == []                                       # no macro: the bridge add-in is gone


def test_a_new_question_resolves_its_workbook_in_the_user_excel_without_the_hidden_one(excel):
    from smartass.model.spec import AnalysisSpec
    asyncio.run(server.fin_analyze(spec=AnalysisSpec.model_validate(spec_dict()), target="active"))
    _ran_in_the_user_excel(excel, 1, "fin_analyze")


def test_a_follow_up_reads_and_redraws_in_the_user_excel_without_the_hidden_one(excel):
    asyncio.run(server.fin_analyze(patch={"plan": {"purpose": "new"}}, target="active"))
    _ran_in_the_user_excel(excel, 2, "fin_analyze follow-up")


def test_a_refresh_reads_and_redraws_in_the_user_excel_without_the_hidden_one(excel):
    out = asyncio.run(server.fin_refresh(target="active"))
    assert [r["ok"] for r in out["refreshed"]] == [True]
    _ran_in_the_user_excel(excel, 2, "fin_refresh")


def test_a_drill_draws_in_the_user_excel_without_the_hidden_one(excel, monkeypatch):
    from types import SimpleNamespace

    async def drill(cell):
        return SimpleNamespace(rows=[], columns=["AMOUNT"], cell_value=0.0, reconciliation={}, warnings=[])
    monkeypatch.setattr(server, "_drill", drill)
    asyncio.run(server.fin_drill({"All Accounts": "ac.C60", "All Entities": "un.SY", "All Time Periods": "Sep 26"},
                                 target="active"))
    _ran_in_the_user_excel(excel, 1, "fin_analyze")


def test_the_critic_reads_its_answer_in_the_user_excel_without_the_hidden_one(excel, monkeypatch):
    from smartass import answers
    monkeypatch.setattr(answers, "get", lambda wb, q: {"question": "q", "critic_brief": "b", "score": None})
    out = asyncio.run(server.fin_critic(target="active"))
    assert out["q"] == 1
    _ran_in_the_user_excel(excel, 1, "fin_critic")


@pytest.mark.parametrize("tool, name, args", [("xl_snapshot", "snapshot", {}), ("xl_revert", "revert", {}),
                                              ("xl_tidy", "tidy", {})])
def test_the_workbench_tools_run_in_the_user_excel_without_the_hidden_one(excel, monkeypatch, tool, name, args):
    from smartass import workbench
    from smartass.xl import jobs
    ran = []

    def job(*a, **k):
        ran.append(threading.current_thread().name)
        return jobs.com(lambda: {"ok": True})
    monkeypatch.setattr(workbench, name, job)
    assert asyncio.run(getattr(server, tool)(**args)) == {"ok": True}
    assert ran == [f"user-job-{tool}"] and excel.reached.calls == []


def test_the_default_session_target_with_no_session_workbook_is_a_new_one_in_the_running_user_excel(excel, monkeypatch,
                                                                                                   tmp_path):
    """Task 5 review, Minor 3: the most common production path - target 'session' (the default) with no session
    workbook yet becomes 'new' in the running Excel; with a known base its file is reserved for the job to build
    (target._reserved_target). Nothing is launched and no Excel of the engine's own is reached."""
    from smartass.core import session
    from smartass.model.spec import AnalysisSpec
    monkeypatch.setattr(session, "workbook", lambda: None)
    monkeypatch.setattr(session, "project_dir", lambda: tmp_path)

    def excel_job(h, pre, target="session"):                 # as exceljob.excel_job when the base is known
        t = exceljob._target(h, target, "Q", build_new=True)
        return {"q": 1, "workbook": t.path, "sheets": ["Q1 Answer"]}
    monkeypatch.setattr(server, "excel_job", excel_job)
    asyncio.run(server.fin_analyze(spec=AnalysisSpec.model_validate(spec_dict())))
    assert excel.reached.calls == [] and excel.threads == ["user-job-fin_analyze"]
    assert len(excel.targets) == 1 and Path(excel.targets[0]).parent == tmp_path
    assert Path(excel.targets[0]).name.endswith(" Q.xlsx") and Path(excel.targets[0]).exists()


@pytest.mark.parametrize("restart", [False, True])
def test_fin_status_starts_no_excel_to_report(excel, monkeypatch, restart):
    from test_v2_server import _status_base
    _status_base(monkeypatch)
    monkeypatch.delenv("SVKIT_PROXY", raising=False)
    out = asyncio.run(server.fin_status(restart=restart))
    assert out["version"] == server.__version__ and excel.reached.calls == []


def test_the_bridge_add_in_is_gone_a_job_attaches_nothing_to_the_user_excel(excel):
    """Task 6 (flips Task 5's cold path (a)): nothing in the user's Excel called the bridge beyond its ping, so the
    attach went with it. A job finds its workbook and attaches, builds or opens nothing more - no hidden Excel, no
    add-in file."""
    from smartass.model.spec import AnalysisSpec
    asyncio.run(server.fin_analyze(spec=AnalysisSpec.model_validate(spec_dict()), target="active"))
    assert excel.reached.calls == [] and excel.app.runs == [] and [b.Name for b in excel.app.books] == ["Answer.xlsx"]
    assert not hasattr(instances, "ensure_bridge_file") and not hasattr(xltarget, "_attach_bridge")
    assert not hasattr(xltarget.Target, "run") and not list(config.SVKIT.glob("*.xlam"))


def test_a_new_workbook_with_no_excel_running_and_no_seed_is_built_and_opened_in_the_users_excel(excel, monkeypatch,
                                                                                               tmp_path):
    """Task 6 (flips Task 5's cold path (b)): no Excel is running and no labelled seed exists. With a stored base the
    first workbook is built without Excel (its Sources log, our label as custom document properties) and the user's
    Excel is launched on it (_launch_excel), as for any file; no hidden Excel is reached."""
    import zipfile

    from smartass.render.xlsx import Base
    monkeypatch.setattr(instances, "excel_instances", lambda *a, **k: [])
    monkeypatch.setattr(config, "NEW_WORKBOOK_SEED", tmp_path / "no-seed.xlsx")
    monkeypatch.setattr(config, "LABEL_ID", "f604")
    monkeypatch.setattr(config, "LABEL_SITE", "1676")
    monkeypatch.setattr(config, "HOME", tmp_path)
    monkeypatch.setattr(xltarget, "_new_path", lambda title, taken=(): tmp_path / f"{title}.xlsx")
    xltarget.remember_new_base(Base("Aptos Narrow", 11.0, 20, 64))
    launched = []

    def launch(path):
        launched.append(Path(path))
        raise HostError("launched the user's Excel")            # stop there: the rest waits for that Excel
    monkeypatch.setattr(xltarget, "_launch_excel", launch)               # the user's Excel, launched on the file
    from smartass.model.spec import AnalysisSpec
    from mcp.server.mcpserver.exceptions import ToolError
    with pytest.raises(ToolError, match="launched the user's Excel"):
        asyncio.run(server.fin_analyze(spec=AnalysisSpec.model_validate(spec_dict()), target="new"))
    assert launched == [tmp_path / "Q.xlsx"]
    with zipfile.ZipFile(launched[0]) as z:
        assert "MSIP_Label_f604_Enabled" in z.read("docProps/custom.xml").decode()
        assert "Sources" in z.read("xl/workbook.xml").decode()
    assert config.NEW_WORKBOOK_SEED.read_bytes() == launched[0].read_bytes()      # the seed for the next time
