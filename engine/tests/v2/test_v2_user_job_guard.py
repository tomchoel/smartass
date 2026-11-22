"""Light-fast phase 2 Task 5: nothing on a tool call's path constructs, imports or starts the engine's hidden Excel.
Every tool that works in the user's Excel - fin_analyze (a new question and a follow-up), fin_drill, fin_refresh,
fin_critic, xl_snapshot / xl_revert / xl_tidy (the workbench) - runs its job through the real runner
(serverkit._call_user -> xl/userjobs.UserJobs, a thread per job) and resolves its workbook through the real
xl/target.resolve_target (the user's Excel found, its turn taken, the bridge add-in attached): the only code of a job
that uses the object the runner hands it (h). The Excel is a small fake; what a job draws after its workbook is found
is replaced (it never uses h). fin_status, with and without restart, starts nothing either. conftest makes a real
Excel start impossible; this pins that the code does not even try: HiddenExcel._ensure_started, ExcelHost._start,
ExcelHost() and serverkit.host() all fail the test.

Two paths still hand a job's work to the hidden Excel when they are needed (today's behaviour, kept until Task 6
replaces them): the bridge add-in's build when its file is missing, and a new analysis file when no Excel is running
and no labelled seed exists. Each is pinned below as a hand-off, so Task 6 has a test to flip."""
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
from smartass.epm import sso
from smartass.xl import hidden, host, instances, sheets
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
    """The user's Excel: one workbook open and active, the bridge add-in answering its ping."""

    def __init__(self, book: FakeBook):
        self.books, self.ActiveWorkbook, self.Ready, self.StatusBar = [book], book, True, False
        self.Workbooks = FakeWorkbooks(self)
        self.runs: list[str] = []

    def Run(self, macro, *args):
        self.runs.append(macro)
        return "pong" if macro.endswith("B_Ping") else "OK"


class Reached:
    """Every way into the hidden Excel, each recorded and failed."""

    def __init__(self, monkeypatch):
        self.calls: list[str] = []
        for owner, name in ((hidden.HiddenExcel, "_ensure_started"), (host.ExcelHost, "_start"),
                            (host.ExcelHost, "__init__")):
            monkeypatch.setattr(owner, name, self._fail(f"{owner.__name__}.{name}"))
        monkeypatch.setattr(serverkit, "host", self._fail("serverkit.host()"))
        monkeypatch.setattr(serverkit, "_host", None)
        monkeypatch.setattr(serverkit, "_users", None)

    def _fail(self, what):
        def f(*a, **k):
            self.calls.append(what)
            raise AssertionError(f"a tool call reached the hidden Excel: {what}")
        return f


@pytest.fixture
def excel(monkeypatch, tmp_path):
    """The user's Excel with the answer workbook open, the bridge add-in built for the current sources, and the jobs'
    threads recorded."""
    monkeypatch.setattr(config, "SVKIT", tmp_path / "svkit")
    config.SVKIT.mkdir()
    bridge = instances.bridge_path(instances._bridge_hash())
    bridge.write_bytes(b"")                                       # built: its build needs no hidden Excel
    path = tmp_path / "Answer.xlsx"
    path.write_bytes(b"")
    app = FakeApp(FakeBook(path))
    app.books.append(FakeBook(bridge))
    monkeypatch.setattr(instances, "excel_instances", lambda *a, **k: [instances.XlInstance(PID, 1, [path.name], app)])
    monkeypatch.setattr(sso, "ensure_global_watcher", lambda: None)
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
    return type("Excel", (), {"app": app, "path": path, "bridge": bridge, "reached": reached, "threads": threads,
                              "targets": targets})


def _ran_in_the_user_excel(excel, n: int, what: str) -> None:
    assert excel.reached.calls == [] and serverkit._host is None
    assert excel.threads == [f"user-job-{what}"] * n
    assert excel.targets == [str(excel.path)] * n
    assert any(r.endswith("B_Ping") for r in excel.app.runs)           # the bridge attached, from its file


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
    assert ran == [f"user-job-{tool}"] and excel.reached.calls == [] and serverkit._host is None


@pytest.mark.parametrize("restart", [False, True])
def test_fin_status_starts_no_excel_to_report(excel, monkeypatch, restart):
    from test_v2_server import _status_base
    _status_base(monkeypatch)
    monkeypatch.delenv("SVKIT_PROXY", raising=False)
    out = asyncio.run(server.fin_status(restart=restart))
    assert out["version"] == server.__version__ and excel.reached.calls == [] and serverkit._host is None


class FakeHidden:
    """The hidden Excel as the runner reaches it (serverkit.host): its work recorded, no Excel."""

    def __init__(self, then):
        self.app = object()
        self.on_host_calls: list[str] = []
        self.then = then

    def on_host(self, fn, timeout=None):
        self.on_host_calls.append(threading.current_thread().name)
        return self.then(fn)


def test_today_a_missing_bridge_file_is_built_by_the_hidden_excel(excel, monkeypatch):
    """Cold path (a), kept until Task 6: the bridge add-in for the current sources is not built yet, so the user job
    hands its build to the hidden Excel (instances.ensure_bridge_file -> h.on_host -> _build_bridge in that Excel's
    thread, h.app its Application). Task 6 flips this test when the attach goes."""
    excel.bridge.unlink()

    def build(fn):                                               # the hidden Excel builds it, then the build returns
        excel.bridge.write_bytes(b"")
        return fn()
    hid = FakeHidden(build)
    monkeypatch.setattr(serverkit, "host", lambda: hid)
    from smartass.model.spec import AnalysisSpec
    asyncio.run(server.fin_analyze(spec=AnalysisSpec.model_validate(spec_dict()), target="active"))
    assert hid.on_host_calls == ["user-job-fin_analyze"] and excel.reached.calls == []
    assert serverkit.user_jobs().app is hid.app


def test_today_a_new_workbook_with_no_excel_running_and_no_seed_is_made_by_the_hidden_excel(excel, monkeypatch,
                                                                                         tmp_path):
    """Cold path (b), kept until Task 6: no Excel is running and no labelled seed exists, so the user job hands the
    new analysis file to the hidden Excel (target._new_analysis_file -> h.on_host -> _new_analysis_file_host). Task 6
    flips this test when a first workbook can be made without it."""
    monkeypatch.setattr(instances, "excel_instances", lambda *a, **k: [])
    monkeypatch.setattr(config, "NEW_WORKBOOK_SEED", tmp_path / "no-seed.xlsx")
    monkeypatch.setattr(xltarget, "_new_path", lambda title, taken=(): tmp_path / f"{title}.xlsx")

    def made(fn):
        raise HostError("handed to the hidden Excel")            # stop there: the rest would start an Excel
    hid = FakeHidden(made)
    monkeypatch.setattr(serverkit, "host", lambda: hid)
    from smartass.model.spec import AnalysisSpec
    from mcp.server.mcpserver.exceptions import ToolError
    with pytest.raises(ToolError, match="handed to the hidden Excel"):
        asyncio.run(server.fin_analyze(spec=AnalysisSpec.model_validate(spec_dict()), target="new"))
    assert hid.on_host_calls == ["user-job-fin_analyze"] and excel.reached.calls == []
