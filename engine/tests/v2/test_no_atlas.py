"""A PC with no atlas (light-fast phase 4 Task 2, the controller's ruling of 22.11): every fin_* entry that needs the
atlas says ONE command - `smartass login` (it signs in and builds the atlas) while a sign-in is missing, else
`smartass atlas build`. "No atlas" is a missing file or one without members (atlas.connect() makes such a file on a
fresh home: fin_status and `smartass atlas status` call it); it is read without creating anything. An atlas that has
members but lacks one a metric needs keeps today's text (a semantic-file problem). fin_critic (answers.db), fin_status
(it reports the atlas) and the xl_* tools (open workbooks) do not need the atlas. EPM, Excel and the sign-ins are
fakes."""
import asyncio
from types import SimpleNamespace

import pytest
from kutil import knowledge
from v2util import spec_dict

from smartass import server
from smartass.core import config, serverkit
from smartass.epm import oauth
from smartass.knowledge import atlas
from smartass.model.spec import AnalysisSpec
from smartass.ops import atlas_build, doctor

LOGIN, BUILD = "`smartass login`", "`smartass atlas build`"
CELL = {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"}


@pytest.fixture(params=["no file", "empty file"])
def no_atlas(request, monkeypatch, tmp_path):
    """A home with no atlas (missing, or the schema alone) and no knowledge.db; the tools' EPM and Excel are fakes
    that fail the test when they are reached."""
    a = tmp_path / "atlas.db"
    if request.param == "empty file":
        atlas.connect(a).close()
    monkeypatch.setattr(config, "ATLAS_DB", a)
    monkeypatch.setattr(config, "KNOWLEDGE_DB", tmp_path / "knowledge.db")
    monkeypatch.setattr(atlas.connect, "__defaults__", (a,))

    def unreachable(*args, **kwargs):
        pytest.fail("EPM or Excel was reached without an atlas")

    async def user(fn, what):
        return fn(None)
    stored = {2: {"spec": {"v2": spec_dict()}, "ids": {}, "owned": None}}
    from smartass.epm import drillstep
    closed = {"FinMax": (2026, 9), "FinLite": (2026, 9), "FinPlan": (2026, 9)}
    monkeypatch.setattr(server, "_provider", lambda: SimpleNamespace(closed_months=lambda: closed, fetch=unreachable))
    monkeypatch.setattr(drillstep, "drill", unreachable)
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "read_stored", lambda h, target="session": (stored, "Q2 Answer", "C:/a/x.xlsx"))
    monkeypatch.setattr(server, "_default_target", lambda target: target)
    return a


def signed_in(monkeypatch, state):
    monkeypatch.setattr(oauth, "configured", lambda: state is not None)
    monkeypatch.setattr(oauth, "status", lambda: [{"pod": "https://plan1.example", "state": state or "none"}])


def _said(call) -> str:
    try:
        out = asyncio.run(call())
    except Exception as e:   # noqa: BLE001 - the tool's error text is what the agent reads
        return str(e)
    return " ".join(r.get("error", "") for r in out["refreshed"])


ENTRIES = {
    "fin_lookup": lambda: serverkit.guard(server.fin_lookup)("opex"),
    "fin_lookup member": lambda: serverkit.guard(server.fin_lookup)("2BUS", kind="member"),
    "fin_analyze spec": lambda: serverkit.guard(server.fin_analyze)(spec=AnalysisSpec.model_validate(spec_dict())),
    "fin_analyze patch": lambda: serverkit.guard(server.fin_analyze)(patch={"plan": {"purpose": "new"}}),
    "fin_drill": lambda: serverkit.guard(server.fin_drill)(CELL),
    "fin_refresh": lambda: serverkit.guard(server.fin_refresh)(),
}


@pytest.mark.parametrize("entry", ENTRIES)
@pytest.mark.parametrize("state, command, other", [(None, LOGIN, BUILD), ("ended", LOGIN, BUILD),
                                                   ("signed_in", BUILD, LOGIN)])
def test_each_entry_that_needs_the_atlas_says_one_command(no_atlas, monkeypatch, entry, state, command, other):
    signed_in(monkeypatch, state)
    said = _said(ENTRIES[entry])
    assert "no atlas" in said and command in said, said
    assert other not in said and "smartass knowledge build" not in said, said
    assert not no_atlas.exists() or no_atlas.stat().st_size > 0        # nothing was created for a missing file


def test_past_answers_need_no_atlas(no_atlas, monkeypatch):
    from smartass import answers
    monkeypatch.setattr(answers, "search", lambda query, limit: [])
    assert asyncio.run(serverkit.guard(server.fin_lookup)("opex", kind="answer")) == {"answers": []}


def test_an_atlas_that_lacks_a_metrics_member_keeps_todays_text(tmp_path, monkeypatch):
    from kutil import METRICS
    _, sem, k = knowledge(tmp_path, monkeypatch, build_db=False)
    (sem / "metrics" / "ghost.toml").write_text(METRICS["ebitda"].replace("EBITDA", "Ghost").replace(
        "ac.Ghost", "ac.NOPE"), encoding="utf-8")
    signed_in(monkeypatch, "signed_in")
    said = _said(lambda: serverkit.guard(server.fin_lookup)("opex"))
    assert "member 'ac.NOPE' is not in the atlas" in said and "smartass knowledge build" in said
    assert "no atlas" not in said


def test_no_atlas_is_read_without_creating_a_file(tmp_path):
    missing, empty, junk = tmp_path / "missing.db", tmp_path / "empty.db", tmp_path / "junk.db"
    atlas.connect(empty).close()
    junk.write_bytes(b"not a database")
    kept, _, _ = knowledge(tmp_path / "k")
    assert [atlas.present(p) for p in (missing, empty, kept)] == [False, False, True]
    assert atlas.present(junk) is True                     # there, but unreadable: its own error, never built over
    assert not missing.exists()


@pytest.mark.parametrize("state, command", [(None, LOGIN), ("none", LOGIN), ("ended", LOGIN), ("signed_in", BUILD)])
def test_the_doctor_says_no_atlas_with_its_one_command_and_knowledge_waits_for_it(no_atlas, monkeypatch, state,
                                                                                   command):
    signed_in(monkeypatch, state)
    c = doctor._atlas()
    assert c.status == "FAIL" and str(no_atlas) in c.detail and command in c.fix
    assert (BUILD if command == LOGIN else LOGIN) not in c.fix and "install" not in c.fix
    k = doctor._knowledge()
    assert k.status == "WARN" and "waits for the atlas" in k.detail and command in k.fix
    assert "semantic" not in k.fix
    assert atlas_build.no_atlas_command() == command.strip("`")
