# engine/tests/v2/test_v2_server.py
import pytest

from smartass import server


def test_nine_tools_registered():
    assert server.TOOL_NAMES == ("fin_analyze", "fin_lookup", "fin_status", "fin_refresh", "fin_drill", "fin_critic",
                                 "xl_snapshot", "xl_revert", "xl_tidy")


def test_instructions_mention_the_excel_toolkit_and_reconnect():
    text = server.INSTRUCTIONS
    assert "xl_snapshot" in text and "Excel MCP" in text and "/mcp" in text
    assert len(text) < 2500


def test_lookup_metrics_and_members(tmp_path, monkeypatch):
    from kutil import knowledge
    knowledge(tmp_path, monkeypatch)
    out = server.lookup("opex", kind="any")
    assert out["metrics"][0]["name"] == "Opex"
    assert out["metrics"][0]["members"]["FinMax"] == ["ac.C50S", "ac.C60"]
    assert server.lookup("2BUS", kind="any")["members"][0]["name"] == "un.SY.2BUS"


def test_lookup_kind_metric_only(tmp_path, monkeypatch):
    from kutil import knowledge
    from smartass.knowledge import atlas
    knowledge(tmp_path, monkeypatch)
    monkeypatch.setattr(atlas, "search", lambda *a, **k: pytest.fail("members not asked"))
    assert server.lookup("ebitda", kind="metric")["members"] == []


def test_instructions_are_short():
    assert len(server.INSTRUCTIONS) < 2500


def test_patched_spec_is_validated_before_any_excel_work():
    from v2util import spec_dict

    from smartass.followup import merge_patch
    from smartass.model.spec import AnalysisSpec
    raw = merge_patch(spec_dict(), {"plan": {"purpose": "New purpose"}})
    assert AnalysisSpec.model_validate(raw).plan.purpose == "New purpose"


def test_instructions_mention_followup_refresh_and_bridge_steps():
    text = server.INSTRUCTIONS
    assert "patch" in text and "fin_refresh" in text and "COGS" in text
    assert len(text) < 2500


def _wire(monkeypatch, stored, fail_prepare=()):
    import asyncio

    async def user(fn, what):
        return fn(None)
    calls = []
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "_provider", lambda: "provider")
    monkeypatch.setattr(server, "read_stored", lambda h, target="session": (stored, "Q2 Answer", "C:/a/x.xlsx"))

    def prep(raw, provider, lines=None):
        if raw["question"] in fail_prepare:
            raise RuntimeError("no data")
        calls.append(("prepare", raw["question"]))
        return raw
    monkeypatch.setattr(server, "run_prepare", prep)

    def redraw(h, pre, q, owned, target="session"):
        calls.append(("redraw", q))
        assert target == "C:/a/x.xlsx"                      # the path read with the specs, not the word 'session'
        assert pre is not None
        return {"q": q, "workbook": "W", "closed": {"FinMax": "09/2026"}}
    monkeypatch.setattr(server, "redraw_job", redraw)
    return calls, asyncio


def test_fin_analyze_needs_exactly_one_of_spec_and_patch(monkeypatch):
    import asyncio

    from smartass.model.spec import AnalysisSpec
    from smartass.model.validate import GuardError
    from v2util import spec_dict
    with pytest.raises(GuardError, match="not both"):
        asyncio.run(server.fin_analyze())
    with pytest.raises(GuardError, match="not both"):
        asyncio.run(server.fin_analyze(spec=AnalysisSpec.model_validate(spec_dict()), patch={"x": 1}))


def test_followup_patches_the_stored_spec_and_redraws(monkeypatch):
    stored = {2: {"spec": {"v2": {"question": "a", "plan": {"purpose": "old"}}}, "ids": {}, "owned": {"rows": 3, "cols": 2}}}
    calls, asyncio = _wire(monkeypatch, stored)
    seen = {}
    monkeypatch.setattr(server, "run_prepare", lambda raw, provider, lines=None, **kw: seen.update(raw=raw) or raw)
    out = asyncio.run(server.fin_analyze(patch={"plan": {"purpose": "new"}}))
    assert seen["raw"]["plan"]["purpose"] == "new" and out["q"] == 2 and calls == [("redraw", 2)]


def test_refresh_continues_past_a_failing_question(monkeypatch):
    stored = {2: {"spec": {"v2": {"question": "bad"}}, "ids": {}, "owned": None},
              3: {"spec": {"v2": {"question": "good"}}, "ids": {}, "owned": None}}
    calls, asyncio = _wire(monkeypatch, stored, fail_prepare=("bad",))
    out = asyncio.run(server.fin_refresh())
    r = out["refreshed"]
    assert [(x["q"], x["ok"]) for x in r] == [(2, False), (3, True)]
    assert "no data" in r[0]["error"] and r[1]["closed"] == {"FinMax": "09/2026"}
    assert calls == [("prepare", "good"), ("redraw", 3)]


def test_refresh_stops_after_a_failed_restore_and_reports_the_rest_as_skipped(monkeypatch):
    from smartass.pipeline import RestoreError
    stored = {q: {"spec": {"v2": {"question": f"q{q}"}}, "ids": {}, "owned": None} for q in (2, 3, 4)}
    calls, asyncio = _wire(monkeypatch, stored)

    def redraw(h, pre, q, owned, target="session"):
        calls.append(("redraw", q))
        if q == 3:
            raise RestoreError("boom; the restore was incomplete - restore Q3 Answer: x")
        return {"q": q, "workbook": "W", "closed": {}}
    monkeypatch.setattr(server, "redraw_job", redraw)
    out = asyncio.run(server.fin_refresh())
    r = out["refreshed"]
    assert [(x["q"], x["ok"]) for x in r] == [(2, True), (3, False), (4, False)]
    assert "restore" in r[1]["error"] and r[2].get("skipped") is True and "Q3" in r[2]["error"]
    assert ("redraw", 4) not in calls


def _status_base(monkeypatch):
    monkeypatch.setattr(server.atlas, "connect", lambda: _FakeDb())
    monkeypatch.setattr(server.atlas, "status", lambda db: {"stale": False, "oldest_age_days": 3.0, "built": [1]})
    from smartass.knowledge import build
    monkeypatch.setattr(build, "summary", lambda path: {"metrics": 2, "facts": 5, "drafts": 0, "built_at": "x"})
    from smartass.ops import doctor
    monkeypatch.setattr(doctor, "run", lambda: [doctor.Check("excel", "OK", "x"),
                                                doctor.Check("vba_access", "FAIL", "blocked", "tick it")])


class _FakeDb:
    def close(self):
        pass


def test_fin_status_restart_behind_the_relay_reports_the_fresh_child(monkeypatch):
    import asyncio
    _status_base(monkeypatch)
    monkeypatch.setenv("SVKIT_PROXY", "1")
    out = asyncio.run(server.fin_status(restart=True))
    assert out["restarted"] == "relay"
    assert out["version"] == server.__version__


def test_fin_status_restart_without_relay_restarts_our_excel(monkeypatch):
    import asyncio
    _status_base(monkeypatch)
    monkeypatch.delenv("SVKIT_PROXY", raising=False)
    closed = []

    class Old:
        def shutdown(self):
            closed.append(True)
    monkeypatch.setattr(server.kit, "_host", Old())
    out = asyncio.run(server.fin_status(restart=True))
    assert closed == [True] and server.kit._host is None
    assert out["restarted"] == "excel"


def test_fin_status_without_restart_changes_nothing(monkeypatch):
    import asyncio
    _status_base(monkeypatch)
    out = asyncio.run(server.fin_status())
    assert "restarted" not in out and out["metrics"] == 2
    assert out["atlas"] == {"oldest_age_days": 3.0, "stale": False, "dimensions_built": 1}
    assert out["doctor"]["overall"] == "FAIL"
    assert [p["name"] for p in out["doctor"]["problems"]] == ["vba_access"]


def test_instructions_mention_the_critic_and_the_prompts():
    text = server.INSTRUCTIONS
    assert "fin_critic" in text and "prompts" in text and "kind='answer'" in text
    assert len(text) < 2500
