import pytest
from mcp.server.mcpserver.exceptions import ToolError

from smartass.core import serverkit as server
from smartass.epm.drillcore import CellTooWide
from smartass.core.errors import HostTimeout


class HostError(Exception):
    pass


def test_explain_names_tool_cause_and_next_step():
    msg = server.explain("fin_analyze", HostError("Excel did not answer"))
    assert msg.startswith("fin_analyze failed - HostError: Excel did not answer")
    assert "Next step:" in msg and "fin_status(restart=True)" in msg
    assert "Next step:" in server.explain("fin_analyze", ValueError("bad"))


def test_the_restart_advice_says_what_a_restart_does_now():
    """Without the relay fin_status(restart=True) starts fresh EPM sessions and a fresh runner of Excel jobs; only
    behind the relay is the server itself replaced (Task 6 review, Minor 3)."""
    assert server.NEXT_STEP["HostError"].startswith(
        "Call fin_status(restart=True) (fresh EPM sessions and Excel-job runner; behind the relay, a fresh server), "
        "then retry; ")


def test_explain_uses_the_advice_of_a_base_class():
    assert "fin_status(restart=True)" in server.explain("fin_analyze", HostTimeout("slow"))
    assert "Narrow the cell" in server.explain("fin_drill", CellTooWide("too wide"))


def test_with_next_step_keeps_existing_advice():
    assert server.with_next_step("Bad member X. Next step: use fin_lookup.") == \
        "Bad member X. Next step: use fin_lookup."
    assert "Next step:" in server.with_next_step("Bad member X")
    assert server.with_next_step("Busy. Try again shortly.") == "Busy. Try again shortly."
    assert server.with_next_step("Failed. If it repeats, run fin_status(restart=True).") == "Failed. If it repeats, run fin_status(restart=True)."


@pytest.mark.anyio
async def test_guard_turns_any_exception_into_a_readable_tool_error():
    async def sv_fake():
        raise RuntimeError("sso_token=abcdefghijklmnopqrstuvwxyz leaked")  # allow-secret
    guarded = server.guard(sv_fake)
    with pytest.raises(ToolError) as e:
        await guarded()
    assert "sv_fake failed - RuntimeError" in str(e.value) and "abcdefghij" not in str(e.value)


@pytest.mark.anyio
async def test_guard_adds_next_step_to_tool_errors_and_runs_sync_functions():
    async def sv_bad():
        raise ToolError("Bad member X")

    def sv_sync(x: int) -> int:
        return x + 1

    with pytest.raises(ToolError, match="Bad member X. Next step:"):
        await server.guard(sv_bad)()
    assert await server.guard(sv_sync)(x=1) == 2


class _FailingHost:
    def __init__(self, exc):
        self.exc = exc

    async def call(self, fn, **kw):
        raise self.exc


async def _tool_error_via_host(monkeypatch, exc) -> str:
    monkeypatch.setattr(server, "_users", _FailingHost(exc))     # the runner of Excel jobs (the only one since Task 6)

    async def sv_fake():
        return await server._call_user(lambda h: None, what="fake")
    with pytest.raises(ToolError) as e:
        await server.guard(sv_fake)()
    return str(e.value)


@pytest.mark.anyio
async def test_host_error_via_call_gets_the_restart_hint_exactly_once(monkeypatch):
    from smartass.core.errors import HostError as RealHostError
    msg = await _tool_error_via_host(monkeypatch, RealHostError("COM error during query: boom"))
    assert msg.startswith("COM error during query: boom. Next step:")
    assert msg.count("fin_status(restart=True)") == 1 and msg.lower().count("next step") == 1


@pytest.mark.anyio
async def test_a_message_with_its_own_advice_gets_no_second_advice(monkeypatch):
    from smartass.core.errors import HostError as RealHostError
    own = "Unexpected KeyError during query: 'x'. If it repeats, run fin_status(restart=True)."
    assert await _tool_error_via_host(monkeypatch, RealHostError(own)) == own

    grid = "Grid has 9 cells (> 1). Narrow the members, split the query, or pass large=true if you really need it."

    async def fin_analyze():
        raise ToolError(grid)
    with pytest.raises(ToolError) as e:
        await server.guard(fin_analyze)()
    assert str(e.value) == grid


@pytest.mark.anyio
async def test_input_errors_never_suggest_a_restart(monkeypatch):
    from smartass.core.errors import InputError
    msg = await _tool_error_via_host(monkeypatch, InputError("Unknown pull sheet 'm'. Pulls: ['ytd']"))
    assert "fin_status(restart=True)" not in msg and "fin_status" not in msg
    assert msg.count("Next step:") == 1 and "Correct the input" in msg


def test_explain_gives_one_next_step_for_unknown_errors():
    msg = server.explain("fin_analyze", ValueError("bad"))
    assert msg.count("Next step:") == 1 and "Retry once" in msg


@pytest.mark.anyio
async def test_drill_errors_keep_their_class_advice():
    from smartass.epm.drillcore import DrillError

    async def fin_drill(cell=None):
        raise DrillError("getData returned HTTP 500 at offset 0.")
    with pytest.raises(ToolError) as e:
        await server.guard(fin_drill)(cell={})
    msg = str(e.value)
    assert "DrillError: getData returned HTTP 500" in msg and "Retry the drill once" in msg
    assert msg.count("Next step:") == 1


def test_a_missing_sign_in_becomes_the_sign_in_message(monkeypatch, tmp_path, caplog):          # Review Focus 4
    import asyncio
    import json

    from smartass.core import config
    from smartass.epm import oauth
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    monkeypatch.setattr(oauth, "sign_in_message", lambda: "EPM sign-in needed. Open https://idcs.example/ui and enter "
                        "the code ABCD (plan1.example, valid 5 min), approve with Telenor SSO, then ask again.")

    async def fin_analyze():
        raise oauth.NotSignedIn("not signed in to https://plan1.example: run 'smartass login'")
    with caplog.at_level("DEBUG"), pytest.raises(ToolError) as e:
        asyncio.run(server.guard(fin_analyze)())
    assert str(e.value).startswith("EPM sign-in needed.") and "Next step" not in str(e.value) and "code ABCD" in str(e.value)
    assert not [r for r in caplog.records if r.exc_info]                          # no traceback logged
    line = json.loads((tmp_path / "runs.jsonl").read_text(encoding="utf-8"))      # in the run log like any failure
    assert line["tool"] == "fin_analyze" and line["ok"] is False and "not signed in" in line["error"]
    assert "ABCD" not in line["error"]


def test_without_a_sign_in_message_the_error_itself_is_the_reply(monkeypatch):
    import asyncio

    from smartass.epm import oauth

    def broken():
        raise RuntimeError("status unreadable")
    for message in (lambda: None, broken):
        monkeypatch.setattr(oauth, "sign_in_message", message)

        async def fin_lookup():
            raise oauth.NotSignedIn("not signed in to https://plan1.example: run 'smartass login'")
        with pytest.raises(ToolError) as e:
            asyncio.run(server.guard(fin_lookup)())
        assert str(e.value) == "not signed in to https://plan1.example: run 'smartass login'"


def test_no_oauth_set_up_is_answered_with_the_set_up_text_and_no_retry_advice(monkeypatch, caplog):
    import asyncio

    from smartass.core import config
    from smartass.epm import oauth
    monkeypatch.setattr(config, "OAUTH_IDCS", "")
    monkeypatch.setattr(config, "OAUTH_CLIENT_ID", "")
    monkeypatch.setattr(config, "OAUTH_SCOPES", {})

    async def fin_drill():
        raise oauth.NotSetUp()
    with caplog.at_level("DEBUG"), pytest.raises(ToolError) as e:
        asyncio.run(server.guard(fin_drill)())
    assert str(e.value) == oauth.NOT_SET_UP and "Retry once" not in str(e.value)
    assert not [r for r in caplog.records if r.exc_info]
