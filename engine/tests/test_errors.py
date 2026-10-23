import pytest
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import server
from smartview_mcp.drill_browser import CellTooWide
from smartview_mcp.host import HostTimeout


class HostError(Exception):
    pass


def test_explain_names_tool_cause_and_next_step():
    msg = server.explain("sv_query", HostError("Excel did not answer"))
    assert msg.startswith("sv_query failed - HostError: Excel did not answer")
    assert "Next step:" in msg and "sv_restart" in msg
    assert "Next step:" in server.explain("sv_query", ValueError("bad"))


def test_explain_uses_the_advice_of_a_base_class():
    assert "sv_restart" in server.explain("sv_query", HostTimeout("slow"))
    assert "Narrow the cell" in server.explain("sv_drill", CellTooWide("too wide"))


def test_with_next_step_keeps_existing_advice():
    assert server.with_next_step("Bad member X. Next step: use sv_find_member.") == \
        "Bad member X. Next step: use sv_find_member."
    assert "Next step:" in server.with_next_step("Bad member X")
    assert server.with_next_step("Busy. Try again shortly.") == "Busy. Try again shortly."
    assert server.with_next_step("Failed. If it repeats, run sv_restart.") == "Failed. If it repeats, run sv_restart."


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


async def _tool_error_via_host(monkeypatch, exc, user_excel=False) -> str:
    monkeypatch.setattr(server, "_host", _FailingHost(exc))

    async def sv_fake():
        if user_excel:
            return await server._call_user(lambda h: None, what="fake")
        return await server._call(lambda h: None, what="fake")
    with pytest.raises(ToolError) as e:
        await server.guard(sv_fake)()
    return str(e.value)


@pytest.mark.anyio
async def test_host_error_via_call_gets_the_restart_hint_exactly_once(monkeypatch):
    from smartview_mcp.host import HostError as RealHostError
    msg = await _tool_error_via_host(monkeypatch, RealHostError("COM error during query: boom"))
    assert msg.startswith("COM error during query: boom. Next step:")
    assert msg.count("sv_restart") == 1 and msg.lower().count("next step") == 1


@pytest.mark.anyio
async def test_a_message_with_its_own_advice_gets_no_second_advice(monkeypatch):
    from smartview_mcp.host import HostError as RealHostError
    own = "Unexpected KeyError during query: 'x'. If it repeats, run sv_restart."
    assert await _tool_error_via_host(monkeypatch, RealHostError(own)) == own

    grid = "Grid has 9 cells (> 1). Narrow the members, split the query, or pass large=true if you really need it."

    async def sv_query():
        raise ToolError(grid)
    with pytest.raises(ToolError) as e:
        await server.guard(sv_query)()
    assert str(e.value) == grid


@pytest.mark.anyio
async def test_input_errors_never_suggest_a_restart(monkeypatch):
    from smartview_mcp.host import InputError
    msg = await _tool_error_via_host(monkeypatch, InputError("Unknown pull sheet 'm'. Pulls: ['ytd']"), user_excel=True)
    assert "sv_restart" not in msg and "sv_status" not in msg
    assert msg.count("Next step:") == 1 and "Correct the input" in msg


def test_explain_gives_one_next_step_for_unknown_errors():
    msg = server.explain("sv_query", ValueError("bad"))
    assert msg.count("Next step:") == 1 and "Retry once" in msg


@pytest.mark.anyio
async def test_drill_errors_keep_their_class_advice(monkeypatch):
    from smartview_mcp import tools_drill
    from smartview_mcp.drill_browser import DrillError

    async def boom(*a, **k):
        raise DrillError("getData returned HTTP 500 at offset 0.")
    monkeypatch.setattr(tools_drill, "_sv_drill", boom)
    with pytest.raises(ToolError) as e:
        await tools_drill.sv_drill(question="q", cell={})
    msg = str(e.value)
    assert "DrillError: getData returned HTTP 500" in msg and "Retry the drill once" in msg
    assert msg.count("Next step:") == 1


def test_answer_spec_errors_are_input_errors_not_host_errors():
    from smartview_mcp import answer
    from smartview_mcp.host import HostError as RealHostError
    from smartview_mcp.host import InputError
    with pytest.raises(InputError) as e:
        answer._expr_to_formula("rev + nope", "D", {"rev": 5})
    assert not isinstance(e.value, RealHostError)


def test_rewrap_keeps_input_errors_as_input_errors():
    from smartview_mcp.host import HostError as RealHostError
    from smartview_mcp.host import InputError, rewrap
    assert type(rewrap(InputError("bad spec"), "bad spec (removed: none)")) is InputError
    assert type(rewrap(RuntimeError("boom"), "boom (removed: none)")) is RealHostError


def test_connect_without_a_session_id_says_sign_in_and_retry():
    from smartview_mcp.host import _explain
    msg = _explain("B_Connect", "ERR|-15|invalid parameter|Argument sID in the request is an empty string or NULL. "
                                "Try to connect again.")
    assert "not signed in" in msg and "retry" in msg


def test_login_dialog_message_asks_the_user_not_the_agent(monkeypatch):
    import pytest

    from smartview_mcp import userexcel
    from smartview_mcp.host import HostError
    monkeypatch.setattr(userexcel, "login_open", lambda pid: True)
    monkeypatch.setattr(userexcel.time, "sleep", lambda s: None)
    with pytest.raises(HostError) as e:
        userexcel.wait_login_cleared(1, limit=0.01)
    assert "Ask the user" in str(e.value) and "PowerPoint" in str(e.value)
