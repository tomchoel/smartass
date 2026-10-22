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
