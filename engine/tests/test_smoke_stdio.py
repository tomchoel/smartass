"""The real stdio server (child_server.py, as run_server.py starts it): named smartass, exactly the v2 tools."""
from stdio_util import initialize_and_list_tools

TOOLS = {"fin_analyze", "fin_lookup", "fin_status", "fin_refresh", "fin_drill", "fin_critic", "xl_snapshot", "xl_revert",
         "xl_tidy"}


def test_stdio_server_is_smartass_with_the_v2_tools(tmp_path):
    init, tools = initialize_and_list_tools(tmp_path / "home")
    assert init["serverInfo"]["name"] == "smartass"
    by = {t["name"]: t for t in tools}
    assert set(by) == TOOLS
    assert "restart" in by["fin_status"]["inputSchema"]["properties"]
    assert {"spec", "patch", "q", "target"} <= set(by["fin_analyze"]["inputSchema"]["properties"])
    assert {"q", "scores", "notes"} <= set(by["fin_critic"]["inputSchema"]["properties"])
    assert not (tmp_path / "home" / "pids.json").exists()          # no Excel was started
