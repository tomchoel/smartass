from stdio_util import list_tools


def test_stdio_server_starts_and_lists_every_tool(tmp_path):
    names = {t["name"] for t in list_tools(tmp_path)}
    assert len(names) >= 36
    assert {"sv_status", "sv_query", "sv_ws_build", "sv_drill", "xl_read", "xl_revert"} <= names
