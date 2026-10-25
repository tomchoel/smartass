from stdio_util import list_tools


def test_stdio_server_starts_and_lists_every_tool(tmp_path):
    names = {t["name"] for t in list_tools(tmp_path)}
    assert len(names) == 32, sorted(names)
    assert not any(n.startswith("sv_adhoc") or n == "sv_show" for n in names)
    assert {"sv_status", "sv_query", "sv_ws_build", "sv_drill", "xl_read", "xl_revert", "fin_pnl",
            "sv_warehouse"} <= names
