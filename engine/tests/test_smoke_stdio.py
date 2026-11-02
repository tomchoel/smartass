from stdio_util import list_tools


def test_stdio_server_starts_and_lists_every_tool(tmp_path):
    names = {t["name"] for t in list_tools(tmp_path)}
    assert len(names) == 25, sorted(names)
    assert not {"sv_ws_pull", "sv_ws_build", "sv_ws_answer", "sv_analyse", "sv_concept"} & set(names)
    assert not {"fin_pnl", "sv_warehouse"} & set(names)
    assert not {"sv_query", "sv_scope", "sv_refresh_workbook"} & set(names)
    assert not any(n.startswith("sv_adhoc") or n == "sv_show" for n in names)
    assert {"sv_status", "sv_drill", "xl_read", "xl_revert", "sv_answer", "sv_gate",
            "sv_remember"} <= names
