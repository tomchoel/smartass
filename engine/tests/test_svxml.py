import pytest
from svxml_util import answer

from smartview_mcp import svxml
from smartview_mcp.grid import Axis, Layout, parse


def _lay():
    return Layout([("All Scenarios", "Actual")], [Axis("All Accounts", ["ac.A", "ac.B"])],
                  [Axis("All Time Periods", ["Aug 26", "Sep 26"])])


def test_grid_cells_mark_pov_headers_and_data():
    vals, types = svxml.grid_cells(_lay())
    assert vals == ["", "Actual", "", "", "Aug 26", "Sep 26", "ac.A", "", "", "ac.B", "", ""]
    assert types == ["7", "0", "7", "7", "0", "0", "0", "2", "2", "0", "2", "2"]


def test_grid_cells_escape_xml_and_refuse_pipes():
    lay = Layout([("All Entities", "un.SY.5P&C_P&C")], [Axis("All Accounts", ["ac.A"])],
                 [Axis("All Time Periods", ["Sep 26"])])
    assert "un.SY.5P&amp;C_P&amp;C" in svxml.grid_cells(lay)[0]
    bad = Layout([], [Axis("All Accounts", ["a|b"])], [Axis("All Time Periods", ["Sep 26"])])
    with pytest.raises(ValueError, match="separator"):
        svxml.grid_cells(bad)


def test_execute_grid_is_a_retrieve_with_default_aliases():
    body = svxml.execute_grid("S1", "FinMax", _lay())
    svxml.check_read_only(body)
    assert body.startswith("<req_ExecuteGrid><sID>S1</sID><action>retrieve</action>")
    assert '<aliasTableName val="Default"/>' in body and "<cube>FinMax</cube>" in body
    assert '<slice rows="4" cols="3">' in body and '<range start="0" end="11">' in body


@pytest.mark.parametrize("body", [
    "<req_WriteBack><sID>x</sID></req_WriteBack>",
    "<req_SetVals/>",
    "<req_ExecuteGrid><sID>x</sID><action>freeformsubmit</action></req_ExecuteGrid>",
    "<req_ExecuteGrid><sID>x</sID></req_ExecuteGrid>",
    "not xml",
    "<req_ExecuteGrid><action>retrieve</action><action>\nfreeformsubmit\n</action></req_ExecuteGrid>",
    "<req_ExecuteGrid><action>retrieve</action><action >submit</action></req_ExecuteGrid>",
    "<req_OpenCube></req_OpenCube><req_WriteBack/>",
])
def test_check_read_only_refuses_everything_else(body):
    with pytest.raises(ValueError, match="refused"):
        svxml.check_read_only(body)


def test_session_requests_pass_the_allowlist():
    for body in (svxml.connect(), svxml.provisioned_data_sources(), svxml.open_application("s", "FinMax"),
                 svxml.open_cube("s", "FinMax", "FinLite")):
        svxml.check_read_only(body)


def test_sid_and_exceptions():
    assert svxml.sid_of("<res_OpenApplication><sID>-4f:1a</sID></res_OpenApplication>") == "-4f:1a"
    with pytest.raises(svxml.SvXmlError) as e:
        svxml.sid_of('<exception errcode="1000" type="error"><desc>Grid is &lt;bad&gt;</desc></exception>')
    assert e.value.code == 1000 and str(e.value) == "Grid is <bad>"
    with pytest.raises(svxml.SvXmlError, match="not a Smart View provider answer"):
        svxml.sid_of("<html>401 Authorization Required</html>")


def test_data_sources_give_the_pod_per_app():
    text = ('<res_GetProvisionedDataSources><Product id="HP"><Server name="FinMax" '
            'context="https://plan1.example/HyperionPlanning/SmartView" appName="FinMax"/></Product>'
            '<Product id="HP"><Server name="FinPlan_for_DataLoad" '
            'context="https://strat.example/HyperionPlanning/SmartView" appName="FinPlan"/></Product>'
            '</res_GetProvisionedDataSources>')
    assert svxml.data_sources(text) == {"FinMax": "https://plan1.example", "FinPlan": "https://strat.example"}


def test_grid_values_round_trip_through_grid_parse():
    m = [[None, "Actual", None], [None, "Aug 26", "Sep 26"],
         ["A account (ac)", 100.5, 110.0], ["B account (ac)", "#noaccess", None]]
    values = svxml.grid_values(answer(m))
    assert values[2] == ["A account (ac)", 100.5, 110.0] and values[3][1] == "#NoAccess"
    p = parse(_lay(), values, suppress_missing=True, suppress_zeros=False)
    assert p.rows == [["ac.A", 100.5, 110.0], ["ac.B", "#NoAccess", None]]
    assert p.row_labels == [["A account (ac)"], ["B account (ac)"]] and p.no_access_cells == 1


def test_grid_values_raise_on_provider_exception():
    with pytest.raises(svxml.SvXmlError, match="at least 2 members"):
        svxml.grid_values('<exception errcode="1000" type="error"><desc>Grid is Invalid. It should have at '
                          'least 2 members valid for selected plan type.</desc></exception>')


def test_hostile_member_name_cannot_smuggle_an_action():
    lay = Layout([("All Entities", "x<action>submit</action>")], [Axis("All Accounts", ["ac.A"])],
                 [Axis("All Time Periods", ["Sep 26"])])
    body = svxml.execute_grid("S1", "FinMax", lay)
    svxml.check_read_only(body)
    assert "<action>submit" not in body


def test_unlisted_requests_are_not_allowed():
    assert "ListCubes" not in svxml.ALLOWED and "GetDefaultGrid" not in svxml.ALLOWED


def test_raise_for_exception_ignores_attribute_order():
    with pytest.raises(svxml.SvXmlError) as e:
        svxml.raise_for_exception('<res_X><exception type="error" errcode="5"><desc>bad</desc></exception></res_X>')
    assert e.value.code == 5 and str(e.value) == "bad"


def test_sid_is_unescaped_and_stripped():
    assert svxml.sid_of("<res_X><sID> a&amp;b </sID></res_X>") == "a&b"


def test_sid_in_is_none_when_absent_and_never_raises():
    assert svxml.sid_in("<res_X><sID> a&amp;b </sID></res_X>") == "a&b"
    assert svxml.sid_in("<res_OpenCube></res_OpenCube>") is None
    assert svxml.sid_in("<res_X><sID>  </sID></res_X>") is None
    assert svxml.sid_in('<exception errcode="1000"><desc>x</desc></exception>') is None


def test_execute_grid_escapes_the_alias_attribute():
    body = svxml.execute_grid("S1", "FinMax", _lay(), alias='a"b')
    assert 'val="a&quot;b"' in body


def test_data_sources_ignore_attribute_order_and_skip_incomplete():
    text = ('<res_X><Server appName="FinMax" name="n" context="https://p.example/HyperionPlanning/SmartView"/>'
            '<Server name="x" appName="NoContext"/><Server name="y" context="https://q.example/HyperionPlanning/x"/>'
            '</res_X>')
    assert svxml.data_sources(text) == {"FinMax": "https://p.example"}


def _ans(rows, cols, start, vals, types):
    end = start + len(vals) - 1
    return (f'<res_ExecuteGrid><slice rows="{rows}" cols="{cols}"><data><range start="{start}" end="{end}">'
            f'<vals enc="1">{"|".join(vals)}</vals><types enc="0">{"|".join(types)}</types></range></data></slice>'
            '</res_ExecuteGrid>')


def test_grid_values_refuse_a_pipe_in_a_returned_label():
    text = _ans(1, 2, 0, ["a", "b", "c"], ["0", "0"]).replace('end="2"', 'end="1"')
    with pytest.raises(svxml.SvXmlError, match="grid answer has 3 values and 2 types"):
        svxml.grid_values(text)


def test_grid_values_too_many_values_is_an_svxml_error_not_index_error():
    text = _ans(1, 2, 0, ["a", "b", "c"], ["0", "0", "0"])
    with pytest.raises(svxml.SvXmlError, match="1x2 slice"):
        svxml.grid_values(text)


def test_grid_values_refuse_several_ranges():
    text = _ans(1, 2, 0, ["a", "b"], ["0", "0"]).replace("</data>", '<range start="0" end="0"><vals>x</vals>'
                                                                  '<types>0</types></range></data>')
    with pytest.raises(svxml.SvXmlError, match="several ranges"):
        svxml.grid_values(text)


def test_grid_values_honour_range_start():
    text = _ans(2, 2, 2, ["c", "d"], ["0", "0"])
    assert svxml.grid_values(text) == [[None, None], ["c", "d"]]


def test_numbers_are_parsed_for_any_data_type():
    text = _ans(1, 3, 0, ["100", "1.2E7", "abc"], ["1", "3", "1"])
    assert svxml.grid_values(text) == [[100.0, 12000000.0, "abc"]]
