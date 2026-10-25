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
