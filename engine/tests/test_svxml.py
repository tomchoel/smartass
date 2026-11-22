import pytest
from svxml_util import answer

from smartass.epm import svxml
from smartass.epm.grid import Axis, Layout, parse


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
    # final review A, M10: the provider's parser sees no action inside CDATA or a comment - refused, any request
    "<req_ExecuteGrid><sID>x</sID><![CDATA[<action>retrieve</action>]]></req_ExecuteGrid>",
    "<req_ExecuteGrid><sID>x</sID><!-- <action>retrieve</action> --></req_ExecuteGrid>",
    "<req_ExecuteGrid><sID>x</sID><action>retrieve</action><!-- x --></req_ExecuteGrid>",
    "<req_OpenCube><sID>x</sID><cube><![cdata[FinMax]]></cube></req_OpenCube>",
    "<req_ConnectToProvider><!DOCTYPE x [<!ENTITY a 'submit'>]></req_ConnectToProvider>",
    "<req_ConnectToProvider><?x y?></req_ConnectToProvider>",
])
def test_check_read_only_refuses_everything_else(body):
    with pytest.raises(ValueError, match="refused"):
        svxml.check_read_only(body)


def test_a_refused_body_is_not_quoted():
    """A malformed body can start with a session id: the refusal names no part of it."""
    with pytest.raises(ValueError, match="refused") as e:
        svxml.check_read_only("<sID>sid-SECRET-123</sID><req_ExecuteGrid></req_ExecuteGrid>")
    assert "SECRET" not in str(e.value) and "sid-" not in str(e.value)


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


SECRETS = ("sid-7f3a9c",
           "eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJ0b20ifQ.c2lnbmF0dXJlX29mX3Rva2Vu",  # allow-secret (an invented token)
           "JSESSIONID=c00k1e9", "ORA_OCIS_42=ocis-c00k1e")
# Answers that are not the provider's, each carrying a session id, a token and cookies: one with an <exception>
# element but no desc passes the transport (svhttp._provider_xml), and its text must still never be quoted.
NOT_PROVIDER = [
    f"<exception><sID>{SECRETS[0]}</sID><token>{SECRETS[1]}</token></exception> Set-Cookie: {SECRETS[2]}; {SECRETS[3]}",
    f"<html><body>Signed out. <input name='sID' value='{SECRETS[0]}'> Bearer {SECRETS[1]} {SECRETS[2]} "
    f"{SECRETS[3]}</body></html>",
    f"sID={SECRETS[0]}&token={SECRETS[1]}&cookie={SECRETS[2]}&{SECRETS[3]}",
    f"<req_Connect><sID>{SECRETS[0]}</sID></req_Connect>{SECRETS[1]}{SECRETS[2]}{SECRETS[3]}",
]


@pytest.mark.parametrize("text", NOT_PROVIDER)
def test_an_answer_that_is_not_the_providers_is_never_quoted(text):
    """Its error text reaches the tool reply (a failed grid's warning, a drill link's error): it says what the answer
    looked like in words of our own, never a part of it - no session id, token or cookie can be in it."""
    with pytest.raises(svxml.SvXmlError) as e:
        svxml.raise_for_exception(text)
    said = str(e.value)
    assert said.startswith("not a Smart View provider answer (") and not any(s in said for s in SECRETS)
    assert not any(part in said for part in ("sID", "token", "cookie", "Bearer", "JSESSIONID"))


def _desc(text: str) -> str:
    """The provider's own <desc> error text as raise_for_exception passes it on (escaped in the XML, as EPM sends
    it)."""
    import html
    answer = f'<exception errcode="1000" type="error"><desc>{html.escape(text)}</desc></exception>'
    with pytest.raises(svxml.SvXmlError) as e:
        svxml.raise_for_exception(answer)
    return str(e.value)


def _invented(n: int = 24) -> str:
    """A random token-like string, made at run time (no realistic secret in a fixture)."""
    import secrets
    return secrets.token_urlsafe(n)


def test_a_session_id_tag_in_the_providers_own_error_text_is_scrubbed():
    sid = _invented()
    said = _desc(f"Session <sID>{sid}</sID> is not valid. Grid is Invalid")
    assert sid not in said and "Grid is Invalid" in said and "Session" in said


def test_a_jwt_shaped_string_in_the_providers_own_error_text_is_scrubbed():
    jwt = ".".join(["eyJ" + _invented(12), "eyJ" + _invented(20), _invented(30)])
    said = _desc(f"token {jwt} rejected for member ac.C60")
    assert jwt not in said and not any(part in said for part in jwt.split(".")) and "ac.C60" in said


def test_a_bearer_token_in_the_providers_own_error_text_is_scrubbed():
    tok = _invented()
    said = _desc(f"Authorization: Bearer {tok} - access denied")
    assert tok not in said and "access denied" in said


@pytest.mark.parametrize("cookies", ["Cookie: JSESSIONID={a}; ORA_OCIS_42={b}",
                                     "Set-Cookie: {a}={b}; Path=/; Secure",
                                     "session JSESSIONID={a} and ORA_OCIS_42={b} ended"])
def test_cookie_pairs_in_the_providers_own_error_text_are_scrubbed(cookies):
    a, b = _invented(8), _invented(8)
    said = _desc("Grid is Invalid. " + cookies.format(a=a, b=b))
    assert a not in said and b not in said and said.startswith("Grid is Invalid.")


@pytest.mark.parametrize("shape", [
    "Grid is Invalid. JSESSIONID = {x}",
    "Grid is Invalid. ssoToken = {x}",
    "Grid is Invalid. sessionId: {x}",
    'Grid is Invalid. {{"token": "{x}"}}',
    "Grid is Invalid. See https://h.example/p?code={x}&state=1",
    "Grid is Invalid. Authorization: Basic {x}",
    "Grid is Invalid. Basic {x}",
    "Grid is Invalid. X-Auth: {x}",
    "Grid is Invalid. https://h.example/p;sID%3D{x}",
    "Grid is Invalid. session &amp;lt;sID&amp;gt;{x}&amp;lt;/sID&amp;gt; ended",
])
def test_every_shape_of_a_session_key_or_query_in_the_providers_error_text_is_scrubbed(shape):
    x = _invented()
    said = _desc(shape.format(x=x))
    assert x not in said and said.startswith("Grid is Invalid.")


def test_a_bearer_token_is_replaced_once_and_cleanly():
    jwt = ".".join(["eyJ" + _invented(12), "eyJ" + _invented(20), _invented(30)])
    said = _desc(f"Authorization denied: Bearer {jwt} - access denied")
    assert "Bearer ***" in said and "removed]" not in said and "access denied" in said


def test_the_providers_own_error_text_is_capped_in_length():
    said = _desc("Grid is Invalid: " + "x" * 5000)
    assert said.startswith("Grid is Invalid: xxx") and len(said) < 700 and "5017 characters" in said
    assert _desc("Grid is Invalid: " + "y" * 400) == "Grid is Invalid: " + "y" * 400      # a usual text is kept whole


def test_data_sources_give_the_pod_per_app():
    text = ('<res_GetProvisionedDataSources><Product id="HP"><Server name="FinMax" '
            'context="https://plan1.example/HyperionPlanning/SmartView" appName="FinMax"/></Product>'
            '<Product id="HP"><Server name="FinPlan_for_DataLoad" '
            'context="https://strat.example/HyperionPlanning/SmartView" appName="FinPlan"/></Product>'
            '</res_GetProvisionedDataSources>')
    assert svxml.data_sources(text) == {"FinMax": [("FinMax", "https://plan1.example")],
                                        "FinPlan_for_DataLoad": [("FinPlan", "https://strat.example")]}


# Shape of the live list (2026-10-25): three servers share appName FinMax - prod, plan7 and the TEST pod.
THREE_FINMAX = ('<res_GetProvisionedDataSources>'
                '<Product id="HP"><Server name="FinMax" context="https://plan1-x.example/HyperionPlanning/SmartView" '
                'appName="FinMax"/></Product>'
                '<Product id="HP"><Server appName="FinMax" name="Plan 7 Prod" '
                'context="https://plan7-x.example/HyperionPlanning/SmartView"/></Product>'
                '<Product id="HP"><Server context="https://plan1-test-x.example/HyperionPlanning/SmartView" '
                'appName="FinMax" name="plan1-test"/></Product>'
                '</res_GetProvisionedDataSources>')


def test_data_sources_are_keyed_by_server_name_not_app():
    assert svxml.data_sources(THREE_FINMAX) == {"FinMax": [("FinMax", "https://plan1-x.example")],
                                                "Plan 7 Prod": [("FinMax", "https://plan7-x.example")],
                                                "plan1-test": [("FinMax", "https://plan1-test-x.example")]}


def test_a_duplicated_server_name_keeps_every_distinct_entry_and_never_raises():
    twice = THREE_FINMAX.replace('name="Plan 7 Prod"', 'name="FinMax"')
    assert svxml.data_sources(twice)["FinMax"] == [("FinMax", "https://plan1-x.example"),
                                                   ("FinMax", "https://plan7-x.example")]
    same = THREE_FINMAX.replace("</res_GetProvisionedDataSources>",
                                '<Server name="FinMax" appName="FinMax" '
                                'context="https://plan1-x.example/HyperionPlanning/SmartView"/>'
                                "</res_GetProvisionedDataSources>")
    assert svxml.data_sources(same)["FinMax"] == [("FinMax", "https://plan1-x.example")]   # a repeat is one entry


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
            '<Server appName="NoName" context="https://r.example/HyperionPlanning/SmartView"/></res_X>')
    assert svxml.data_sources(text) == {"n": [("FinMax", "https://p.example")]}


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


ENUM = ('<res_EnumMembers><overflow>0</overflow><mbrs enc="1">[CurrRepMonth].[Y26M09]</mbrs><atts enc="0">0</atts>'
        '<mbrDesc enc="1">Y26M09</mbrDesc></res_EnumMembers>')


def test_enum_children_is_read_only_and_names_the_provider_dimension():
    body = svxml.enum_children("S1", "TimePeriods", "CurrRepMonth")
    svxml.check_read_only(body)
    assert body.startswith("<req_EnumMembers><sID>S1</sID><dim>TimePeriods</dim>")
    assert '<filter name="Children"><arg id="0">CurrRepMonth</arg></filter>' in body


def test_members_parses_the_pipe_list():
    assert svxml.members(ENUM) == ["[CurrRepMonth].[Y26M09]"]
    assert svxml.members(ENUM.replace("[CurrRepMonth].[Y26M09]", "a&amp;b|c")) == ["a&b", "c"]
    assert svxml.members('<res_EnumMembers><mbrs enc="1"></mbrs></res_EnumMembers>') == []
    with pytest.raises(svxml.SvXmlError) as e:
        svxml.members('<exception errcode="4997" type="error"><desc>dimension invalid</desc></exception>')
    assert e.value.code == 4997
    with pytest.raises(svxml.SvXmlError, match="no member list"):
        svxml.members("<res_EnumMembers></res_EnumMembers>")


def test_echo_transport_answers_a_retrieve_with_the_value_function(tmp_path):
    from svxml_util import EchoTransport, atlas_with_dimensions, pnl_value

    from smartass.epm import svhttp
    from smartass.epm.grid import Axis
    sv = svhttp.SvHttp(EchoTransport(pnl_value), home="https://plan1.example",
                       db_path=atlas_with_dimensions(tmp_path / "a.db"))
    lay, p, _ = sv.retrieve("FinMax", {"All Balance Types": "Accumulated"},
                            [Axis("All Entities", ["un.SY.2WHS"]), Axis("All Accounts", ["ac.C30", "ac.STA20Y"])],
                            [Axis("All Time Periods", ["Mar 26"])], False, False)
    assert p.rows == [["un.SY.2WHS", "ac.C30", 3000.0], ["un.SY.2WHS", "ac.STA20Y", 10.0]]
    assert sv.closed_month() == "Sep 26"


def test_drill_reports_body_is_smart_views_own_shape():
    body = svxml.drill_reports("S1", "FinLite", 19, [("All Time Periods", "col", "Sep 26"),
                                                     ("All Accounts", "row", "C60080 Consultancy IS/IT (ac)"),
                                                     ("All Entities", "pov", "SY Telenor Norge AS (un)")])
    assert body == ('<req_GetDrillThroughReports><sID>S1</sID><usr></usr><pwd></pwd><rowGov></rowGov>'
                    '<timeGov></timeGov><cube>FinLite</cube><ords>19</ords><alsTbl>Default</alsTbl><pov/><cells>'
                    '<dim name="All Time Periods" col="0" enc="0">Sep 26</dim>'
                    '<dim name="All Accounts" row="0" enc="0">C60080 Consultancy IS/IT (ac)</dim>'
                    '<dim name="All Entities" pov="pov" enc="0">SY Telenor Norge AS (un)</dim>'
                    '</cells></req_GetDrillThroughReports>')
    svxml.check_read_only(body)                       # allowed: it only lists the reports


def test_drill_reports_escapes_labels():
    body = svxml.drill_reports("S", "FinLite", 3, [("All Entities", "pov", "R&D <x>")])
    assert "R&amp;D &lt;x&gt;" in body


def test_execute_drill_through_report_stays_refused():
    with pytest.raises(ValueError, match="ExecuteDrillThroughReport"):
        svxml.check_read_only("<req_ExecuteDrillThroughReport><sID>S</sID></req_ExecuteDrillThroughReport>")


def test_drill_url_unescapes_the_first_report():
    text = ('<?xml version="1.0"?><res_GetDrillThroughReports><authRequired>false</authRequired>'
            '<drillThroughURL name="Drill Through to source" type="post" launchType="both" urlenc="1">'
            './aif/x/FdmeeMain.jsp?module=aif.drilldown&amp;DRILLURL=Drill&amp;attribute=Accounts.id.ac.C1'
            '</drillThroughURL></res_GetDrillThroughReports>')
    assert svxml.drill_url(text) == ("Drill Through to source",
                                     "aif/x/FdmeeMain.jsp?module=aif.drilldown&DRILLURL=Drill&attribute=Accounts.id.ac.C1")


def test_drill_url_none_without_a_report():
    assert svxml.drill_url("<res_GetDrillThroughReports><authRequired>false</authRequired>"
                           "</res_GetDrillThroughReports>") is None


def test_drill_url_raises_provider_exceptions():
    with pytest.raises(svxml.SvXmlError, match="not valid"):
        svxml.drill_url('<exception errcode="1000" type="error"><desc>Your session is not valid.</desc></exception>')


def test_the_raw_desc_is_kept_for_the_classification_and_never_shown():
    """Final review A, M1: the error keeps the provider's raw text (unescaped, uncut) for svhttp's classification;
    its text, args and repr are the scrubbed one."""
    import secrets
    x = secrets.token_hex(12)
    with pytest.raises(svxml.SvXmlError) as e:
        svxml.raise_for_exception(f'<exception errcode="1000"><desc>Member Cookie: {x} &amp;amp; more</desc>'
                                  '</exception>')
    assert e.value.raw == f"Member Cookie: {x} &amp; more" and e.value.code == 1000
    assert x not in str(e.value) and x not in repr(e.value) and x not in str(e.value.args)
    assert svxml.SvXmlError("no member list").raw == "no member list"
