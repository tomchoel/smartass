"""drill_direct: drill lines from EPM's own drill query, without a page per piece (figures invented)."""
import json

import pytest

from smartview_mcp import drill_direct as dd
from smartview_mcp.drill_browser import CellTooWide, DrillError

URL = "aif/m/FdmeeMain.jsp?module=aif.drilldown&DRILLURL=Drill&sso_token=T&attribute=Entities.id.un.SY"


def test_sui_url_drops_drillurl_and_asks_for_json():
    u = dd.sui_url("https://epm", "./" + URL, 123)
    assert u == ("https://epm/aif/m/FdmeeMain.jsp?module=aif.drilldown&sso_token=T&attribute=Entities.id.un.SY"
                 "&FORMAT=SUI&_=123")


def test_sui_url_passes_values_through_byte_for_byte():
    u = dd.sui_url("https://epm/", "./aif/x.jsp?a=b%20c&DRILLURL=Drill&t=A+B/c=%2F&e=", 5)
    assert u == "https://epm/aif/x.jsp?a=b%20c&t=A+B/c=%2F&e=&FORMAT=SUI&_=5"


def _sui(**kw):
    return json.dumps({"DRILL_LOG": "/tmp/x", "APPLICATION_ID": 52, "PARTITIONKEY": 39, "status": 0,
                       "QUERY": "SELECT 1", "PARAMSARRAY": ["39", "un.SY"], "PARAMSTYPEARRAY": ["BigDecimal", "String"],
                       **kw})


def test_parse_sui_gives_epms_query_unchanged():
    q = dd.parse_sui(_sui())
    assert q.app_id == 52 and q.partition == 39
    assert json.loads(q.body) == {"QUERY": "SELECT 1", "PARAMS": ["39", "un.SY"],
                                  "PARAMSTYPE": ["BigDecimal", "String"]}


def test_parse_sui_no_data_is_a_note_not_an_error():
    with pytest.raises(dd.NoDrillData, match="not loaded to this intersection"):
        dd.parse_sui(json.dumps({"status": 1, "DATA": [], "details": "Drill query does not return any data. "
                                 "Data was not loaded to this intersection using Data Integration."}))


def test_parse_sui_too_wide():
    with pytest.raises(CellTooWide, match="EPMFDM-140373"):
        dd.parse_sui(json.dumps({"status": 1, "details": "EPMFDM-140373:The drill from summary operation ..."}))


def test_parse_sui_not_json_means_signed_out():
    with pytest.raises(dd.SignedOut):
        dd.parse_sui("<html>epmexit.jsp</html>")


def test_parse_sui_other_refusal():
    with pytest.raises(DrillError, match="boom"):
        dd.parse_sui(json.dumps({"status": 1, "details": "boom"}))
