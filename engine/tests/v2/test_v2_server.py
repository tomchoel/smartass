# engine/tests/v2/test_v2_server.py
import pytest
from v2util import resolver

from smartass import server


def test_three_tools_registered():
    assert server.TOOL_NAMES == ("fin_analyze", "fin_lookup", "fin_status")


def test_lookup_metrics_and_members(monkeypatch):
    monkeypatch.setattr(server, "_resolver", lambda: resolver())
    monkeypatch.setattr(server, "_search_members", lambda text, cube, limit: [
        {"connection": "FinMax", "dimension": "All Entities", "name": "un.SY.2BUS", "label": "2BUS Business Norway"}])
    out = server.lookup("opex", kind="any")
    assert out["metrics"][0]["name"] == "Opex"
    assert out["metrics"][0]["members"]["FinMax"] == ["ac.C50S", "ac.C60"]
    assert out["members"][0]["name"] == "un.SY.2BUS"


def test_lookup_kind_metric_only(monkeypatch):
    monkeypatch.setattr(server, "_resolver", lambda: resolver())
    monkeypatch.setattr(server, "_search_members", lambda *a: pytest.fail("members not asked"))
    assert server.lookup("ebitda", kind="metric")["members"] == []


def test_instructions_are_short():
    assert len(server.INSTRUCTIONS) < 2500
