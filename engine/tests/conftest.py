import os

import pytest


@pytest.hookimpl(trylast=True)
def pytest_configure(config):
    # Keep the suite out of the user's real state folder (%LOCALAPPDATA%\smartass): set before smartview_mcp is imported.
    os.environ["SMARTASS_HOME"] = str(config._tmp_path_factory.mktemp("smartass_home"))


@pytest.fixture(autouse=True)
def _no_live_drill_routes(monkeypatch):
    """Tests never reach EPM: the HTTP drill link/values and the direct drill are off unless a test turns them on
    (drillstep then uses its Excel/page routes, which the tests fake)."""
    from smartview_mcp import config
    monkeypatch.setattr(config, "DRILL_DIRECT", False)
    monkeypatch.setenv("SV_TEST_NO_HTTP_DRILL", "1")
