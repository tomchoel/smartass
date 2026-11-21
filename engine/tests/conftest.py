import os

import pytest


@pytest.hookimpl(trylast=True)
def pytest_configure(config):
    # Keep the suite out of the user's real state folder (%LOCALAPPDATA%\smartass): set before smartass.core.config is imported.
    os.environ["SMARTASS_HOME"] = str(config._tmp_path_factory.mktemp("smartass_home"))


@pytest.fixture(autouse=True)
def _no_live_drill_routes(monkeypatch):
    """Tests never reach EPM: the HTTP drill link/values are off unless a test turns them on (drillstep then uses its
    Excel route, which the tests fake), and the token drill raises unless a test that drills sets its own fake."""
    from smartass.core import config
    from smartass.epm import drill_token

    def _no_epm():
        raise AssertionError("tests never reach EPM")
    monkeypatch.setattr(drill_token, "shared", _no_epm)
    monkeypatch.setattr(config, "ANSWER_PREFETCH", False)        # tests never reach the live HTTP route
    monkeypatch.setattr(config, "INPROC", False)                 # batches run over COM: no add-in build, no compiler
    monkeypatch.setenv("SV_TEST_NO_HTTP_DRILL", "1")
