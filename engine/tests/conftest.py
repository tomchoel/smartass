import os

import pytest


@pytest.hookimpl(trylast=True)
def pytest_configure(config):
    # Keep the suite out of the user's real state folder (%LOCALAPPDATA%\smartass): set before smartview_mcp is imported.
    os.environ["SMARTASS_HOME"] = str(config._tmp_path_factory.mktemp("smartass_home"))
