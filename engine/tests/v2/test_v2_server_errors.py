import asyncio

import pytest

from smartass import server
from smartass.model.errors import GuardError
from smartview_mcp.host import InputError


def test_guard_error_in_a_job_becomes_input_error(monkeypatch):
    seen = {}

    async def fake_call_user(fn, what):
        try:
            return fn(None)
        except Exception as e:                 # what v1's host would see
            seen["type"] = type(e)
            raise

    monkeypatch.setattr(server.kit, "_call_user", fake_call_user)

    def job(h):
        raise GuardError("no v2 answer here", "ask a new question")

    with pytest.raises(InputError, match="no v2 answer here - ask a new question"):
        asyncio.run(server._user(job, "x"))
    assert seen["type"] is InputError
