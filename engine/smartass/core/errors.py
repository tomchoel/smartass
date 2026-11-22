"""The engine's own errors, neutral: no COM, no Excel. A job in the user's Excel, the EPM routes and the server raise
and catch these."""
from __future__ import annotations


class HostError(Exception):
    """Raised for failures the model should see (message is meant to be actionable)."""


class HostTimeout(HostError):
    pass


class InputError(ValueError):
    """Bad input from the caller (spec, address, name, path) - not a host failure, so no restart advice."""
