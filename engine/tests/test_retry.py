"""core/retry.tried_again: one bounded retry for a file operation Windows refuses for a moment (the seed's replace, a
new workbook's delete, setup's config replace)."""
import pytest

from smartass.core import retry


def _failing(times: int, error=PermissionError):
    calls = []

    def fn():
        calls.append(1)
        if len(calls) <= times:
            raise error(13, "Access is denied")
        return "done"
    return fn, calls


def test_it_returns_what_the_call_returns_once_it_goes_through(monkeypatch):
    pauses = []
    monkeypatch.setattr(retry.time, "sleep", pauses.append)
    fn, calls = _failing(2)
    assert retry.tried_again(fn, 10, 0.2) == "done" and len(calls) == 3 and pauses == [0.2, 0.2]


def test_after_its_tries_the_last_error_is_raised_with_no_pause_after_it(monkeypatch):
    pauses = []
    monkeypatch.setattr(retry.time, "sleep", pauses.append)
    fn, calls = _failing(100)
    with pytest.raises(PermissionError):
        retry.tried_again(fn, 4, 0.1)
    assert len(calls) == 4 and pauses == [0.1] * 3


def test_another_error_is_raised_at_once_and_on_names_what_is_tried_again(monkeypatch):
    monkeypatch.setattr(retry.time, "sleep", lambda s: None)
    fn, calls = _failing(100, FileNotFoundError)
    with pytest.raises(FileNotFoundError):
        retry.tried_again(fn, 4, 0.1)
    assert len(calls) == 1
    fn, calls = _failing(2, FileNotFoundError)
    assert retry.tried_again(fn, 4, 0.1, on=OSError) == "done" and len(calls) == 3
