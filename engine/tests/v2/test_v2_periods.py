import pytest

from smartass.model.periods import label, parse, resolve, shift, span

CLOSED = (2026, 9)


def test_label_and_parse_round_trip():
    assert label((2026, 9)) == "Sep 26"
    assert parse("Sep 26") == (2026, 9)
    assert parse(" Jan 25 ") == (2025, 1)


def test_parse_rejects_other_text():
    with pytest.raises(ValueError, match="use 'Sep 26'"):
        parse("September 2026")


def test_shift_crosses_years():
    assert shift((2026, 1), -1) == (2025, 12)
    assert shift((2026, 9), -12) == (2025, 9)
    assert shift((2025, 12), 1) == (2026, 1)


def test_span_inclusive_and_ordered():
    assert span((2025, 11), (2026, 2)) == [(2025, 11), (2025, 12), (2026, 1), (2026, 2)]
    with pytest.raises(ValueError):
        span((2026, 2), (2026, 1))


@pytest.mark.parametrize("expr, first, last, n", [
    ("closed", (2026, 9), (2026, 9), 1),
    ("closed-ytd", (2026, 1), (2026, 9), 9),
    ("closed-12m", (2025, 10), (2026, 9), 12),
    ("fy", (2026, 1), (2026, 12), 12),
    ("Mar 26", (2026, 3), (2026, 3), 1),
    ("Jan 26..Sep 26", (2026, 1), (2026, 9), 9),
])
def test_resolve(expr, first, last, n):
    got = resolve(expr, CLOSED)
    assert (got[0], got[-1], len(got)) == (first, last, n)
