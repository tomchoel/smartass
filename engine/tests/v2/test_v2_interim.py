from smartass.knowledge.interim import catalog_metrics


def test_metrics_join_cubes_on_concept():
    by = {m.name: m for m in catalog_metrics()}
    assert by["Revenue"].members == {"FinMax": ("ac.C30",), "FinPlan": ("Revenue",), "FinLite": ("ac.C30",)}
    assert by["Opex"].members == {"FinMax": ("ac.C70",), "FinPlan": ("OPEX",), "FinLite": ("ac.C70",)}
    assert by["EBITDA"].members["FinPlan"] == ("EBITDA",)
    assert "FinPlan" not in by["Personnel cost"].members        # FinMax-only line


def test_signs_materiality_and_stock():
    by = {m.name: m for m in catalog_metrics()}
    assert by["Revenue"].sign == "income" and by["Opex"].sign == "cost" and by["EBITDA"].sign == "income"
    assert by["Other items"].sign == "neutral"
    assert by["Opex"].materiality == (5_000_000.0, 0.10)
    fte = by["FTE"]
    assert (fte.balance, fte.aggregation, fte.sign) == ("Accumulated", "last", "neutral")
    assert all(m.status == "approved" for m in by.values())


def test_finlite_shares_finmax_members():
    """Changed from the plan's 'FinLite is None': FinLite uses the same accounts as FinMax, so every non-stock metric
    with FinMax members gets them for FinLite; stock metrics (FTE) are not in FinLite."""
    by = {m.name: m for m in catalog_metrics()}
    assert by["Revenue"].members["FinLite"] == by["Revenue"].members["FinMax"] == ("ac.C30",)
    assert by["Opex"].members["FinLite"] == ("ac.C70",)
    assert "FinLite" not in by["FTE"].members
    for m in by.values():
        if "FinMax" in m.members and m.aggregation != "last":
            assert m.members["FinLite"] == m.members["FinMax"], m.name
