from smartass.knowledge.interim import catalog_metrics


def test_metrics_join_cubes_on_concept():
    by = {m.name: m for m in catalog_metrics()}
    assert by["Revenue"].members == {"FinMax": ("ac.C30",), "FinPlan": ("Revenue",)}
    assert by["Opex"].members == {"FinMax": ("ac.C70",), "FinPlan": ("OPEX",)}
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
    by = {m.name: m for m in catalog_metrics()}
    assert by["Revenue"].members.get("FinLite") is None      # FinLite is opted in per metric in sub-spec 2
