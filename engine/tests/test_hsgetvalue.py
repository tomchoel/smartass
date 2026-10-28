from smartview_mcp.hsgetvalue import FINMAX_DEFAULTS, build_hsgetvalue_formula


def test_finmax_defaults_defined():
    assert "All Scenarios" in FINMAX_DEFAULTS
    assert FINMAX_DEFAULTS["All Currencies"] == "NOK"
    assert FINMAX_DEFAULTS["All Balance Types"] == "Periodic"


def test_build_hsgetvalue_formula_basic():
    coords = {
        "Accounts": "ac.EBITDA",
        "Entities": "un.SY.2WHS",
        "Period": "Aug 26",
    }
    formula = build_hsgetvalue_formula("FinMax", coords)
    assert formula.startswith('=HsGetValue("FinMax", ')
    assert '"All Accounts#ac.EBITDA"' in formula
    assert '"All Entities#un.SY.2WHS"' in formula
    assert '"All Time Periods#Aug 26"' in formula
    # Default dimensions should be filled
    assert '"All Functions#All Functions"' in formula
    assert '"All Products#All Products"' in formula


def test_build_hsgetvalue_formula_override_defaults():
    coords = {
        "All Accounts": "ac.C30",
        "All Entities": "un.SY.2TEC",
        "All Time Periods": "Sep 26",
        "Currencies": "EUR",
    }
    formula = build_hsgetvalue_formula("FinMax", coords)
    assert '"All Currencies#EUR"' in formula
    assert '"All Currencies#NOK"' not in formula
