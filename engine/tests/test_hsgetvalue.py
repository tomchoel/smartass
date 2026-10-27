from smartview_mcp.hsgetvalue import FINMAX_DEFAULTS, build_hsgetvalue_formula


def test_finmax_defaults_defined():
    assert "Scenarios" in FINMAX_DEFAULTS
    assert FINMAX_DEFAULTS["Currencies"] == "NOK"


def test_build_hsgetvalue_formula_basic():
    coords = {
        "Accounts": "ac.EBITDA",
        "Entities": "un.SY.2WHS",
        "Period": "Aug",
        "Years": "FY26",
    }
    formula = build_hsgetvalue_formula("FinMax", coords)
    assert formula.startswith('=HsGetValue("FinMax", ')
    assert '"Accounts#ac.EBITDA"' in formula
    assert '"Entities#un.SY.2WHS"' in formula
    assert '"Period#Aug"' in formula
    assert '"Years#FY26"' in formula
    # Default dimensions should be filled
    assert '"Functions#fu.NONE"' in formula
    assert '"Products#pd.NONE"' in formula


def test_build_hsgetvalue_formula_override_defaults():
    coords = {
        "Accounts": "ac.C30",
        "Entities": "un.SY.2TEC",
        "Period": "Sep",
        "Years": "FY26",
        "Currencies": "EUR",
    }
    formula = build_hsgetvalue_formula("FinMax", coords)
    assert '"Currencies#EUR"' in formula
    assert '"Currencies#NOK"' not in formula
