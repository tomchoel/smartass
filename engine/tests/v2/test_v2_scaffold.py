def test_package_imports():
    import smartass.model

    assert smartass.model.__doc__.startswith("Pure model layer")
