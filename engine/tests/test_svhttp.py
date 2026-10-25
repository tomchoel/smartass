from smartview_mcp import atlas


def _store(db, conn, dims):
    for d in dims:
        atlas.store(db, conn, d, [{"name": d, "label": d, "parent": None, "depth": 0, "order": 0}], 0.1)


def test_atlas_dimensions_sorted_per_connection(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    _store(db, "FinMax", ["All Time Periods", "All Accounts"])
    _store(db, "FinPlan", ["Account"])
    assert atlas.dimensions(db, "FinMax") == ["All Accounts", "All Time Periods"]
    assert atlas.dimensions(db, "FinLite") == []
