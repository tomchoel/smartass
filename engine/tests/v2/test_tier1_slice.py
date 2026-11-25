import pytest
from svxml_util import EchoTransport

from smartass.data.provider import SvHttpProvider
from smartass.data.slice import browse_tree, run_slice
from smartass.epm import svhttp
from smartass.knowledge import atlas
from smartass.model.validate import GuardError


def _rich_atlas_db(tmp_path):
    p = tmp_path / "atlas.db"
    db = atlas.connect(p)
    atlas.store(db, "FinMax", "All Accounts", [
        {"name": "All Accounts", "label": "All Accounts", "parent": None, "depth": 0, "order": 0},
        {"name": "ac.C40S", "label": "C40S COGS (ac)", "parent": "All Accounts", "depth": 1, "order": 1},
        {"name": "ac.43200", "label": "43200 Power COGS (ac)", "parent": "ac.C40S", "depth": 2, "order": 2},
        {"name": "ac.41000", "label": "41000 Interconnect (ac)", "parent": "ac.C40S", "depth": 2, "order": 3},
    ], 0.1)
    atlas.store(db, "FinMax", "All Entities", [
        {"name": "All Entities", "label": "All Entities", "parent": None, "depth": 0, "order": 0},
        {"name": "un.SY.2WHS", "label": "2WHS Wholesale (un.SY)", "parent": "All Entities", "depth": 1, "order": 1},
        {"name": "un.SY.5BMT_T", "label": "5BMT_T Core (un.SY)", "parent": "un.SY.2WHS", "depth": 2, "order": 2},
    ], 0.1)
    atlas.store(db, "FinMax", "All Counterparts", [
        {"name": "All Counterparts", "label": "All Counterparts", "parent": None, "depth": 0, "order": 0},
        {"name": "cp.TN_GROUP", "label": "TN_GROUP Telenor Group (cp)", "parent": "All Counterparts", "depth": 1, "order": 1},
    ], 0.1)
    for dim in ("All Time Periods", "All Scenarios", "All Currencies", "All Balance Types", "All Products"):
        atlas.store(db, "FinMax", dim, [{"name": dim, "label": dim, "parent": None, "depth": 0, "order": 0}], 0.1)
    atlas.store_cards(db, "FinMax", "All Accounts", [
        ("ac.43200", "ac.C40S", "+", "flow", "Expense", 0),
        ("ac.41000", "ac.C40S", "+", "flow", "Expense", 0),
    ])
    return p, db


class _DummyResolver:
    def __init__(self, db):
        self.db = db

    def canonical(self, member, cube):
        return member

    def exists(self, member, cube):
        return True

    def members(self, metric, cube):
        from smartass.model.resolver import UnknownMetric
        if metric == "COGS":
            return ("ac.C40S",)
        raise UnknownMetric(metric)

    def scenarios(self):
        return {"FinMax": {}, "FinPlan": {"budget": ["Target", "Final"], "forecast": ["Forecast", "Final"]}}

    def latest_forecast(self):
        return "Forecast Q326"

    def caveats(self, touched, metric_ids=frozenset(), months=None):
        return [{"id": "c1", "text": "test caveat", "severity": "info"}]

    def card_caveats(self, touched):
        return []

    def note(self, name, cube):
        return "power note" if name == "ac.43200" else None

    def compact_card(self, name, cube):
        return None


def test_atlas_dimensions_and_tree_walk(tmp_path):
    _, db = _rich_atlas_db(tmp_path)
    r = _DummyResolver(db)
    dims = browse_tree(r, cube="FinMax")
    assert any(d["dimension"] == "All Counterparts" for d in dims["dimensions"])

    sub = browse_tree(r, cube="FinMax", parent="ac.C40S", depth=1)
    assert [m["name"] for m in sub["members"]] == ["ac.43200", "ac.41000"]
    assert sub["members"][0]["card"]["account_type"] == "Expense"
    assert sub["members"][0]["note"] == "power note"

    hits = browse_tree(r, cube="FinMax", dimension="All Counterparts", query="TN_GROUP")
    assert hits["members"][0]["name"] == "cp.TN_GROUP"


def test_run_slice_supports_arbitrary_dimensions_expand_and_compare(tmp_path):
    db_path, db = _rich_atlas_db(tmp_path)
    seen = []

    def value(rows, cols, pov):
        seen.append((rows, cols, pov))
        if rows[0] == "ac.43200":
            return -120.0 if cols[0] == "Sep 26" else -100.0
        return -80.0 if cols[0] == "Sep 26" else -80.0

    sv = svhttp.SvHttp(EchoTransport(value), home="https://plan1.example", db_path=db_path)
    prov = SvHttpProvider(sv)
    res = run_slice(
        prov, _DummyResolver(db), {"FinMax": (2026, 9)},
        cube="FinMax",
        rows={"All Accounts": ["COGS"], "All Counterparts": ["cp.TN_GROUP"]},
        expand={"All Accounts": "children"},
        pov={"All Entities": "un.SY.5BMT_T"},
        periods="closed",
        compare=["ly"],
    )
    assert res["row_count"] == 2
    assert any("un.SY.5BMT_T" in pov for _, _, pov in seen)
    r0 = next(r for r in res["rows"] if r["All Accounts"] == "ac.43200")
    assert r0["All Counterparts"] == "cp.TN_GROUP"
    assert r0["Sep 26"] == -120.0
    assert r0["Sep 26 [ly]"] == -100.0
    assert r0["Sep 26 var vs ly"] == -20.0
    assert r0["Sep 26 var% vs ly"] == pytest.approx(-0.2)


def test_run_slice_enforces_closed_month_unless_preliminary(tmp_path):
    db_path, db = _rich_atlas_db(tmp_path)
    sv = svhttp.SvHttp(EchoTransport(lambda *a: 1.0), home="https://plan1.example", db_path=db_path)
    prov = SvHttpProvider(sv)
    with pytest.raises(GuardError, match="after the last closed month Sep 26"):
        run_slice(prov, _DummyResolver(db), {"FinMax": (2026, 9)},
                  cube="FinMax", rows={"All Accounts": ["ac.43200"]}, periods="Oct 26")
    ok = run_slice(prov, _DummyResolver(db), {"FinMax": (2026, 9)},
                   cube="FinMax", rows={"All Accounts": ["ac.43200"]}, periods="Oct 26", preliminary=True)
    assert any("preliminary" in w for w in ok["warnings"])
