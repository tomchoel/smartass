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


def test_run_slice_retains_comparison_only_rows_and_handles_missing_variance(tmp_path):
    db_path, db = _rich_atlas_db(tmp_path)
    with atlas.connect(db_path) as c:
        atlas.store(c, "FinMax", "All Entities", [
            {"name": "All Entities", "label": "All Entities", "parent": None, "depth": 0, "order": 0},
            {"name": "un.SY.VB20", "label": "VB20 Core (un.SY)", "parent": "All Entities", "depth": 1, "order": 1},
            {"name": "un.SY.A406", "label": "A406 Access (un.SY)", "parent": "All Entities", "depth": 1, "order": 2},
        ], 0.1)

    def value(rows, cols, pov):
        ent = rows[0]
        per = cols[0]
        if ent == "un.SY.A406":
            return -50.0 if per == "Apr 26" else -40.0
        # VB20 has LY only (Apr 25), #Missing in Apr 26!
        return -110.0 if per == "Apr 25" else None

    sv = svhttp.SvHttp(EchoTransport(value), home="https://plan1.example", db_path=db_path)
    prov = SvHttpProvider(sv)
    res = run_slice(
        prov, _DummyResolver(db), {"FinMax": (2026, 9)},
        cube="FinMax",
        rows={"All Entities": ["un.SY.VB20", "un.SY.A406"]},
        pov={"All Accounts": "ac.43200"},
        periods="Apr 26",
        compare=["ly"],
        suppress_missing=True,
        limit=1,
    )
    # Must preserve VB20 (axis order: first) even though count>limit; row_count=2 before slicing to limit=1
    assert res["row_count"] == 2 and res["truncated"] is True
    r0 = res["rows"][0]
    assert r0["All Entities"] == "un.SY.VB20"
    assert r0["Apr 26"] is None and r0["Apr 26 [ly]"] == -110.0
    assert r0["Apr 26 var vs ly"] == pytest.approx(110.0)
    assert r0["Apr 26 var% vs ly"] == pytest.approx(1.0)

    # When comparison (LY) is #Missing and AC has a value: variance is AC - 0, var_pct is None!
    def value_ac_only(rows, cols, pov):
        return 50.0 if cols[0] == "Apr 26" else None
    sv2 = svhttp.SvHttp(EchoTransport(value_ac_only), home="https://plan1.example", db_path=db_path)
    res2 = run_slice(SvHttpProvider(sv2), _DummyResolver(db), {"FinMax": (2026, 9)},
                     cube="FinMax", rows={"All Entities": ["un.SY.VB20"]}, pov={"All Accounts": "ac.43200"},
                     periods="Apr 26", compare=["ly"], suppress_missing=True)
    r_ac = res2["rows"][0]
    assert r_ac["Apr 26 var vs ly"] == pytest.approx(50.0)
    assert r_ac["Apr 26 var% vs ly"] is None


def test_expand_leaves_past_1000_nodes_and_tree_truncation_metadata(tmp_path):
    p = tmp_path / "big_atlas.db"
    db = atlas.connect(p)
    # 1100 intermediate parent nodes BEFORE a single leaf node!
    nodes = [{"name": "All Products", "label": "All Products", "parent": None, "depth": 0, "order": 0},
             {"name": "pd.SY", "label": "SY (pd)", "parent": "All Products", "depth": 1, "order": 1}]
    for i in range(1100):
        nodes.append({"name": f"pd.SY.INT{i:04d}", "label": f"INT{i}", "parent": "pd.SY", "depth": 2, "order": i})
    nodes.append({"name": "pd.SY.U9500", "label": "U9500 Leaf", "parent": "pd.SY.INT0000", "depth": 3, "order": 0})
    atlas.store(db, "FinMax", "All Products", nodes, 0.1)
    for dname in ("All Accounts", "All Entities", "All Time Periods", "All Scenarios", "All Currencies", "All Balance Types"):
        atlas.store(db, "FinMax", dname, [{"name": dname, "label": dname, "parent": None, "depth": 0, "order": 0}], 0.1)
    atlas.store(db, "FinMax", "All Accounts", [
        {"name": "All Accounts", "label": "All Accounts", "parent": None, "depth": 0, "order": 0},
        {"name": "ac.C30", "label": "C30", "parent": "All Accounts", "depth": 1, "order": 1},
    ], 0.1)

    r = _DummyResolver(db)
    # Tree reports truncation + effective limits honestly
    t = browse_tree(r, cube="FinMax", parent="pd.SY", depth=3, limit=50)
    assert t["truncated"] is True and t["has_more"] is True and t["returned"] == 50
    assert t["effective_depth"] == 3 and t["effective_limit"] == 50

    # expand="leaves" filters to leaves BEFORE limiting, finding pd.SY.U9500 past 1100 intermediates!
    sv = svhttp.SvHttp(EchoTransport(lambda rows, cols, pov: 14.0 if "pd.SY.U9500" in rows else None),
                       home="https://plan1.example", db_path=p)
    prov = SvHttpProvider(sv)
    res = run_slice(prov, r, {"FinMax": (2026, 9)}, cube="FinMax", rows={"All Products": ["pd.SY"]},
                    expand={"All Products": "leaves"}, pov={"All Accounts": "ac.C30"}, periods="Sep 26")
    assert any(row["All Products"] == "pd.SY.U9500" for row in res["rows"])


def test_slice_ytd_total_sums_flows_and_skips_summing_levels(tmp_path):
    db_path, db = _rich_atlas_db(tmp_path)
    atlas.store_cards(db, "FinMax", "All Accounts", [
        ("ac.43200", "ac.C40S", "+", "flow", "Expense", 0),
        ("ac.STA20Y", "All Accounts", "+", "balance_last", "Asset", 0),
    ])
    with atlas.connect(db_path) as c:
        c.execute("INSERT OR IGNORE INTO member VALUES ('FinMax','All Accounts','ac.STA20Y','STA20Y FTE','All Accounts',1,99)")
        c.commit()

    def value(rows, cols, pov):
        return -10.0 if rows[0] == "ac.43200" else 26.0

    sv = svhttp.SvHttp(EchoTransport(value), home="https://plan1.example", db_path=db_path)
    res = run_slice(SvHttpProvider(sv), _DummyResolver(db), {"FinMax": (2026, 9)},
                    cube="FinMax", rows={"All Accounts": ["ac.43200", "ac.STA20Y"]},
                    periods="Jan 26..Mar 26", total=True)
    by_acc = {r["All Accounts"]: r for r in res["rows"]}
    assert by_acc["ac.43200"]["Total"] == pytest.approx(-30.0)
    assert by_acc["ac.STA20Y"]["Total"] is None
    assert by_acc["ac.STA20Y"]["Period End"] == pytest.approx(26.0)


def test_expand_reconciliation_exact_and_skipped_for_shared(tmp_path):
    db_path, db = _rich_atlas_db(tmp_path)
    # Clean single-placement parent ac.C40S -> ac.43200 (-120) + ac.41000 (-80) = -200
    def value(rows, cols, pov):
        return {"ac.C40S": -200.0, "ac.43200": -120.0, "ac.41000": -80.0}.get(rows[0])

    sv = svhttp.SvHttp(EchoTransport(value), home="https://plan1.example", db_path=db_path)
    res = run_slice(SvHttpProvider(sv), _DummyResolver(db), {"FinMax": (2026, 9)},
                    cube="FinMax", rows={"All Accounts": ["ac.C40S"]}, expand={"All Accounts": "children"},
                    pov={"All Entities": "un.SY.5BMT_T"}, periods="closed")
    assert res["reconciliation"][0]["status"] == "exact" and res["reconciliation"][0]["gap"] == 0.0

    # Shared member under All Products -> reconciliation explicitly skipped!
    with atlas.connect(db_path) as c:
        atlas.store(c, "FinMax", "All Products", [
            {"name": "All Products", "label": "All Products", "parent": None, "depth": 0, "order": 0},
            {"name": "pd.SY.P1", "label": "P1", "parent": "All Products", "depth": 1, "order": 1},
            {"name": "pd.SY.S1", "label": "S1", "parent": "pd.SY.P1", "depth": 2, "order": 1},
            {"name": "pd.SY.S1", "label": "S1", "parent": "All Products", "depth": 1, "order": 2},
        ], 0.1)
    sv2 = svhttp.SvHttp(EchoTransport(lambda *a: 10.0), home="https://plan1.example", db_path=db_path)
    res2 = run_slice(SvHttpProvider(sv2), _DummyResolver(db), {"FinMax": (2026, 9)},
                     cube="FinMax", rows={"All Products": ["pd.SY.P1"]}, expand={"All Products": "children"},
                     pov={"All Accounts": "ac.C40S", "All Entities": "un.SY.5BMT_T"}, periods="closed")
    assert res2["reconciliation"][0]["status"] == "skipped"
    assert "shared" in res2["reconciliation"][0]["reason"].lower()
