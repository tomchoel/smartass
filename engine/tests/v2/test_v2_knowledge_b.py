"""Knowledge layer phase B (docs/plans/2026-11-10-knowledge-b-cards-metrics.md)."""
import pytest
from smartass.knowledge.cards import account_type, operators, time_balance

KIDS = {"P": ["A", "B", "C"], "Q": ["D", "E"], "Z": ["F", "G"], "N": ["H"]}


def test_a_plain_sum_is_plus_for_every_child():
    got = operators(KIDS, {"P": 6.0, "A": 1.0, "B": 2.0, "C": 3.0})
    assert got[("P", "A")] == got[("P", "B")] == got[("P", "C")] == "+"


def test_one_child_left_out_is_tilde():
    got = operators(KIDS, {"Q": 5.0, "D": 5.0, "E": 7.0})
    assert got[("Q", "D")] == "+" and got[("Q", "E")] == "~"


def test_one_child_subtracted_is_minus():
    got = operators(KIDS, {"Q": -2.0, "D": 5.0, "E": 7.0})
    assert got[("Q", "E")] == "-" and got[("Q", "D")] == "+"


def test_a_zero_parent_over_children_with_data_ignores_them_all():
    got = operators(KIDS, {"Z": 0.0, "F": 4.0, "G": -1.0})
    assert got[("Z", "F")] == got[("Z", "G")] == "~"


def test_no_data_and_unresolved_are_unknown():
    got = operators(KIDS, {"P": 100.0, "A": 1.0, "B": 2.0, "C": 3.0})
    assert got[("P", "A")] is None and got[("N", "H")] is None


def test_two_children_that_could_explain_the_gap_are_not_guessed():
    got = operators({"P": ["A", "B", "C"]}, {"P": 7.0, "A": 3.0, "B": 3.0, "C": 7.0})   # leave out A or B?
    assert got == {("P", "A"): None, ("P", "B"): None, ("P", "C"): None}


def test_a_blocked_cell_says_nothing_about_its_parent():
    # a #NoAccess parent reads as no value; without the block it would look like a zero parent -> all '~'
    got = operators({"P": ["A", "B"]}, {"A": 4.0, "B": 1.0}, blocked=[{"P"}])
    assert got == {("P", "A"): None, ("P", "B"): None}


def test_a_second_slice_settles_what_the_first_could_not():
    first = operators(KIDS, {"Q": 5.0, "D": 5.0})                     # E has no data: + for D, E unknown
    both = operators(KIDS, {"Q": 5.0, "D": 5.0}, {"Q": 5.0, "D": 5.0, "E": 7.0}, prior=first)
    assert first[("Q", "E")] is None and both[("Q", "E")] == "~"


def test_time_balance():
    months = [100.0] + [1.0] * 11                                     # a level booked as movements (FinMax FTE)
    assert time_balance(sum(months), months, movements=True) == "balance_last"
    flow = [10.0] * 12
    assert time_balance(120.0, flow, movements=True) == "flow"
    level = [5.0, 6.0, 7.0, 8.0, 9.0, 9.0, 9.0, 9.0, 9.0, 9.0, 9.0, 10.0]  # FinPlan subscriptions: year = December
    assert time_balance(10.0, level) == "balance_last"
    assert time_balance(sum(flow), flow) == "flow"
    assert time_balance(None, [None] * 12) is None


def test_time_balance_real_finmax_cases():
    """Probe 10.11, FinMax un.SY Actual Yr 25 Periodic: STA20Y (FTE) is a level, C70 (opex) a flow."""
    sta20y = [2857.3, -5.92, -13.2, -17.38, 0.65, -2.55, -11.4, -3.1, -45.8, 20.2, 7.44, -4.12]
    assert time_balance(2782.12, sta20y, movements=True) == "balance_last"
    c70 = [-576936628.52, -572493138.1, -616739429.71, -586487288.89, -586482799.47, -583876694.42, -467559217.5,
           -587523845.59, -634055040.31, -623499052.22, -627667099.78, -688602819.56]
    assert time_balance(-7151923054.07, c70, movements=True) == "flow"


def test_a_january_only_booking_is_undecidable_not_a_level():
    once = [-50000.0] + [None] * 11                                   # an annual licence booked in January
    assert time_balance(-50000.0, once, movements=True) is None
    assert time_balance(-50000.0, [None] * 11 + [-50000.0]) is None   # FinPlan: one month says nothing


def test_time_balance_real_finplan_case():
    """Probe 10.11, FinPlan Target/Final FY26: Fixed Broadband subscriptions YearTotal = December; OPEX = the sum."""
    fbb = [700187.23, 698784.31, 697729.71, 696786.67, 695852.95, 694021.28, 692096.57, 691079.75, 690501.48,
           688848.66, 688171.16, 686996.55]
    assert time_balance(686996.55, fbb) == "balance_last"
    opex = [-582803492.09, -590294719.30, -605919633.95, -605431245.93, -624155541.18, -582812267.83, -487715419.41,
            -587923235.02, -640074598.20, -612754330.98, -604144384.23, -605672388.79]
    assert time_balance(sum(opex), opex) == "flow"


def test_account_type_from_branch_and_sign():
    assert account_type(["ac.C30", "ac.C48", "ac.EBITDA"], 5.0) == "revenue"
    assert account_type(["ac.C70", "ac.EBITDA_BOI"], -5.0) == "expense"
    assert account_type(["ac.STA20Y", "ac.STA", "ac.Statistics"], 3.0) == "statistic"
    assert account_type(["ac.C181", "ac.C1_C2"], 3.0) == "balance"
    assert account_type(["ac.X"], None) is None


def test_account_type_uses_every_hierarchy_and_signs_only_inside_the_pnl():
    # STA20Y's first parent (by atlas order) is the export tree: the union of ancestors still finds ac.STA
    assert account_type(["ac.STA20Y", "ac.Export", "All Accounts", "ac.STA", "ac.Statistics"], 3.0) == "statistic"
    assert account_type(["ac.X1", "ac.C89T", "ac.C3_C8"], -3.0) == "expense"
    assert account_type(["Revenue", "Gross_Profit", "P&L"], 9.0) == "revenue"                 # FinPlan
    assert account_type(["Fixed Broadband subscriptions", "Statistics"], 9.0) == "statistic"   # FinPlan
    assert account_type(["ac.EBITDA", "Ratios", "Account"], 0.3) is None      # a ratio is no revenue
    assert account_type(["ac.X", "ac.Export"], 5.0) is None                   # outside the P&L no sign guess


# --- Task 3: fetch the slices and store the cards in atlas.db --------------------------------------------------------

from smartass.knowledge import atlas, cards_fetch  # noqa: E402


class FakeSv:
    """fetch(conn, pov, rows, cols) -> a Fetched-like grid in Smart View's shape: row 0 POV, rows 1..nc column
    headers, then one row per member with its label in column 0 and the values from column nr on
    (values[1 + nc + j][nr + k], as svhttp.Fetched). `values`: {(column member, row member): value}, or a function
    (pov, column member, row member) -> value."""

    def __init__(self, values):
        self.values, self.calls = values, []

    def get(self, pov, col, name):
        if callable(self.values):
            return self.values(pov, col, name)
        return self.values.get((col, name), "#Missing")

    def fetch(self, conn, pov, rows, cols):
        self.calls.append((conn, dict(pov), list(rows[0].members), list(cols[0].members)))
        names, cs = rows[0].members, cols[0].members

        class L:
            nr, nc = 1, 1
        grid = [[None] * (1 + len(cs)), [None, *cs]] + [[n, *(self.get(pov, c, n) for c in cs)] for n in names]
        return type("F", (), {"layout": L, "values": grid})()


def _atlas(tmp_path, rows):
    db = atlas.connect(tmp_path / "a.db")
    db.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)", rows)
    db.commit()
    return db


def test_derive_dimension_stores_operators_and_shared(tmp_path):
    db = _atlas(tmp_path, [
        ("FinMax", "All Entities", "All Entities", "All Entities", None, 0, 0),
        ("FinMax", "All Entities", "un.SY", "SY", "All Entities", 1, 1),
        ("FinMax", "All Entities", "A", "A", "un.SY", 2, 2),
        ("FinMax", "All Entities", "Ai", "Ai elim", "un.SY", 2, 3),
        ("FinMax", "All Entities", "A", "A", "OTHER", 2, 4)])
    sv = FakeSv({("ac.C70", "un.SY"): 10.0, ("ac.C70", "A"): 10.0, ("ac.C70", "Ai"): 4.0,
                 ("ac.C70", "All Entities"): 10.0})
    n = cards_fetch.derive_dimension(sv, db, "FinMax", "All Entities", chunk=2, year=2025)["rows"]
    rows = {(r[0], r[1]): r for r in db.execute("SELECT name, parent, operator, shared, account_type FROM card")}
    assert rows[("A", "un.SY")][2] == "+" and rows[("Ai", "un.SY")][2] == "~"
    assert rows[("A", "un.SY")][3] == 1 and rows[("Ai", "un.SY")][3] == 0
    assert rows[("A", "un.SY")][4] is None                                # account type only for accounts
    assert n == 4 and len(sv.calls) == 2                                  # 4 names, chunk 2: 2 grids, both columns
    assert sv.calls[0][3] == ["ac.C70", "ac.C30"] and sv.calls[0][1]["All Time Periods"] == "Yr 25"


def test_a_shared_account_is_typed_over_every_hierarchy(tmp_path):
    """ac.STA20Y's first parent by atlas order is ac.Export (real atlas 10.11); its type comes from ac.STA."""
    A = "All Accounts"
    db = _atlas(tmp_path, [
        ("FinMax", A, A, A, None, 0, 0),
        ("FinMax", A, "ac.Export", "x", A, 1, 1),
        ("FinMax", A, "ac.STA20Y", "fte", "ac.Export", 2, 2),
        ("FinMax", A, "ac.Statistics", "s", A, 1, 3),
        ("FinMax", A, "ac.STA", "s", "ac.Statistics", 2, 4),
        ("FinMax", A, "ac.STA20Y", "fte", "ac.STA", 3, 5)])
    sv = FakeSv({("Yr 25", "ac.STA20Y"): 2782.0, ("Yr 25", "ac.STA"): 2782.0})
    cards_fetch.derive_dimension(sv, db, "FinMax", A, year=2025)
    types = {r[0] for r in db.execute("SELECT account_type FROM card WHERE name='ac.STA20Y'")}
    assert types == {"statistic"}
    assert db.execute("SELECT operator FROM card WHERE name='ac.STA20Y' AND parent='ac.STA'").fetchone()[0] == "+"


def test_a_noaccess_parent_is_unknown_not_a_zero_parent(tmp_path):
    E = "All Entities"
    db = _atlas(tmp_path, [("FinMax", E, "P", "P", E, 1, 1), ("FinMax", E, "A", "A", "P", 2, 2),
                           ("FinMax", E, "B", "B", "P", 2, 3)])
    sv = FakeSv({("ac.C70", "P"): "#NoAccess", ("ac.C70", "A"): 4.0, ("ac.C70", "B"): 1.0})
    cards_fetch.derive_dimension(sv, db, "FinMax", E, year=2025)
    assert {r[0] for r in db.execute("SELECT operator FROM card WHERE parent='P'")} == {None}


def test_finplan_accounts_take_a_second_scenario_for_what_actual_lacks(tmp_path):
    db = _atlas(tmp_path, [("FinPlan", "Account", "Statistics", "s", "Account", 1, 1),
                           ("FinPlan", "Account", "Subs", "s", "Statistics", 2, 2)])

    def val(pov, col, name):                       # only Target has data
        return {"Statistics": 9.0, "Subs": 9.0}[name] if pov["Scenario"] == "Target" else "#Missing"
    sv = FakeSv(val)
    cards_fetch.derive_dimension(sv, db, "FinPlan", "Account", year=2025)
    assert [(c[1]["Scenario"], c[1]["Years"], c[3]) for c in sv.calls] == [
        ("Actual", "FY25", ["YearTotal"]), ("Target", "FY26", ["YearTotal"])]
    assert db.execute("SELECT operator, account_type FROM card WHERE name='Subs'").fetchone() == ("+", "statistic")


def test_finlite_copies_finmax_cards_where_its_tree_is_the_same(tmp_path):
    E = "All Entities"
    db = _atlas(tmp_path, [("FinLite", E, "A", "A", "P", 2, 1), ("FinLite", E, "B", "B", "Q", 2, 2)])
    atlas.store_cards(db, "FinMax", E, [("A", "P", "+", None, None, 0), ("B", "P", "~", None, None, 0)])
    atlas.store_cards(db, "FinLite", E, [("old", "P", "+", None, None, 0)])
    assert cards_fetch.copy_finlite(db) == 1
    assert db.execute("SELECT name, parent, operator FROM card WHERE connection='FinLite'").fetchall() == [
        ("A", "P", "+")]


def test_cards_status_counts(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    atlas.store_cards(db, "FinMax", "All Entities", [("A", "P", "+", None, None, 0), ("B", "P", None, None, None, 0)])
    st = atlas.cards_status(db)
    assert st[0]["rows"] == 2 and st[0]["operator"] == {"+": 1, "unknown": 1}


def test_hidden_children_allow_only_plus_for_the_visible_ones():
    kids = {"P": ["A", "B", "H"]}
    got = operators(kids, {"P": 5.0, "A": 2.0, "B": 3.0}, blocked=[{"H"}])
    assert got == {("P", "A"): "+", ("P", "B"): "+", ("P", "H"): None}
    # real case 10.11 (FinMax ac.C70 Yr 25): un.L9 is empty, un.L9.C #NoAccess, un.L9.L9i -48 292. Read as 0 the
    # probe called L9i '~' (a zero parent); with the hidden child nothing can be said.
    real = {"un.L9": ["un.L9.C", "un.L9.506010", "un.L9.L9i"]}
    got = operators(real, {"un.L9.L9i": -48292.28}, blocked=[{"un.L9.C"}])
    assert set(got.values()) == {None}


def test_the_nearest_branch_root_decides_the_account_type():
    """Live atlas 10.11: the FinMax P&L is shared into equity (ac.C89T -> ac.C203 -> ... -> ac.C1_C2), so ac.C70
    and ac.C30 have ac.C1_C2 above them too; the nearest root decides."""
    c70 = ["ac.C70", "ac.EBITDA_BOI", "ac.EBITDA", "ac.EBIT", "ac.C85", "ac.C87", "ac.C89T", "ac.C203", "ac.C3_C8",
           "ac.C207", "ac.C208", "ac.C20S", "ac.C2S", "ac.C1_C2"]
    assert account_type(c70, -1.0) == "expense"
    assert account_type({"ac.C85": 0, "ac.C87": 1, "ac.C89T": 2, "ac.C3_C8": 3, "ac.C203": 3, "ac.C1_C2": 8},
                        -4.0) == "expense"                                     # P&L sign, not balance
    assert account_type({"ac.C181": 0, "ac.C18S": 1, "ac.C1_C2": 4, "ac.Export": 1}, 3.0) == "balance"
    assert cards_fetch.ancestry("x", {"x": ["e", "s"], "e": ["root"], "s": ["root"]}) == {"x": 0, "e": 1, "s": 1,
                                                                                          "root": 2}


def test_a_member_whose_name_clashes_with_an_alias_is_found_and_left_out():
    class Clash(FakeSv):
        def fetch(self, conn, pov, rows, cols):
            if "SDx Revenue" in rows[0].members:
                self.calls.append(("fail", len(rows[0].members)))
                raise RuntimeError("Smart View provider: Found duplicate aliases or member name conflicting with "
                                   "alias under parent 'ac.Statistics'.")
            return super().fetch(conn, pov, rows, cols)
    names = ["A", "B", "SDx Revenue", "C", "D"] + [f"m{i}" for i in range(40)]   # 1 clash of 45: under the cap
    sv = Clash({("YearTotal", n): 1.0 for n in names})
    vals, blocked = cards_fetch.grid(sv, "FinPlan", {}, "Account", names, "Period", ["YearTotal"])
    assert blocked == {"YearTotal": {"SDx Revenue"}}
    assert vals["A"]["YearTotal"] == 1.0 and vals["D"]["YearTotal"] == 1.0 and vals["SDx Revenue"]["YearTotal"] is None
    import pytest
    with pytest.raises(RuntimeError, match="other"):
        cards_fetch.grid(type("S", (), {"fetch": lambda *a: (_ for _ in ()).throw(RuntimeError("other"))})(),
                         "FinPlan", {}, "Account", names, "Period", ["YearTotal"])
    known = {"SDx Revenue"}
    sv.calls.clear()
    cards_fetch.grid(sv, "FinPlan", {}, "Account", names, "Period", ["YearTotal"], clashes=known)
    assert not [c for c in sv.calls if c[0] == "fail"]                  # a known clash is not searched again


def test_a_mixed_pnl_total_has_no_account_type():
    from smartass.knowledge.cards import account_types
    kids = {"ac.C3_C8": ["ac.EBITDA"], "ac.EBITDA": ["ac.C30", "ac.C70"], "ac.C70": ["ac.X"]}
    anc = {"ac.EBITDA": {"ac.EBITDA": 0, "ac.C3_C8": 1}, "ac.C30": {"ac.C30": 0, "ac.EBITDA": 1, "ac.C3_C8": 2},
           "ac.C70": {"ac.C70": 0, "ac.EBITDA": 1, "ac.C3_C8": 2},
           "ac.X": {"ac.X": 0, "ac.C70": 1, "ac.EBITDA": 2, "ac.C3_C8": 3}}
    got = account_types(kids, anc, {"ac.EBITDA": 5.0, "ac.C30": 9.0, "ac.C70": -4.0, "ac.X": -4.0})
    assert got == {"ac.EBITDA": None, "ac.C30": "revenue", "ac.C70": "expense", "ac.X": "expense"}


def test_derive_all_runs_only_the_wanted_cubes_and_copies_finlite(tmp_path):
    E = "All Entities"
    db = _atlas(tmp_path, [("FinMax", E, "A", "A", "P", 2, 1), ("FinLite", E, "A", "A", "P", 2, 1)])
    atlas.store_cards(db, "FinMax", E, [("A", "P", "+", None, None, 0)])
    sv = FakeSv({})
    done = cards_fetch.derive_all(sv, db, ["FinLite"], progress=lambda m: None)
    assert sv.calls == [] and [(d["dimension"], d["rows"]) for d in done] == [("(copy of FinMax)", 1)]


def test_derive_all_records_a_failing_dimension_and_goes_on(tmp_path):
    db = _atlas(tmp_path, [("FinPlan", "Entity", "A", "A", "P", 2, 1), ("FinPlan", "Account", "X", "X", "Y", 2, 1)])

    class Boom(FakeSv):
        def fetch(self, conn, pov, rows, cols):
            if rows[0].dimension == "Account":
                raise RuntimeError("down")
            return super().fetch(conn, pov, rows, cols)
    done = cards_fetch.derive_all(Boom({("OPEX", "A"): 1.0}), db, ["FinPlan"], progress=lambda m: None, year=2025)
    assert done[0] == {"connection": "FinPlan", "dimension": "Account", "error": "down"}
    assert done[1]["dimension"] == "Entity" and done[1]["rows"] == 1


def test_atlas_build_derives_cards_for_the_cubes_built_and_never_fails_on_them():
    from smartass.ops import atlas_build
    seen, msgs = [], []
    done = [{"connection": "FinMax", "dimension": "All Accounts", "members": 3},
            {"connection": "FinPlan", "dimension": "Account", "error": "x"}, {"connection": "FinLite", "error": "y"}]
    atlas_build.cards_after_build(done, derive=lambda cubes: seen.append(cubes) or [], progress=msgs.append)
    assert seen == [["FinMax"]]

    def boom(cubes):
        raise RuntimeError("no browser")
    assert atlas_build.cards_after_build(done, derive=boom, progress=msgs.append) == []
    assert "no browser" in msgs[-1] and msgs[-1].startswith("warning")
    assert atlas_build.cards_after_build([{"connection": "FinMax", "error": "z"}], derive=boom) == []


def test_cli_atlas_cards(monkeypatch, capsys):
    from smartass.ops import cli
    got = []
    monkeypatch.setattr(cards_fetch, "run", lambda cubes: got.append(cubes) or [
        {"connection": "FinMax", "dimension": "All Accounts", "rows": 3}])
    monkeypatch.setattr("smartass.core.logsetup.configure", lambda *a, **k: None)
    assert cli.main(["atlas", "cards", "--cube", "FinMax"]) == 0
    assert got == [["FinMax"]] and "cards: 1 derived, 0 failed" in capsys.readouterr().out


# --- Task 4: time balance per account -------------------------------------------------------------------------------

MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def test_derive_time_balance_finmax_from_one_periodic_grid(tmp_path):
    A = "All Accounts"
    db = _atlas(tmp_path, [("FinMax", A, "ac.STA", "s", A, 1, 1), ("FinMax", A, "ac.STA20Y", "f", "ac.STA", 2, 2),
                           ("FinMax", A, "ac.STA20Y", "f", "ac.Export", 2, 3), ("FinMax", A, "ac.C70", "o", A, 1, 4)])
    atlas.store_cards(db, "FinMax", A, [("ac.STA20Y", "ac.STA", "+", None, "statistic", 1),
                                        ("ac.STA20Y", "ac.Export", None, None, "statistic", 1),
                                        ("ac.C70", A, None, None, "expense", 0)])
    sta = dict(zip([f"{m} 25" for m in MON], [2857.3, -5.92, -13.2, -17.38, 0.65, -2.55, -11.4, -3.1, -45.8, 20.2,
                                               7.44, -4.12]))
    c70 = {f"{m} 25": -100.0 for m in MON}

    def val(pov, col, name):
        series = {"ac.STA20Y": sta, "ac.C70": c70}.get(name)
        if series is None:
            return "#Missing"
        return sum(series.values()) if col == "Yr 25" else series[col]
    sv = FakeSv(val)
    n = cards_fetch.derive_time_balance(sv, db, "FinMax", year=2025)
    tb = dict(db.execute("SELECT name || '/' || parent, time_balance FROM card"))
    assert tb == {"ac.STA20Y/ac.STA": "balance_last", "ac.STA20Y/ac.Export": "balance_last", "ac.C70/All Accounts": "flow"}
    assert n == 2 and len(sv.calls) == 1
    assert sv.calls[0][3] == [f"{m} 25" for m in MON] + ["Yr 25"] and sv.calls[0][1]["All Balance Types"] == "Periodic"


def test_derive_time_balance_finplan_falls_back_to_target(tmp_path):
    db = _atlas(tmp_path, [("FinPlan", "Account", "Subs", "s", "Statistics", 2, 1),
                           ("FinPlan", "Account", "OPEX", "o", "EBITDA", 2, 2)])
    atlas.store_cards(db, "FinPlan", "Account", [("Subs", "Statistics", "+", None, "statistic", 0),
                                                 ("OPEX", "EBITDA", "+", None, "expense", 0)])
    subs = [700.0, 699.0, 698.0, 697.0, 696.0, 694.0, 692.0, 691.0, 690.0, 689.0, 688.0, 687.0]

    def val(pov, col, name):
        if name == "OPEX" and pov["Scenario"] == "Actual":
            return -1200.0 if col == "YearTotal" else -100.0
        if name == "Subs" and pov["Scenario"] == "Target":
            return subs[-1] if col == "YearTotal" else subs[MON.index(col)]
        return "#Missing"
    sv = FakeSv(val)
    cards_fetch.derive_time_balance(sv, db, "FinPlan", year=2025)
    assert dict(db.execute("SELECT name, time_balance FROM card")) == {"Subs": "balance_last", "OPEX": "flow"}
    assert [(c[1]["Scenario"], c[1]["Years"], c[3][-1]) for c in sv.calls] == [("Actual", "FY25", "YearTotal"),
                                                                              ("Target", "FY26", "YearTotal")]
    assert sv.calls[1][2] == ["Subs"]                                     # Target only for what Actual left open


def test_finlite_gets_the_time_balance_too(tmp_path):
    A = "All Accounts"
    db = _atlas(tmp_path, [("FinMax", A, "ac.C70", "o", A, 1, 1), ("FinLite", A, "ac.C70", "o", A, 1, 1),
                           ("FinMax", "All Entities", "un.SY", "s", "All Entities", 1, 1)])
    sv = FakeSv(lambda pov, col, name: -1200.0 if col == "Yr 25" else -100.0 if name == "ac.C70" else "#Missing")
    done = cards_fetch.derive_all(sv, db, ["FinMax", "FinLite"], progress=lambda m: None, year=2025)
    assert db.execute("SELECT time_balance FROM card WHERE connection='FinLite' AND name='ac.C70'").fetchone() == (
        "flow",)
    assert [d["dimension"] for d in done] == ["All Accounts", "All Entities", "All Accounts (time balance)",
                                              "(copy of FinMax)"]


def test_time_balance_real_finplan_cases_10_11():
    """FinPlan BANORWAY, Actual/Working FY25 and Target/Final FY26 (live 10.11): months then the YearTotal."""
    def tb(row):
        return time_balance(row[-1], row[:-1])
    fte_avg = [2909.3, 2868.38, 2855.18, 2802.8, 2803.45, 2800.9, 2789.5, 2786.4, 2740.6, 2760.8, 2768.24, 2764.12,
               2804.139166666667]
    fwa_arpu = [738.3996949471823, 738.7201471417136, 738.6076040360317, 739.2705336475208, 738.7440901416387,
                739.0299285822648, 741.0005219875544, 740.5686280958253, 740.120848743199, 739.8723807212073,
                739.7974594405429, 739.1417880826715, 739.4394687972793]
    assert tb(fte_avg) == "average" and tb(fwa_arpu) == "average"       # No of FTE Total Average; FWA ARPU Recurring
    ebitda_pct = [0.5384128146010413, 0.5369949760678753, 0.5342829897255181, 0.5400535998976133, 0.5292658734744092,
                  0.5409895477296583, 0.6016765204704133, 0.5490648543826423, 0.5213593509518368, 0.522311178105174,
                  0.5215563394845106, 0.5347901936009707, 0.5392363107319369]     # Target FY26
    assert tb(ebitda_pct) is None   # a ratio of the year totals: a flat tolerance of 1 called it balance_last, 1e-4
    #                                 relative called it average
    gross_add_acc = [500.0, 800.0, 1300.0, 1800.0, 2200.0, 2600.0, 3100.0, 3300.0, 3900.0, 4200.0, 4600.0, 4900.0,
                     4900.0]
    assert tb(gross_add_acc) == "balance_last"
    fte_moves = [2909.3, -40.92, -13.2, -52.38, 0.65, -2.55, -11.4, -3.1, -45.8, 20.2, 7.44, -4.12, -4.12]
    assert tb(fte_moves) == "balance_last"     # 'No of FTEs': movements in a TB-Last member - the cube's own rule
    cost_per_fte = [-111557.6218, -113076.5877, -116903.2866, -120624.0505, -118020.8526, -103613.4321, -67788.5337,
                    -103330.5642, -128445.2749, -117952.3016, -123580.4877, -106543.6777, -1331436.6711]
    assert tb(cost_per_fte) == "flow"
    assert tb([739.0] * 13) is None                                      # flat: last = mean, undecidable


def test_pnl_accounts_are_flows_whatever_the_numbers_say():
    from smartass.knowledge.cards import settle_time_balance
    ac31204 = [3835408.0, 2102.03, None, None, None, 749.0, 749.0, 948.0, 948.0, 749.0, 749.0, 749.0]  # FinMax 10.11
    raw = time_balance(3843151.03, ac31204, movements=True)
    assert raw == "balance_last" and settle_time_balance(raw, "revenue") == "flow"
    assert settle_time_balance(None, "expense") == "flow" and settle_time_balance("balance_last", "statistic") == \
        "balance_last" and settle_time_balance(None, None) is None


def test_small_values_are_not_equal_within_a_whole_unit():
    got = operators({"KPI": ["a", "b"]}, {"KPI": 0.5, "a": 0.3, "b": 0.4})   # ratios: 0.7 is not 0.5
    assert set(got.values()) == {None}
    got = operators({"P": ["a", "b"]}, {"P": 1e9 + 0.4, "a": 6e8, "b": 4e8})  # NOK rounding still passes
    assert set(got.values()) == {"+"}
    got = operators({"Z": ["a", "b"]}, {"Z": 5.7e-7, "a": 3e6, "b": -1e6})   # a zero parent (ac.C40510i 10.11)
    assert set(got.values()) == {"~"}


def test_monthly_levels_that_the_cube_sums_read_as_flow():
    """Known limit (live 10.11): FinMax ac.STA21Y_stage stores the FTE level in every month and Yr sums them
    (33 725.66). From the numbers this is a flow like any steady monthly series; a manual card (Task 5) corrects it."""
    stage = [2857.29, 2851.37, 2838.17, 2820.79, 2821.44, 2818.89, 2807.49, 2804.4, 2758.59, 2778.8, 2786.23, 2782.12]
    assert time_balance(sum(stage), stage, movements=True) == "flow"


def test_a_dimension_without_members_keeps_its_cards(tmp_path):
    """A failed atlas dump stores the dimension with no members: deriving it must not wipe the last good cards."""
    E = "All Entities"
    db = _atlas(tmp_path, [("FinMax", "All Accounts", "ac.C70", "o", "All Accounts", 1, 1),
                           ("FinLite", E, "A", "A", "P", 2, 1)])
    atlas.store_cards(db, "FinMax", E, [("A", "P", "+", None, None, 0)])
    sv = FakeSv(lambda pov, col, name: -1200.0 if col == "Yr 25" else -100.0)
    done = cards_fetch.derive_all(sv, db, ["FinMax", "FinLite"], progress=lambda m: None, year=2025)
    assert "no FinMax/All Entities members" in next(d["error"] for d in done if d["dimension"] == E)
    assert db.execute("SELECT connection, operator FROM card WHERE name='A' ORDER BY 1").fetchall() == [
        ("FinLite", "+"), ("FinMax", "+")]


# --- Review fixes (rv_probe.py cases and I1-M10) ---------------------------------------------------------------------

CLASH = "Duplicate name Revenue found in grid; cannot resolve. Please replace it with qualified name."


def test_a_clash_in_the_column_or_pov_is_raised_not_searched():
    """rv_probe case 4: the clash is a column member, so every row 'clashes'; the search cost 2N-1 fetches and
    blocked every row. One retry with the dimension's own top member shows the rows are not the problem."""
    import pytest

    class ColClash(FakeSv):
        def fetch(self, conn, pov, rows, cols):
            self.calls.append(1)
            raise RuntimeError(CLASH)
    sv = ColClash({})
    with pytest.raises(RuntimeError, match="qualified name"):
        cards_fetch.grid(sv, "FinPlan", {}, "Entity", [f"m{i}" for i in range(100)], "Account", ["OPEX", "Revenue"],
                         chunk=50)
    assert len(sv.calls) == 2


def test_too_many_clashes_are_raised():
    import pytest

    class Many(FakeSv):
        def fetch(self, conn, pov, rows, cols):
            if any(n.startswith("bad") for n in rows[0].members):
                raise RuntimeError(CLASH)
            return super().fetch(conn, pov, rows, cols)
    names = [f"ok{i}" for i in range(90)] + [f"bad{i}" for i in range(10)]          # 10 % clash
    with pytest.raises(ValueError, match="clash"):
        cards_fetch.grid(Many({}), "FinPlan", {}, "Account", names, "Period", ["YearTotal"])


def test_a_dimension_without_data_in_any_slice_keeps_its_cards(tmp_path):
    """Every cell #NoAccess or #Missing (an expired session, a lost access right): no new cards. Measured live
    10.11: FinMax entities have 94-97 % of all names without a number (16 578 #NoAccess), but 76 % / 87 % of the
    visible ones, so the guard counts visible names only."""
    import pytest
    E = "All Entities"
    db = _atlas(tmp_path, [("FinMax", E, "P", "P", E, 1, 1), ("FinMax", E, "A", "A", "P", 2, 2)])
    atlas.store_cards(db, "FinMax", E, [("A", "P", "+", None, None, 0)])
    for cell in ("#NoAccess", "#Missing"):
        with pytest.raises(ValueError, match="no data"):
            cards_fetch.derive_dimension(FakeSv(lambda pov, col, name, c=cell: c), db, "FinMax", E, year=2025)
    assert db.execute("SELECT operator FROM card WHERE name='A'").fetchone() == ("+",)


def test_mostly_hidden_members_still_derive_when_the_visible_ones_have_data(tmp_path):
    E = "All Entities"
    db = atlas.connect(tmp_path / "b.db")
    db.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)", [("FinMax", E, f"m{i}", "m", "P", 2, i)
                                                                 for i in range(20)])
    db.commit()
    cells = {f"m{i}": v for i, v in enumerate(["#NoAccess"] * 18 + [1.0, "#Missing"])}   # 90 % hidden
    got = cards_fetch.derive_dimension(FakeSv(lambda pov, col, name: cells[name]), db, "FinMax", E, year=2025)
    assert got["rows"] == 20


def test_rederiving_operators_keeps_the_time_balance_until_a_new_one_is_derived(tmp_path):
    """rv_probe case 3: the time balance step fails (SSO) after the operators were stored: the old time balance
    stays; the run stops at the sign-in error (M6)."""
    A = "All Accounts"
    db = _atlas(tmp_path, [("FinMax", A, "ac.STA", "s", A, 1, 1), ("FinMax", A, "ac.STA20Y", "f", "ac.STA", 2, 2)])
    atlas.store_cards(db, "FinMax", A, [("ac.STA20Y", "ac.STA", "+", "balance_last", "statistic", 0),
                                        ("ac.STA", A, "+", "balance_last", "statistic", 0)])

    class Half(FakeSv):
        def fetch(self, conn, pov, rows, cols):
            if len(cols[0].members) > 2:
                raise RuntimeError("SSO expired")
            return super().fetch(conn, pov, rows, cols)
    done = cards_fetch.derive_all(Half({("Yr 25", "ac.STA20Y"): 5.0, ("Yr 25", "ac.STA"): 5.0}), db, ["FinMax"],
                                  progress=lambda m: None, year=2025)
    assert done[-1]["error"] == "SSO expired" and done[-1]["dimension"] == "All Accounts (time balance)"
    assert not any(d["dimension"] == "(copy of FinMax)" for d in done)            # stopped
    assert set(db.execute("SELECT name, time_balance FROM card WHERE connection='FinMax'")) == {
        ("ac.STA20Y", "balance_last"), ("ac.STA", "balance_last")}


def test_store_cards_carries_the_time_balance_of_each_row():
    import sqlite3
    db = sqlite3.connect(":memory:")
    db.executescript(atlas.SCHEMA)
    atlas.store_cards(db, "FinMax", "A", [("x", "p", "+", "flow", None, 0), ("y", "p", "+", "flow", None, 0)])
    atlas.store_cards(db, "FinMax", "A", [("x", "p", "~", None, None, 0), ("z", "p", "+", None, None, 0),
                                          ("y", "p", "+", "balance_last", None, 0)])
    assert set(db.execute("SELECT name, operator, time_balance FROM card")) == {
        ("x", "~", "flow"), ("z", "+", None), ("y", "+", "balance_last")}


def test_finlite_copy_skips_a_dimension_whose_finlite_atlas_failed(tmp_path):
    """rv_probe case 1: FinLite's All Entities dump failed (no members): its cards stay; All Accounts is copied."""
    db = _atlas(tmp_path, [("FinLite", "All Accounts", "X", "X", "P", 1, 1)])
    atlas.store_cards(db, "FinMax", "All Entities", [("A", "P", "+", None, None, 0)])
    atlas.store_cards(db, "FinMax", "All Accounts", [("X", "P", "+", "flow", "expense", 0)])
    atlas.store_cards(db, "FinLite", "All Entities", [("A", "P", "+", None, None, 0)])
    assert cards_fetch.copy_finlite(db) == 1
    assert set(db.execute("SELECT dimension, name FROM card WHERE connection='FinLite'")) == {
        ("All Entities", "A"), ("All Accounts", "X")}
    # a build error on the FinLite dimension keeps its cards too, even with members left
    db.execute("INSERT INTO build VALUES ('FinLite', 'All Accounts', 0, 1, 0, 'timed out')")
    db.commit()
    atlas.store_cards(db, "FinMax", "All Accounts", [("X", "P", "~", None, None, 0)])
    cards_fetch.copy_finlite(db)
    assert db.execute("SELECT operator FROM card WHERE connection='FinLite' AND name='X'").fetchone() == ("+",)


def test_a_coincidental_plus_contradicted_by_another_slice_is_reset():
    """rv_probe case 2: slice 1 (P 0 = 4 - 4) says '+'; slice 2 (P 0, children 6 and -2) contradicts it."""
    from smartass.knowledge.cards import contradicted
    kids = {"P": ["A", "B"], "Q": ["C"]}
    sl = [{"P": 0.0, "A": 4.0, "B": -4.0, "Q": 1.0, "C": 1.0}, {"P": 0.0, "A": 6.0, "B": -2.0, "Q": 2.0, "C": 2.0}]
    bad: set = set()
    ops = operators(kids, *sl, conflicts=bad)
    assert ops == {("P", "A"): None, ("P", "B"): None, ("Q", "C"): "+"} and bad == {"P"}
    first = {("P", "A"): "+", ("P", "B"): "+", ("Q", "C"): "+"}           # what slice 1 alone gave
    assert contradicted(kids, first, sl) == {"P"}
    assert contradicted(kids, first, sl, blocked=[set(), {"B"}]) == set()      # B hidden in slice 2: no check


def test_derive_dimension_resets_and_counts_conflicts(tmp_path):
    E = "All Entities"
    db = _atlas(tmp_path, [("FinMax", E, "P", "P", E, 1, 1), ("FinMax", E, "A", "A", "P", 2, 2),
                           ("FinMax", E, "B", "B", "P", 2, 3)])
    vals = {("ac.C70", "P"): 0.0, ("ac.C70", "A"): 4.0, ("ac.C70", "B"): -4.0,
            ("ac.C30", "P"): 0.0, ("ac.C30", "A"): 6.0, ("ac.C30", "B"): -2.0}
    got = cards_fetch.derive_dimension(FakeSv(vals), db, "FinMax", E, year=2025)
    assert got["conflicts"] == 1
    assert {r[0] for r in db.execute("SELECT operator FROM card WHERE parent='P'")} == {None}


def test_copy_failure_is_a_step_error_and_a_sign_in_error_stops_the_run(tmp_path, monkeypatch):
    E = "All Entities"
    db = _atlas(tmp_path, [("FinMax", "All Accounts", "ac.C70", "o", "All Accounts", 1, 1),
                           ("FinMax", E, "un.SY", "s", E, 1, 1)])

    def boom(db):
        raise RuntimeError("disk full")
    monkeypatch.setattr(cards_fetch, "copy_finlite", boom)
    sv = FakeSv(lambda pov, col, name: -1.0)
    done = cards_fetch.derive_all(sv, db, ["FinMax"], progress=lambda m: None, year=2025)
    assert done[-1] == {"connection": "FinLite", "dimension": "(copy of FinMax)", "error": "disk full"}

    class NoBrowser(FakeSv):
        def fetch(self, conn, pov, rows, cols):
            raise RuntimeError("No headless browser could start for the drill")
    done = cards_fetch.derive_all(NoBrowser({}), db, None, progress=lambda m: None, year=2025)
    assert len(done) == 1 and "browser" in done[0]["error"]


def test_the_year_is_the_last_full_year_at_or_before_the_closed_month():
    class Sv:
        def __init__(self, m):
            self.m = m

        def closed_month(self, conn="FinMax"):
            return self.m
    assert cards_fetch.year_for(Sv("Oct 26")) == 2025
    assert cards_fetch.year_for(Sv("Jan 26")) == 2025
    assert cards_fetch.year_for(Sv("Dec 25")) == 2025
    assert cards_fetch.year_for(Sv("Dec 26")) == 2026


# --- Task 5: manual overrides merged over the derived cards in knowledge.db ------------------------------------------

import sqlite3  # noqa: E402

from smartass.knowledge.build import build  # noqa: E402
from smartass.knowledge.resolver import KnowledgeResolver  # noqa: E402
from smartass.knowledge.semantic import CardOverride, SemanticError, load_all  # noqa: E402

STA_ATLAS = [("FinMax", "All Accounts", "ac.Statistics", "Statistics", "All Accounts", 2, 1),
             ("FinMax", "All Accounts", "ac.STA", "STA", "ac.Statistics", 3, 1),
             ("FinMax", "All Accounts", "ac.STA20Y", "STA20Y FTE", "ac.STA", 4, 1),
             ("FinMax", "All Accounts", "ac.STA20Y", "STA20Y FTE", "ac.Export", 3, 2),
             ("FinMax", "All Accounts", "ac.STA21Y_stage", "stage", "ac.Not_Mapped", 3, 3),
             ("FinLite", "All Accounts", "ac.STA21Y_stage", "stage", "ac.Not_Mapped", 3, 1)]
CARD_ROWS = [("FinMax", "All Accounts", "ac.STA", "ac.Statistics", "+", None, "statistic", 0, 1.0),
             ("FinMax", "All Accounts", "ac.STA20Y", "ac.STA", "+", "balance_last", "statistic", 1, 1.0),
             ("FinMax", "All Accounts", "ac.STA20Y", "ac.Export", None, "balance_last", "statistic", 1, 1.0),
             ("FinMax", "All Accounts", "ac.STA21Y_stage", "ac.Not_Mapped", "+", "flow", None, 0, 1.0),
             ("FinLite", "All Accounts", "ac.STA21Y_stage", "ac.Not_Mapped", "+", "flow", None, 0, 1.0)]


def _knowledge(tmp_path, cards_toml: dict[str, str] | None = None, cards=True):
    from kutil import copy_required_rules
    a, sem, k = tmp_path / "atlas.db", tmp_path / "semantic", tmp_path / "knowledge.db"
    con = atlas.connect(a)
    con.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)", STA_ATLAS)
    if cards:
        con.executemany("INSERT INTO card VALUES (?,?,?,?,?,?,?,?,?)", CARD_ROWS)
    con.commit()
    con.close()
    copy_required_rules(sem)
    for name, text in (cards_toml or {}).items():
        (sem / "cards").mkdir(exist_ok=True)
        (sem / "cards" / f"{name}.toml").write_text(text, encoding="utf-8")
    return a, sem, k


def _resolver(tmp_path, cards_toml=None, **kw):
    a, sem, k = _knowledge(tmp_path, cards_toml, **kw)
    report = build(sem, a, k)
    assert report.ok, report.errors
    return KnowledgeResolver(sqlite3.connect(k), sqlite3.connect(a), forecast_year=26), report


STA_TILDE = """
[[card]]
cube = "FinMax"
member = "ac.STA"
parent = "ac.Statistics"
operator = "~"
note = "statistics never add up to the P&L"
source = "probe 10.11"
"""


def test_card_override_model():
    CardOverride(cube="FinMax", member="un.SY.X", parent="un.SY", operator="~", note="elimination", source="s")
    with pytest.raises(ValueError):
        CardOverride(cube="FinMax", member="x", operator="*", source="s")
    with pytest.raises(ValueError):
        CardOverride(cube="FinMax", member="x", operator="+", source="s", colour="red")


def test_a_manual_card_wins_over_the_derived_one_field_by_field(tmp_path):
    r, _ = _resolver(tmp_path, {"x": STA_TILDE})
    c = r.card("ac.STA", "FinMax", "ac.Statistics")
    assert c["operator"] == "~" and c["provenance"]["operator"].startswith("manual@")
    assert c["account_type"] == "statistic" and c["provenance"]["account_type"].startswith("derived@")
    assert c["note"] == "statistics never add up to the P&L"


def test_without_an_override_the_derived_card_stands(tmp_path):
    r, _ = _resolver(tmp_path)
    c = r.card("ac.STA", "FinMax", "ac.Statistics")
    assert c["operator"] == "+" and c["provenance"]["operator"].startswith("derived@")
    assert r.card("ac.nothing", "FinMax") is None
    assert r.card("AC.sta20y", "FinMax", "ac.sta")["time_balance"] == "balance_last"     # any case
    assert r.card("ac.STA20Y", "FinMax", "ac.Export")["shared"] is True


def test_an_override_without_parent_covers_every_placement_and_finlite_shares_finmax(tmp_path):
    toml = '[[card]]\ncube = "FinMax"\nmember = "ac.STA21Y_stage"\ntime_balance = "balance_last"\nsource = "t"\n'
    r, _ = _resolver(tmp_path, {"x": toml})
    assert r.card("ac.STA21Y_stage", "FinMax")["time_balance"] == "balance_last"
    assert r.card("ac.STA21Y_stage", "FinLite", "ac.Not_Mapped")["time_balance"] == "balance_last"
    assert r.card("ac.STA21Y_stage", "FinMax")["provenance"]["time_balance"].startswith("manual@")


def test_a_specific_parent_beats_a_general_override(tmp_path):
    toml = ('[[card]]\ncube = "FinMax"\nmember = "ac.STA20Y"\noperator = "-"\nsource = "t"\n'
            '[[card]]\ncube = "FinMax"\nmember = "ac.STA20Y"\nparent = "ac.STA"\noperator = "~"\nsource = "t"\n')
    r, _ = _resolver(tmp_path, {"x": toml})
    assert r.card("ac.STA20Y", "FinMax", "ac.STA")["operator"] == "~"
    assert r.card("ac.STA20Y", "FinMax", "ac.Export")["operator"] == "-"


def test_conflicting_overrides_and_unknown_members_stop_the_build(tmp_path):
    toml = STA_TILDE + STA_TILDE.replace('"~"', '"-"')
    a, sem, k = _knowledge(tmp_path, {"x": toml})
    report = build(sem, a, k)
    assert not report.ok and any("ac.STA" in e and "operator" in e for e in report.errors)
    gone = '[[card]]\ncube = "FinMax"\nmember = "ac.Gone"\noperator = "~"\nsource = "t"\n'
    a2, sem2, k2 = _knowledge(tmp_path / "two", {"x": gone})
    report = build(sem2, a2, k2)
    assert not report.ok and any("ac.Gone" in e for e in report.errors)
    same = STA_TILDE + STA_TILDE                      # the same value twice is not a conflict
    a3, sem3, k3 = _knowledge(tmp_path / "three", {"x": same})
    assert build(sem3, a3, k3).ok


def test_a_bad_cards_file_is_a_semantic_error(tmp_path):
    from kutil import copy_required_rules
    copy_required_rules(tmp_path)
    (tmp_path / "cards").mkdir()
    (tmp_path / "cards" / "x.toml").write_text('[[card]]\ncube = "FinMax"\nmember = "a"\noperator = "*"\nsource="s"\n')
    with pytest.raises(SemanticError):
        load_all(tmp_path)


def test_an_atlas_without_cards_still_builds(tmp_path):
    r, report = _resolver(tmp_path, cards=False)
    assert r.card("ac.STA", "FinMax", "ac.Statistics") is None and report.ok


def test_the_repo_card_overrides_are_valid_and_cover_the_stage_statistics():
    from pathlib import Path
    k = load_all(Path(__file__).resolve().parents[3] / "knowledge" / "semantic")
    stage = [c for c in k.cards if c.member.endswith("_stage")]
    assert stage and all(c.time_balance == "balance_last" and c.cube == "FinMax" for c in stage)
    assert not any(c.member == "ac.STA" for c in k.cards)       # ac.STA under ac.Statistics is derived as '~'


# --- Task 6: the model consolidates with the real operators ----------------------------------------------------------

def _weight_resolver(tmp_path, ops):
    """FinMax entity un.SY with children Ai, Bi, Ci and one shared child under two parents; ops: {child: operator}."""
    a, sem, k = tmp_path / "atlas.db", tmp_path / "semantic", tmp_path / "knowledge.db"
    from kutil import copy_required_rules
    con = atlas.connect(a)
    rows = [("FinMax", "All Entities", "un.SY", "SY", "All Entities", 1, 1)]
    rows += [("FinMax", "All Entities", n, n, "un.SY", 2, i) for i, n in enumerate(["Ai", "Bi", "Ci"])]
    rows += [("FinMax", "All Entities", "Ci", "Ci", "un.SY.Other", 3, 9)]
    con.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)", rows)
    con.executemany("INSERT INTO card VALUES (?,?,?,?,?,?,?,?,?)",
                    [("FinMax", "All Entities", n, p, op, None, None, int(n == "Ci"), 1.0) for (n, p), op in ops.items()])
    con.commit()
    con.close()
    copy_required_rules(sem)
    assert build(sem, a, k).ok
    return KnowledgeResolver(sqlite3.connect(k), sqlite3.connect(a), forecast_year=26)


def test_weight_follows_the_card_operator_per_parent(tmp_path):
    r = _weight_resolver(tmp_path, {("Ai", "un.SY"): "~", ("Bi", "un.SY"): "-", ("Ci", "un.SY"): "+",
                                    ("Ci", "un.SY.Other"): "~"})
    assert r.weight("Ai", "FinMax", parent="un.SY") == 0.0
    assert r.weight("Bi", "FinMax", parent="un.SY") == -1.0
    assert r.weight("Ci", "FinMax", parent="un.SY") == 1.0
    assert r.weight("Ci", "FinMax", parent="un.SY.Other") == 0.0           # per placement of a shared member
    assert r.weight("Ai", "FinMax", parent="Elsewhere") == 1.0              # no card for this pair
    assert r.weight("Ai", "FinMax") == 0.0                                  # no parent: the shallowest placement
    assert r.weight("Nobody", "FinMax", parent="un.SY") == 1.0
    assert r.weight("Ai", "FinMax", parent="un.sy") == 0.0                  # any case


def test_an_unknown_operator_weighs_one(tmp_path):
    r = _weight_resolver(tmp_path, {("Ai", "un.SY"): None})
    assert r.weight("Ai", "FinMax", parent="un.SY") == 1.0


def test_the_blocks_pass_the_entity_as_the_parent():
    from v2util import CLOSED, spec_dict

    from smartass.model.blocks.common import BlockOut, ck, counted, entity_lines, evaluate
    from smartass.model.context import Ctx
    from smartass.model.frame import Fact, Frame
    from smartass.model.resolver import MetricDef, StaticResolver
    from smartass.model.spec import AnalysisSpec

    class Cards(StaticResolver):
        def weight(self, member, cube, parent=None):
            return 0.0 if (member, parent) == ("B", "P") else 1.0       # B is eliminated in P only
    d = spec_dict(blocks=[{"type": "matrix", "query": "fm", "metric": "Revenue"}])
    d["queries"][0].update(entity="P", periods="Sep 26")
    r = Cards([MetricDef("Revenue", {"FinMax": ("ac.C30",)})], tree={"P": ("A", "B")})
    fr = Frame([Fact("FinMax", "fm", "ac.C30", e, "Sep 26", "Actual", "", v)
                for e, v in (("A", 100.0), ("B", 30.0), ("P", 100.0))])
    c = Ctx(AnalysisSpec.model_validate(d), r, CLOSED, fr)
    q = c.query("fm")
    assert counted(c, q, ["A", "B"]) == ["A"]
    lines = entity_lines(c, q, "Revenue", ["value"], "b0")
    out = BlockOut("b0", "table", "t", ["Actual"], lines=lines)
    assert evaluate([out], fr).get(ck("b0", "check", 0)) == 0.0


# --- Task 7: cards in fin_lookup and in caveats -----------------------------------------------------------------------

LOOKUP_ATLAS = [("FinMax", "All Accounts", "ac.Statistics", "Statistics", "All Accounts", 2, 1),
                ("FinMax", "All Accounts", "ac.STA", "STA statistics", "ac.Statistics", 3, 1),
                ("FinMax", "All Accounts", "ac.STA20Y", "STA20Y FTE", "ac.STA", 4, 1),
                ("FinMax", "All Accounts", "ac.STA21Y", "STA21Y FTE other", "ac.STA", 4, 2),
                ("FinMax", "All Accounts", "ac.STA22Y", "STA22Y FTE minus", "ac.STA", 4, 3),
                ("FinMax", "All Accounts", "ac.STA23Y", "STA23Y plain", "ac.STA", 4, 4),
                ("FinMax", "All Accounts", "ac.AVG1", "AVG1 average", "ac.Statistics", 3, 2)]
LOOKUP_CARDS = [("FinMax", "All Accounts", "ac.STA", "ac.Statistics", "+", "balance_last", "statistic", 0, 1.0),
                ("FinMax", "All Accounts", "ac.STA20Y", "ac.STA", "+", "balance_last", "statistic", 0, 1.0),
                ("FinMax", "All Accounts", "ac.STA21Y", "ac.STA", "~", "balance_last", "statistic", 0, 1.0),
                ("FinMax", "All Accounts", "ac.STA22Y", "ac.STA", "-", "balance_last", "statistic", 0, 1.0),
                ("FinMax", "All Accounts", "ac.STA23Y", "ac.STA", "+", None, None, 0, 1.0),
                ("FinMax", "All Accounts", "ac.AVG1", "ac.Statistics", "+", "average", "statistic", 0, 1.0)]
FTE_METRIC = '''name = "FTE"
description = "A level."
unit = "FTE"
balance = "Accumulated"
aggregation = "last"
status = "approved"
[finmax]
members = ["ac.STA20Y"]
'''


def _lookup(tmp_path):
    from kutil import copy_required_rules
    a, sem, k = tmp_path / "atlas.db", tmp_path / "semantic", tmp_path / "knowledge.db"
    con = atlas.connect(a)
    con.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)", LOOKUP_ATLAS)
    con.executemany("INSERT INTO card VALUES (?,?,?,?,?,?,?,?,?)", LOOKUP_CARDS)
    con.commit()
    con.close()
    copy_required_rules(sem)
    (sem / "metrics").mkdir()
    (sem / "metrics" / "fte.toml").write_text(FTE_METRIC, encoding="utf-8")
    assert build(sem, a, k).ok
    return KnowledgeResolver(sqlite3.connect(k), sqlite3.connect(a), forecast_year=26)


def test_member_hits_carry_a_compact_card(tmp_path):
    r = _lookup(tmp_path)
    hit = next(m for m in r.search("STA", kind="member", cube="FinMax")["members"] if m["name"] == "ac.STA")
    assert hit["card"] == {"time_balance": "balance_last", "account_type": "statistic",
                           "children_ignored": ["ac.STA21Y"], "children_subtracted": ["ac.STA22Y"]}
    plain = next(m for m in r.search("STA23Y", kind="member", cube="FinMax")["members"] if m["name"] == "ac.STA23Y")
    assert plain["card"] == {}                                    # nothing known: None fields are omitted
    root = r.search("Statistics", kind="member", cube="FinMax")["members"][0]
    assert root["name"] == "ac.Statistics" and "children_ignored" not in root["card"]


def test_a_level_gets_a_time_balance_caveat_unless_a_last_metric_covers_it(tmp_path):
    r = _lookup(tmp_path)
    got = r.card_caveats({"FinMax": {"ac.STA21Y", "ac.STA20Y", "ac.AVG1", "ac.STA23Y", "un.SY"}})
    assert [c["id"] for c in got] == ["card_ac.AVG1", "card_ac.STA21Y"]       # STA20Y is the FTE metric (last)
    lvl = next(c for c in got if c["id"] == "card_ac.STA21Y")
    assert lvl["severity"] == "info" and lvl["text"] == (
        "STA21Y FTE other is a level (time balance last): read the period-end value, never sum months.")
    assert "average" in next(c for c in got if c["id"] == "card_ac.AVG1")["text"]


def test_the_pipeline_adds_the_card_caveats_once(tmp_path):
    from types import SimpleNamespace

    from smartass import pipeline
    r = _lookup(tmp_path)
    req = SimpleNamespace(cube="FinLite", entities=["un.SY"], accounts=["ac.STA21Y"], products=[], periods=["Sep 26"])
    spec = SimpleNamespace(blocks=[], queries=[])
    got = pipeline.caveats(spec, [req, req], r)
    assert [c["id"] for c in got if c["id"].startswith("card_")] == ["card_ac.STA21Y"]


# --- review fixes: no atlas scans, new cards rebuild, ~ children out of shown lines, grouped caveats ------------------

def test_card_lookups_never_touch_the_atlas_tree(tmp_path, monkeypatch):
    import time
    r = _lookup(tmp_path)

    def boom(*a, **k):
        raise AssertionError("atlas parent() scan")
    monkeypatch.setattr(r, "parent", boom)
    touched = {"FinMax": {f"un.SY.E{i}" for i in range(500)} | {"ac.STA21Y"}}
    t = time.perf_counter()
    got = r.card_caveats(touched)
    assert time.perf_counter() - t < 0.2 and [c["id"] for c in got] == ["card_ac.STA21Y"]
    assert r.compact_card("ac.STA", "FinMax")["children_ignored"] == ["ac.STA21Y"]
    assert r.card("ac.STA21Y", "FinMax")["operator"] == "~"                   # no parent: read from knowledge.db


def test_a_new_card_derivation_rebuilds_knowledge_db(tmp_path, monkeypatch):
    from kutil import knowledge

    from smartass.knowledge.resolver import REBUILT_NOTE, ensure_built
    a, _, _k = knowledge(tmp_path, monkeypatch)
    con = sqlite3.connect(a)
    con.execute("DELETE FROM build")
    con.execute("INSERT INTO build VALUES ('FinMax','All Accounts',1790000000.0,9,1.0,NULL)")
    con.commit()
    notes: list[str] = []
    ensure_built(notes)
    notes.clear()
    ensure_built(notes)
    assert notes == []
    con.execute("INSERT INTO card VALUES ('FinMax','All Accounts','ac.C30','ac.C48','+',NULL,NULL,0,1790000900.0)")
    con.commit()
    ensure_built(notes)
    assert notes == [REBUILT_NOTE]
    from smartass.knowledge.build import summary
    assert summary(_k)["atlas_built_ts"] == "1790000000.0|1790000900.0"
    con.close()


def test_the_fte_metric_also_covers_finlite(tmp_path):
    r = _lookup(tmp_path)
    assert r.card_caveats({"FinLite": {"ac.STA20Y"}, "FinMax": {"ac.STA20Y"}}) == []


def test_many_levels_make_one_grouped_caveat(tmp_path):
    r = _lookup(tmp_path)
    r._levels |= {("FinMax", f"ac.x{i}") for i in range(5)}
    r.card = lambda m, c, parent=None: {"time_balance": "balance_last"}
    got = r.card_caveats({"FinMax": {f"ac.x{i}" for i in range(5)}})
    assert [c["id"] for c in got] == ["card_levels"] and got[0]["text"].startswith("5 accounts are levels or averages:")
    assert "and 2 more" in got[0]["text"]


def test_the_average_caveat_says_the_average(tmp_path):
    got = _lookup(tmp_path).card_caveats({"FinMax": {"ac.AVG1"}})
    assert got[0]["text"] == "AVG1 average is an average (time balance average): read the period's average, never sum months."


def test_an_own_cube_override_beats_a_parented_finmax_one_on_finlite_rows(tmp_path):
    toml = ('[[card]]\ncube = "FinMax"\nmember = "ac.STA21Y_stage"\nparent = "ac.Not_Mapped"\n'
            'time_balance = "average"\nsource = "t"\n'
            '[[card]]\ncube = "FinLite"\nmember = "ac.STA21Y_stage"\ntime_balance = "balance_last"\nsource = "t"\n')
    r, _ = _resolver(tmp_path, {"x": toml})
    assert r.card("ac.STA21Y_stage", "FinLite", "ac.Not_Mapped")["time_balance"] == "balance_last"
    assert r.card("ac.STA21Y_stage", "FinMax", "ac.Not_Mapped")["time_balance"] == "average"


def _tilde_ctx(block):
    from v2util import CLOSED, spec_dict

    from smartass.model.context import Ctx
    from smartass.model.frame import Fact, Frame
    from smartass.model.resolver import MetricDef, StaticResolver
    from smartass.model.spec import AnalysisSpec
    d = spec_dict(blocks=[block])
    d["queries"][0].update(entity="P", periods="Sep 26")
    r = StaticResolver([MetricDef("Revenue", {"FinMax": ("ac.C30",)})], tree={"P": ("A", "B", "T")},
                       weights={"T": 0.0})
    fr = Frame([Fact("FinMax", "fm", "ac.C30", e, p, "Actual", "", v) for e, v in
                (("A", 100.0), ("B", 30.0), ("T", 500.0), ("P", 130.0)) for p in ("Sep 26", "Sep 25")])
    return Ctx(AnalysisSpec.model_validate(d), r, CLOSED, fr)


def test_a_tilde_child_is_not_shown_ranked_or_in_other():
    from smartass.model.blocks.common import ck, entity_lines
    c = _tilde_ctx({"type": "matrix", "query": "fm", "metric": "Revenue"})
    lines = entity_lines(c, c.query("fm"), "Revenue", ["value"], "b0")
    assert [ln.entity for ln in lines if ln.kind not in ("grand", "check")] == ["A", "B"]
    ranked = entity_lines(c, c.query("fm"), "Revenue", ["value"], "b0", top=1, sort="desc")
    assert [ln.entity for ln in ranked if ln.key.startswith("e")] == ["A"]
    other = next(ln for ln in ranked if ln.key == "other")
    from smartass.model.blocks.common import BlockOut, evaluate
    cells = evaluate([BlockOut("b0", "table", "t", ["Actual"], lines=ranked)], c.frame)
    assert cells.get(ck("b0", "other", 0)) == 30.0 and other


def test_a_tilde_child_is_not_a_driver():
    from smartass.model.blocks.drivers import build_drivers
    c = _tilde_ctx({"type": "drivers", "query": "fm", "metric": "Revenue", "base": "ly"})
    out = build_drivers(c.spec.blocks[0], c, "b0")
    assert "T" not in [ln.entity for ln in out.lines]


# --- Task 8: new metrics, verified live ---------------------------------------------------------------------------

def test_a_metric_verified_stamp_round_trips_through_its_toml(tmp_path):
    from smartass.knowledge.semantic import Metric, load_semantic, to_toml, terms
    m = Metric(id="ebit_x", name="EBIT X", aliases=["driftsresultat x"], description="d", status="approved",
               approved={"by": "auto", "on": "2026-11-10", "via": "live check phase B"},
               verified={"by": "auto", "on": "2026-11-10"}, finmax={"members": ["ac.EBIT"]})
    (tmp_path / "metrics").mkdir()
    (tmp_path / "metrics" / "ebit_x.toml").write_text(to_toml(m), encoding="utf-8")
    [back], _ = load_semantic(tmp_path)
    assert back == m and back.verified == {"by": "auto", "on": "2026-11-10"}
    assert terms(back) == ["EBIT X", "driftsresultat x"]


PHASE_B = ("ebit", "depreciation_writedowns", "profit_before_tax", "net_financial_items", "net_income", "capex",
           "total_assets", "subscriptions_fixed", "subscriptions_internet", "subscriptions_tv",
           "subscriptions_fixed_phone", "arpu_internet", "arpu_tv", "fte_finplan")


def test_the_phase_b_metrics_are_verified_and_read_levels_at_the_period_end():
    from smartass.core import config
    from smartass.knowledge.semantic import load_semantic
    metrics = {m.id: m for m in load_semantic(config.SEMANTIC_DIR)[0]}
    for mid in PHASE_B:
        m = metrics[mid]
        assert m.status == "approved" and m.verified and m.approved["by"] == "auto", mid
    assert (metrics["total_assets"].balance, metrics["total_assets"].aggregation) == ("Accumulated", "last")
    assert (metrics["fte_finplan"].balance, metrics["fte_finplan"].aggregation) == ("Accumulated", "last")
    for mid in PHASE_B:                 # FinMax statistics hold a level per month: Accumulated would add them up
        if mid.startswith(("subscriptions_", "arpu_")):
            assert (metrics[mid].balance, metrics[mid].aggregation) == ("Periodic", "last"), mid
            assert metrics[mid].pct_floor is not None and metrics[mid].materiality[0] < 1000, mid
    assert "FinPlan" not in metrics["net_income"].members()       # FinPlan Net Income has no tax (= PBT)
    assert metrics["capex"].sign == "neutral"                     # capex is stored positive
