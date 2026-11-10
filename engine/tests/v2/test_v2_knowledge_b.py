"""Knowledge layer phase B (docs/plans/2026-11-10-knowledge-b-cards-metrics.md)."""
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
    n = cards_fetch.derive_dimension(sv, db, "FinMax", "All Entities", chunk=2, year=2025)
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
    names = ["A", "B", "SDx Revenue", "C", "D"]
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
    assert sv.calls == [] and done == [{"connection": "FinLite", "dimension": "(copy of FinMax)", "rows": 1}]


def test_derive_all_records_a_failing_dimension_and_goes_on(tmp_path):
    db = _atlas(tmp_path, [("FinPlan", "Entity", "A", "A", "P", 2, 1), ("FinPlan", "Account", "X", "X", "Y", 2, 1)])

    class Boom(FakeSv):
        def fetch(self, conn, pov, rows, cols):
            if rows[0].dimension == "Account":
                raise RuntimeError("down")
            return super().fetch(conn, pov, rows, cols)
    done = cards_fetch.derive_all(Boom({}), db, ["FinPlan"], progress=lambda m: None)
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
