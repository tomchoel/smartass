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
