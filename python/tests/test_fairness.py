from roster.fairness import eligible_dates, ratio_less, candidate_order
from roster.model import M_ALL, M_DAY
from tests.helpers import make_inputs


def test_eligible_dates_excludes_exempt_and_full_day_leave_only():
    inp = make_inputs([("Ann", ""), ("Bob", "")], ["2027-01-01", "2027-02-06", "2027-02-07"],
                      exempt=[("Ann", 0)], leave=[("Ann", 1, M_ALL), ("Bob", 1, M_DAY)])
    assert eligible_dates(inp) == [1, 3]


def test_ratio_less_uses_cross_multiplication():
    assert ratio_less(1, 3, 1, 2) and not ratio_less(1, 2, 1, 3) and not ratio_less(2, 4, 1, 2)


def test_order_ratio_then_standby_then_name_zero_eligible_last():
    inp = make_inputs([("Dan", ""), ("bob", ""), ("Ann", ""), ("Cat", ""), ("Eve", "")],
                      ["2027-01-01"])
    #            Dan bob Ann Cat Eve
    active = [1, 0, 0, 0, 0]
    standby = [0, 1, 0, 0, 0]
    eligible = [4, 4, 4, 0, 4]
    assert [inp.staff[i] for i in candidate_order(inp, active, standby, eligible)] == \
        ["Ann", "Eve", "bob", "Dan", "Cat"]


def test_two_zero_eligible_ordered_by_standby_then_name():
    inp = make_inputs([("Zed", ""), ("Amy", ""), ("Bea", ""), ("Cal", "")], ["2027-01-01"])
    active = [0, 0, 0, 0]
    standby = [0, 2, 0, 1]
    eligible = [0, 0, 3, 3]
    # Bea (0/3, standby 0), Cal (0/3, standby 1), then zero-eligible Zed (standby 0), Amy (2)
    assert [inp.staff[i] for i in candidate_order(inp, active, standby, eligible)] == \
        ["Bea", "Cal", "Zed", "Amy"]


def test_zero_eligible_same_standby_falls_to_name():
    inp = make_inputs([("Zed", ""), ("amy", "")], ["2027-01-01"])
    assert candidate_order(inp, [5, 0], [1, 1], [0, 0]) == [1, 0]


def test_equal_ratio_different_denominators_fall_to_standby():
    inp = make_inputs([("Ann", ""), ("Bob", ""), ("Cat", "")], ["2027-01-01"])
    active = [2, 1, 0]
    standby = [0, 3, 5]
    eligible = [4, 2, 4]
    # Cat 0/4 first; Ann 2/4 and Bob 1/2 are equal, Ann has fewer standby
    assert candidate_order(inp, active, standby, eligible) == [2, 0, 1]


def test_name_tiebreak_uses_normalised_name():
    inp = make_inputs([(" bob ", ""), ("Ann", "")], ["2027-01-01"])
    assert candidate_order(inp, [0, 0], [0, 0], [2, 2]) == [1, 0]


def test_full_tie_keeps_sheet_order():
    inp = make_inputs([("Ann", ""), ("ANN", "")], ["2027-01-01"])
    assert candidate_order(inp, [0, 0], [0, 0], [2, 2]) == [0, 1]
