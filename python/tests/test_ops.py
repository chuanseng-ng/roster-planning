from roster.engine import new_result, solve_year
from roster.model import GAP, GAP_SENIOR, GAP_COAG, EMPTY_CELL
from roster.ops import (expand_weekends, roster_to_lastyear, approval_broken_rules,
                        apply_approval)
from roster.review import check_rules
from tests.helpers import make_inputs


def names(inp, res, d):
    out = []
    for v in res.grid[d]:
        out.append(inp.staff[v] if v >= 0 else {GAP: "GAP", GAP_SENIOR: "GAP:S", GAP_COAG: "GAP:C"}[v])
    return out


def solved(inp):
    res = new_result(inp)
    solve_year(inp, res, 0)
    return res


def dates_of(rows):
    return [r[1] for r in rows]


def test_friday_adds_saturday_and_sunday():
    rows, skipped = expand_weekends([["Good Friday", "2027-03-26", ""]], 2027)
    assert rows == [["Good Friday", "2027-03-26", ""], ["Good Friday", "2027-03-27", "Y"],
                    ["Good Friday", "2027-03-28", "Y"]]
    assert skipped == []


def test_monday_adds_saturday_and_sunday_before_sorted():
    rows, skipped = expand_weekends([["Bank", "2027-08-09", ""]], 2027)
    assert rows == [["Bank", "2027-08-07", "Y"], ["Bank", "2027-08-08", "Y"],
                    ["Bank", "2027-08-09", ""]]
    assert skipped == []


def test_saturday_and_sunday_add_only_monday_once():
    rows, skipped = expand_weekends([["CNY1", "2027-02-06", ""], ["CNY2", "2027-02-07", ""]], 2027)
    assert dates_of(rows) == ["2027-02-06", "2027-02-07", "2027-02-08"]
    assert rows[2] == ["CNY1", "2027-02-08", "Y"]
    assert skipped == []


def test_new_year_friday_and_saturday_christmas():
    rows, _ = expand_weekends([["NY", "2027-01-01", ""]], 2027)
    assert dates_of(rows) == ["2027-01-01", "2027-01-02", "2027-01-03"]
    rows, _ = expand_weekends([["Xmas", "2027-12-25", ""]], 2027)
    assert dates_of(rows) == ["2027-12-25", "2027-12-26", "2027-12-27"]
    assert rows[1:] == [["Xmas", "2027-12-26", "Y"], ["Xmas", "2027-12-27", "Y"]]


def test_dates_outside_year_are_skipped():
    rows, skipped = expand_weekends([["Eve", "2027-12-31", ""]], 2027)
    assert rows == [["Eve", "2027-12-31", ""]]
    assert skipped == ["2028-01-01", "2028-01-02"]


def test_skipped_date_listed_once():
    # Fri 2027-12-31 and Sat 2027-12-25 are unrelated; Sat 2028-01-01 style duplicate:
    rows, skipped = expand_weekends([["A", "2027-12-31", ""], ["B", "2027-12-31", ""]], 2027)
    assert skipped == ["2028-01-01", "2028-01-02"]
    assert len(rows) == 2


def test_idempotent_and_y_rows_not_expanded():
    first, _ = expand_weekends([["Good Friday", "2027-03-26", ""]], 2027)
    second, skipped = expand_weekends(first, 2027)
    assert second == first and skipped == []
    rows, _ = expand_weekends([["Good Friday", "2027-03-26", "y"]], 2027)
    assert rows == [["Good Friday", "2027-03-26", "y"]]


def test_expand_sorts_stably_by_date():
    rows, _ = expand_weekends([["Late", "2027-05-20", ""], ["B", "2027-03-10", ""],
                               ["A", "2027-03-10", ""]], 2027)
    assert rows == [["B", "2027-03-10", ""], ["A", "2027-03-10", ""], ["Late", "2027-05-20", ""]]


STAFF_A = [("S1", "SC"), ("S2", "SC")] + [("B%d" % i, "") for i in range(1, 7)]
STAFF_B = [("S1", "S"), ("S2", "S"), ("C1", "C"), ("C2", "C")] + [("B%d" % i, "") for i in range(1, 6)]
D3 = ["2027-01-01", "2027-03-10", "2027-05-20"]


def test_roster_to_lastyear_excludes_standby_and_gap():
    inp = make_inputs(STAFF_A, D3, exempt=[("S1", 0)], holiday=["NY", "Mid", "May"])
    res = solved(inp)
    rows = roster_to_lastyear(inp, res)
    assert rows[:5] == [["NY", "2027-01-01", "S2", "Day 1"], ["NY", "2027-01-01", "B1", "Day 2"],
                        ["NY", "2027-01-01", "B2", "Evening"], ["NY", "2027-01-01", "B3", "Night 1"],
                        ["Mid", "2027-03-10", "S1", "Day 1"]]
    assert len(rows) == 4 + 5 + 5
    for r in rows:
        assert not r[3].startswith("Standby")


def test_approve_exempt_senior_into_gap_senior():
    inp = make_inputs(STAFF_B, D3, exempt=[("S1", 0), ("S1", 1), ("S1", 2)])
    res = solved(inp)
    assert names(inp, res, 0) == ["S2", "C1", "B1", "C2", "GAP:S", "B2", "B3"]
    before = [list(r) for r in res.grid]
    approvals = []
    ok, msg = apply_approval(inp, res, approvals, 0, 4, 0)
    assert (ok, msg) == (True, "")
    assert [(a.d, a.k, a.s, a.rules) for a in approvals] == [(0, 4, 0, [4])]
    assert res.grid[0] == [before[0][0], before[0][1], before[0][2], before[0][3], 0,
                           before[0][5], before[0][6]]
    assert res.approved[0][4] is True
    assert res.active_count[0] == 1
    assert names(inp, res, 1)[4] == "GAP:S" and names(inp, res, 2)[4] == "GAP:S"
    assert check_rules(inp, res, approvals) == []


def test_approve_non_senior_into_gap_senior_records_rule_1():
    inp = make_inputs(STAFF_B, D3, exempt=[("S1", 0), ("S1", 1), ("S1", 2)])
    res = solved(inp)
    b1 = inp.staff.index("B4")       # not rostered on date 0
    assert b1 not in res.grid[0]
    approvals = []
    ok, msg = apply_approval(inp, res, approvals, 0, 4, b1)
    assert ok and msg == ""
    assert approvals[0].rules == [1]
    assert check_rules(inp, res, approvals) == []


def test_apply_approval_failures():
    inp = make_inputs(STAFF_A, D3, exempt=[("S1", 0)])
    res = solved(inp)
    approvals = []
    assert apply_approval(inp, res, approvals, 0, 0, 3) == (False, "slot is no longer a GAP")
    assert names(inp, res, 0)[4] == "GAP:S"
    assert apply_approval(inp, res, approvals, 0, 4, 0) == (False, "S1 already has a slot on this date")
    assert approvals == []


def test_later_approval_stays_after_earlier_approval():
    inp = make_inputs(STAFF_B, D3, exempt=[("S1", 0), ("S1", 1), ("S1", 2)])
    res = solved(inp)
    approvals = []
    assert apply_approval(inp, res, approvals, 1, 4, 0) == (True, "")
    assert res.grid[1][4] == 0
    assert apply_approval(inp, res, approvals, 0, 4, 0) == (True, "")
    assert res.grid[1][4] == 0 and res.approved[1][4]
    assert res.grid[0][4] == 0
    assert [(a.d, a.s) for a in approvals] == [(1, 0), (0, 0)]
    assert approvals[0].rules == [4]


def test_earlier_approval_merges_new_rule_into_later_one():
    staff = [("X", "S"), ("Y", "S")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-03-10", "2027-03-11"])
    res = solved(inp)
    res.grid[0][3] = GAP                   # hand-made GAP in Night 1 on the first date
    for k in range(len(inp.slots)):
        if res.grid[0][k] == 0:
            res.grid[0][k] = GAP           # free X on date 0
    res.grid[1][0] = 0                     # X already approved in Day 1 on date 1
    res.approved[1][0] = True
    from roster.review import Approval
    approvals = [Approval(1, 0, 0, [])]
    assert approval_broken_rules(inp, res, 0, 3, 0) == [7]
    ok, msg = apply_approval(inp, res, approvals, 0, 3, 0)
    assert ok
    assert msg == "approval for X on 2027-03-11 now also breaks rule 7"
    assert approvals[0].rules == [7]
    assert approvals[1].rules == [7]
