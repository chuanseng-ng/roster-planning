from roster.engine import new_result, solve_year
from roster.model import GAP, M_DAY, to_serial
from roster.review import (ReviewRow, Approval, build_review, check_rules, ratio_text,
                           build_summary)
from tests.helpers import make_inputs


def solved(inp):
    res = new_result(inp)
    solve_year(inp, res, 0)
    return res


BASE = [("A1", "SC"), ("A2", "SC"), ("B1", ""), ("B2", ""), ("B3", ""), ("B4", ""), ("B5", ""), ("B6", "")]
# Default layout columns: 0 Day 1, 1 Day 2, 2 Evening, 3 Night 1, 4 Night 2, 5 Standby D/E, 6 Standby N.
D1 = to_serial("2027-01-01")


def no_senior_inputs(s1_flags="S"):
    staff = [("S1", s1_flags), ("S2", "S"), ("C1", "C"), ("C2", "C")] + [("B%d" % i, "") for i in range(1, 6)]
    return make_inputs(staff, ["2027-01-01"], exempt=[("S1", 0)])


# --- build_review ---------------------------------------------------------

def test_review_no_senior_gap_row():
    inp = no_senior_inputs()
    res = solved(inp)
    assert build_review(inp, res) == [
        ReviewRow(D1, "Night", "Night 2", "No senior", "S1: exempt; S2: already rostered", "S1 (rule 4 exemption)")]


def test_review_two_rules_is_in_reason_but_not_candidates():
    inp = no_senior_inputs("Sn")
    res = solved(inp)
    assert build_review(inp, res) == [
        ReviewRow(D1, "Night", "Night 2", "No senior", "S1: exempt; S2: already rostered", "")]


def test_review_only_senior_in_standby_is_already_rostered():
    staff = [("S1", "SC"), ("S2", "SC")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-01-01"], exempt=[("S1", 0)])
    res = solved(inp)
    assert build_review(inp, res) == [
        ReviewRow(D1, "Night", "Night 2", "No senior", "S1: already rostered; S2: already rostered", "")]


def test_review_plain_shortage():
    staff = [("A1", "SCn")] + [("N%02d" % i, "n") for i in range(1, 19)] + [("SC1", "SC")]
    inp = make_inputs(staff, ["2027-01-01"])
    res = solved(inp)
    rows = build_review(inp, res)
    assert len(rows) == 2
    assert rows[0].group == "Night" and rows[0].slot == "Night 2" and rows[0].rule == "No eligible staff"
    assert "N04: cannot do Night" in rows[0].reason
    assert "A1: already rostered; N01: already rostered" in rows[0].reason
    assert "N04 (rule 6 cannot do)" in rows[0].candidates
    assert "N03" not in rows[0].candidates and "A1" not in rows[0].candidates
    assert rows[1].slot == "Standby N"
    assert "cannot do Standby N" in rows[1].reason


def test_review_not_placed_and_nobody_considered():
    inp = make_inputs([("B1", "")], ["2027-01-01"], counts=(0, 0, 1, 0, 0))
    res = solved(inp)
    assert build_review(inp, res) == [
        ReviewRow(D1, "Night", "Night", "No senior", "no staff with this requirement", ""),
        ReviewRow(D1, "Night", "", "No special coag", "no free slot in this group", "")]
    # A considered person with no blocking rule reads "not placed".
    inp2 =make_inputs([("B1", "")], ["2027-01-01"], counts=(0, 0, 1, 0, 0))
    res2 = new_result(inp2)
    res2.grid[0][0] = GAP
    rows = build_review(inp2, res2)
    assert rows == [ReviewRow(D1, "Night", "Night", "No eligible staff", "B1: not placed", "")]


def test_review_cut_short_row_first_and_date_order():
    inp = make_inputs(BASE, ["2027-01-01", "2027-03-10"])
    res = solved(inp)
    assert build_review(inp, res) == []
    res.cut_short[1] = True
    assert build_review(inp, res) == [
        ReviewRow(to_serial("2027-03-10"), "", "", "Search cut short", "best roster found so far was kept", "")]
    inp2 = make_inputs([("B1", "")], ["2027-01-01", "2027-01-05"], counts=(0, 0, 1, 0, 0))
    res2 = solved(inp2)
    res2.cut_short[1] = True
    rows = build_review(inp2, res2)
    assert [(r.date, r.rule) for r in rows] == [
        (to_serial("2027-01-01"), "No senior"), (to_serial("2027-01-01"), "No special coag"),
        (to_serial("2027-01-05"), "Search cut short"), (to_serial("2027-01-05"), "No senior"),
        (to_serial("2027-01-05"), "No special coag")]


# --- check_rules ----------------------------------------------------------

def test_check_rules_clean_on_engine_output():
    inp = make_inputs(BASE, ["2027-01-01", "2027-03-10", "2027-03-11"])
    assert check_rules(inp, solved(inp), []) == []
    staff = [("S1", "SC"), ("S2", "SC")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-01-01"], exempt=[("S1", 0)])
    assert check_rules(inp, solved(inp), []) == []
    staff = [("S1", "S"), ("S2", "S"), ("C1", "C"), ("C2", "C")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-01-01"], exempt=[("C1", 0), ("C2", 0)])
    assert check_rules(inp, solved(inp), []) == []
    inp = make_inputs([("A1", "SC"), ("B1", ""), ("B2", "")], ["2027-01-01"], counts=(1, 0, 0, 0, 1))
    assert check_rules(inp, solved(inp), []) == []


def test_check_rule_3():
    inp = make_inputs(BASE, ["2027-01-01"])
    res = solved(inp)
    res.grid[0][1] = 0
    assert check_rules(inp, res, []) == ["2027-01-01 Day 1: A1 breaks rule 3",
                                         "2027-01-01 Day 2: A1 breaks rule 3"]


def test_check_rules_4_5_6():
    inp = make_inputs(BASE, ["2027-01-01"])
    res = solved(inp)
    inp.exempt[2][0] = True
    assert check_rules(inp, res, []) == ["2027-01-01 Day 2: B1 breaks rule 4"]
    inp.exempt[2][0] = False
    inp.leave[2][0] = M_DAY
    assert check_rules(inp, res, []) == ["2027-01-01 Day 2: B1 breaks rule 5"]
    inp.leave[2][0] = 0
    inp.cannot[2] = 1
    assert check_rules(inp, res, []) == ["2027-01-01 Day 2: B1 breaks rule 6"]


def test_check_rule_7():
    inp = make_inputs(BASE + [("B7", ""), ("B8", "")], ["2027-02-06", "2027-02-07"])
    res = solved(inp)
    res.grid[1][1] = 4      # B3 worked Night 2 the day before
    assert check_rules(inp, res, []) == ["2027-02-07 Day 2: B3 breaks rule 7"]


def test_check_rules_1_and_2_and_approvals():
    inp = make_inputs(BASE, ["2027-01-01"])
    res = solved(inp)
    res.grid[0][0] = 7      # B6 replaces A1 in Day 1: group 0 loses its senior and coag
    both = ["2027-01-01 Day + evening: rule 1 not met", "2027-01-01 Day + evening: rule 2 not met"]
    assert check_rules(inp, res, []) == both
    assert check_rules(inp, res, [Approval(0, 0, 7, [1])]) == both[1:]
    assert check_rules(inp, res, [Approval(0, 1, 2, [2])]) == both[:1]
    assert check_rules(inp, res, [Approval(0, 4, 4, [1, 2])]) == both     # other group: no excuse
    assert check_rules(inp, res, [Approval(0, 0, 7, [1, 2])]) == []


def test_check_rule_2_alone_and_rules_off():
    inp = make_inputs(BASE, ["2027-01-01"])
    res = solved(inp)
    inp.coag[0] = False
    assert check_rules(inp, res, []) == ["2027-01-01 Day + evening: rule 2 not met"]
    inp.settings.rules[2] = False
    assert check_rules(inp, res, []) == []


def test_check_rules_approval_excuses_only_its_rule_and_cell():
    inp = make_inputs(BASE, ["2027-01-01"])
    res = solved(inp)
    inp.exempt[2][0] = True
    inp.leave[2][0] = M_DAY
    msgs = ["2027-01-01 Day 2: B1 breaks rule 4", "2027-01-01 Day 2: B1 breaks rule 5"]
    assert check_rules(inp, res, []) == msgs
    assert check_rules(inp, res, [Approval(0, 1, 2, [4])]) == msgs[1:]
    assert check_rules(inp, res, [Approval(0, 1, 2, [5])]) == msgs[:1]
    assert check_rules(inp, res, [Approval(0, 2, 3, [4, 5])]) == msgs      # different cell
    assert check_rules(inp, res, [Approval(0, 1, 2, [4, 5])]) == []


def test_check_rules_skips_groups_with_gaps_or_unmet():
    inp = make_inputs([("B1", "")], ["2027-01-01"], counts=(0, 0, 1, 0, 0))
    res = solved(inp)
    assert check_rules(inp, res, []) == []
    res.grid[0][0] = 0
    assert check_rules(inp, res, []) == []              # unmet entry for the group
    res.unmet[0] = []
    assert check_rules(inp, res, []) == ["2027-01-01 Night: rule 1 not met",
                                         "2027-01-01 Night: rule 2 not met"]


def test_check_rules_message_order():
    inp = make_inputs(BASE, ["2027-01-01", "2027-03-10"])
    res = solved(inp)
    res.grid[0][1] = 0
    res.grid[1][0] = 7
    res.grid[1][1] = 7
    out = check_rules(inp, res, [])
    assert out == ["2027-01-01 Day 1: A1 breaks rule 3", "2027-01-01 Day 2: A1 breaks rule 3",
                   "2027-03-10 Day 1: B6 breaks rule 3", "2027-03-10 Day 2: B6 breaks rule 3",
                   "2027-03-10 Day + evening: rule 1 not met", "2027-03-10 Day + evening: rule 2 not met"]


# --- ratio_text and summary -----------------------------------------------

def test_ratio_text():
    assert ratio_text(1, 3) == "0.3333"
    assert ratio_text(2, 3) == "0.6667"
    assert ratio_text(1, 8) == "0.1250"
    assert ratio_text(1, 16) == "0.0625"
    assert ratio_text(3, 2) == "1.5000"
    assert ratio_text(0, 0) == ""


def test_summary_two_dates():
    inp = make_inputs(BASE, ["2027-01-01", "2027-03-10"])
    res = solved(inp)
    rows, spread = build_summary(inp, res)
    assert rows == [["A1", "2", "2", "1.0000", "0"], ["A2", "2", "2", "1.0000", "0"],
                    ["B1", "1", "2", "0.5000", "1"], ["B2", "1", "2", "0.5000", "1"],
                    ["B3", "1", "2", "0.5000", "0"], ["B4", "1", "2", "0.5000", "1"],
                    ["B5", "1", "2", "0.5000", "1"], ["B6", "1", "2", "0.5000", "0"]]
    assert spread == "0.5000"


def test_summary_no_eligible_staff():
    inp = make_inputs([("B1", "")], ["2027-01-01"], exempt=[("B1", 0)], counts=(0, 0, 1, 0, 0))
    res = solved(inp)
    rows, spread = build_summary(inp, res)
    assert rows == [["B1", "0", "0", "", "0"]]
    assert spread == "0.0000"
