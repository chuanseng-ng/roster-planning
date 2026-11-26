from roster.model import EMPTY_CELL, M_ALL, M_DAY, M_EVE, M_NIGHT
from roster.rules import blocking_rules, forward_night_block
from tests.helpers import make_inputs

DATES = ["2027-02-05", "2027-02-06", "2027-02-07"]
# slot indexes with default counts: 0 Day 1, 1 Day 2, 2 Evening, 3 Night 1, 4 Night 2,
# 5 Standby D/E, 6 Standby N
ALL_SLOTS = [0, 1, 2, 3, 4, 5, 6]


def empty_grid(n):
    return [[EMPTY_CELL] * 7 for _ in range(n)]


def blocked_slots(inp, grid, s, d):
    return [k for k in ALL_SLOTS if blocking_rules(inp, grid, s, d, k)]


def test_rule3_other_cell_blocks_same_cell_does_not():
    inp = make_inputs([("Ann", ""), ("Bob", "")], DATES)
    grid = empty_grid(3)
    grid[1][0] = 0
    assert blocking_rules(inp, grid, 0, 1, 2) == [3]
    assert blocking_rules(inp, grid, 0, 1, 0) == []
    assert blocking_rules(inp, grid, 1, 1, 2) == []


def test_rule3_other_date_does_not_block():
    inp = make_inputs([("Ann", "")], DATES)
    grid = empty_grid(3)
    grid[0][0] = 0
    assert blocking_rules(inp, grid, 0, 1, 2) == []


def test_rule4_exempt_blocks_active_not_standby():
    inp = make_inputs([("Ann", "")], DATES, exempt=[("Ann", 1)])
    grid = empty_grid(3)
    assert blocked_slots(inp, grid, 0, 1) == [0, 1, 2, 3, 4]
    assert blocking_rules(inp, grid, 0, 1, 0) == [4]
    assert blocking_rules(inp, grid, 0, 1, 5) == []
    assert blocking_rules(inp, grid, 0, 1, 6) == []
    assert blocked_slots(inp, grid, 0, 0) == []


def test_rule5_leave_masks():
    grid = empty_grid(3)
    inp = make_inputs([("Ann", "")], DATES, leave=[("Ann", 1, M_DAY)])
    assert blocked_slots(inp, grid, 0, 1) == [0, 1, 5]
    assert blocking_rules(inp, grid, 0, 1, 0) == [5]
    inp = make_inputs([("Ann", "")], DATES, leave=[("Ann", 1, M_EVE)])
    assert blocked_slots(inp, grid, 0, 1) == [2, 5]
    inp = make_inputs([("Ann", "")], DATES, leave=[("Ann", 1, M_NIGHT)])
    assert blocked_slots(inp, grid, 0, 1) == [3, 4, 6]
    inp = make_inputs([("Ann", "")], DATES, leave=[("Ann", 1, M_ALL)])
    assert blocked_slots(inp, grid, 0, 1) == ALL_SLOTS
    assert blocked_slots(inp, grid, 0, 0) == []


def test_rule6_cannot_flags():
    grid = empty_grid(3)
    inp = make_inputs([("Ann", "d")], DATES)
    assert blocked_slots(inp, grid, 0, 1) == [0, 1, 5]
    assert blocking_rules(inp, grid, 0, 1, 0) == [6]
    inp = make_inputs([("Ann", "e")], DATES)
    assert blocked_slots(inp, grid, 0, 1) == [2, 5]
    inp = make_inputs([("Ann", "n")], DATES)
    assert blocked_slots(inp, grid, 0, 1) == [3, 4, 6]
    inp = make_inputs([("Ann", "den")], DATES)
    assert blocked_slots(inp, grid, 0, 1) == ALL_SLOTS


def test_rule7_night_blocks_next_day_non_night():
    inp = make_inputs([("Ann", "")], DATES)
    grid = empty_grid(3)
    grid[0][3] = 0
    assert blocking_rules(inp, grid, 0, 1, 0) == [7]
    assert blocking_rules(inp, grid, 0, 1, 1) == [7]
    assert blocking_rules(inp, grid, 0, 1, 2) == [7]
    assert blocking_rules(inp, grid, 0, 1, 5) == [7]
    assert blocking_rules(inp, grid, 0, 1, 3) == []
    assert blocking_rules(inp, grid, 0, 1, 4) == []
    assert blocking_rules(inp, grid, 0, 1, 6) == []


def test_rule7_second_night_slot_also_counts():
    inp = make_inputs([("Ann", "")], DATES)
    grid = empty_grid(3)
    grid[0][4] = 0
    assert blocking_rules(inp, grid, 0, 1, 0) == [7]


def test_rule7_standby_night_does_not_block():
    inp = make_inputs([("Ann", "")], DATES)
    grid = empty_grid(3)
    grid[0][6] = 0
    assert blocked_slots(inp, grid, 0, 1) == []


def test_rule7_first_date_blocks_nothing():
    inp = make_inputs([("Ann", "")], DATES)
    grid = empty_grid(3)
    assert blocked_slots(inp, grid, 0, 0) == []


def test_rule7_gap_between_dates_blocks_nothing():
    inp = make_inputs([("Ann", "")], ["2027-02-05", "2027-02-07"])
    grid = empty_grid(2)
    grid[0][3] = 0
    assert blocked_slots(inp, grid, 0, 1) == []


def test_rule7_off_blocks_nothing():
    inp = make_inputs([("Ann", "")], DATES, rules_off=(7,))
    grid = empty_grid(3)
    grid[0][3] = 0
    assert blocked_slots(inp, grid, 0, 1) == []


def test_several_rules_ascending():
    inp = make_inputs([("Ann", "d")], DATES, exempt=[("Ann", 1)])
    grid = empty_grid(3)
    assert blocking_rules(inp, grid, 0, 1, 0) == [4, 6]
    grid[1][2] = 0
    grid[0][3] = 0
    assert blocking_rules(inp, grid, 0, 1, 0) == [3, 4, 6, 7]


def test_filled_cell_itself_is_ignored():
    inp = make_inputs([("Ann", "")], DATES)
    grid = empty_grid(3)
    grid[1][0] = 0
    assert blocking_rules(inp, grid, 0, 1, 0) == []


def approved_grid(n):
    return [[False] * 7 for _ in range(n)]


def test_forward_night_block_true_cases():
    inp = make_inputs([("Ann", "")], DATES)
    grid = empty_grid(3)
    for k in (0, 1, 2, 5):
        grid = empty_grid(3)
        ap = approved_grid(3)
        grid[1][k] = 0
        ap[1][k] = True
        assert forward_night_block(inp, grid, ap, 0, 0, 3) is True
        assert forward_night_block(inp, grid, ap, 0, 0, 4) is True


def test_forward_night_block_not_approved():
    inp = make_inputs([("Ann", "")], DATES)
    grid = empty_grid(3)
    ap = approved_grid(3)
    grid[1][0] = 0
    assert forward_night_block(inp, grid, ap, 0, 0, 3) is False


def test_forward_night_block_next_day_night_slots():
    inp = make_inputs([("Ann", "")], DATES)
    for k in (3, 4, 6):
        grid = empty_grid(3)
        ap = approved_grid(3)
        grid[1][k] = 0
        ap[1][k] = True
        assert forward_night_block(inp, grid, ap, 0, 0, 3) is False


def test_forward_night_block_dates_not_consecutive():
    inp = make_inputs([("Ann", "")], ["2027-02-05", "2027-02-07"])
    grid = empty_grid(2)
    ap = approved_grid(2)
    grid[1][0] = 0
    ap[1][0] = True
    assert forward_night_block(inp, grid, ap, 0, 0, 3) is False


def test_forward_night_block_slot_not_night():
    inp = make_inputs([("Ann", "")], DATES)
    grid = empty_grid(3)
    ap = approved_grid(3)
    grid[1][0] = 0
    ap[1][0] = True
    for k in (0, 1, 2, 5, 6):
        assert forward_night_block(inp, grid, ap, 0, 0, k) is False


def test_forward_night_block_last_date_and_rule_off():
    inp = make_inputs([("Ann", "")], DATES)
    grid = empty_grid(3)
    ap = approved_grid(3)
    assert forward_night_block(inp, grid, ap, 0, 2, 3) is False
    inp = make_inputs([("Ann", "")], DATES, rules_off=(7,))
    grid[1][0] = 0
    ap[1][0] = True
    assert forward_night_block(inp, grid, ap, 0, 0, 3) is False


def test_forward_night_block_other_staff_in_next_cell():
    inp = make_inputs([("Ann", ""), ("Bob", "")], DATES)
    grid = empty_grid(3)
    ap = approved_grid(3)
    grid[1][0] = 1
    ap[1][0] = True
    assert forward_night_block(inp, grid, ap, 0, 0, 3) is False
