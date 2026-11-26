import copy
import random

from roster import engine, model
from roster.engine import new_result, solve_year, max_matching
from roster.model import GAP, GAP_SENIOR, GAP_COAG, EMPTY_CELL
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


BASE = [("A1", "SC"), ("A2", "SC"), ("B1", ""), ("B2", ""), ("B3", ""), ("B4", ""), ("B5", ""), ("B6", "")]
# Default layout columns: 0 Day 1, 1 Day 2, 2 Evening, 3 Night 1, 4 Night 2, 5 Standby D/E, 6 Standby N.


def test_single_date_fills_in_candidate_order():
    inp = make_inputs(BASE, ["2027-01-01"])
    res = solved(inp)
    assert names(inp, res, 0) == ["A1", "B1", "B2", "A2", "B3", "B4", "B5"]
    assert res.nodes[0] == 7 and not res.cut_short[0] and res.unmet[0] == []
    assert res.active_count == [1, 1, 1, 1, 1, 0, 0, 0]
    assert res.standby_count == [0, 0, 0, 0, 0, 1, 1, 0]


def test_second_date_prefers_lowest_ratio_then_fewest_standby():
    inp = make_inputs(BASE, ["2027-01-01", "2027-03-10"])
    res = solved(inp)
    assert names(inp, res, 1) == ["A1", "B6", "B4", "A2", "B5", "B1", "B2"]


def test_only_senior_for_one_group_reserves_gap_and_uses_coag_standby():
    staff = [("S1", "SC"), ("S2", "SC")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-01-01"], exempt=[("S1", 0)])
    res = solved(inp)
    assert names(inp, res, 0) == ["S2", "B1", "B2", "B3", "GAP:S", "B4", "S1"]
    assert not res.cut_short[0]


def test_no_coag_at_all_reserves_both_standby_slots():
    staff = [("S1", "S"), ("S2", "S")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-01-01"])
    res = solved(inp)
    assert names(inp, res, 0) == ["S1", "B1", "B2", "S2", "B3", "GAP:C", "GAP:C"]


def test_matching_values():
    assert max_matching(3, [[0], [0, 1], [1]]) == 2
    assert max_matching(2, [[], []]) == 0
    assert max_matching(3, [[0, 1], [0], [2]]) == 3


def test_coag_fallback_in_both_groups():
    # Coag staff C1, C2 are exempt, so they sort last and cannot take active slots.
    # Order: B1..B6, S1, S2, C1, C2. S1 -> Day 1, C1 -> Standby D/E, S2 -> Night 1,
    # C2 -> Standby N, then B1, B2, B3 into Day 2, Evening, Night 2. First leaf equals bound.
    staff = [("S1", "S"), ("S2", "S"), ("C1", "C"), ("C2", "C")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-01-01"], exempt=[("C1", 0), ("C2", 0)])
    res = solved(inp)
    assert names(inp, res, 0) == ["S1", "B1", "B2", "S2", "B3", "C1", "C2"]
    assert res.nodes[0] == 7 and not res.cut_short[0] and res.unmet[0] == []


def test_coag_unmet_with_no_standby_slots():
    # Order B1, S1, S2. S1 -> Day 1 (senior DE), coag DE unmet -> GAP:C in Evening (last empty
    # active DE slot), S2 -> Night 1, coag N unmet -> GAP:C in Night 2, B1 -> Day 2.
    inp = make_inputs([("S1", "S"), ("S2", "S"), ("B1", "")], ["2027-01-01"], counts=(2, 1, 2, 0, 0))
    res = solved(inp)
    assert names(inp, res, 0) == ["S1", "B1", "GAP:C", "S2", "GAP:C"]
    assert res.unmet[0] == [] and res.nodes[0] == 5


def test_unmet_without_a_slot():
    # One Night slot only: senior step puts GAP:S in it; coag step then has no slot left.
    inp = make_inputs([("B1", "")], ["2027-01-01"], counts=(0, 0, 1, 0, 0))
    res = solved(inp)
    assert names(inp, res, 0) == ["GAP:S"]
    assert res.unmet[0] == [(1, GAP_COAG)]


def test_plain_shortage_stops_at_the_bound():
    # A1 (senior+coag, cannot do nights) -> Day 1; SC1 (only night-capable person) -> Night 1
    # as senior N; Night 2 and Standby N nobody. First leaf (2,1,0) equals the bound.
    staff = [("A1", "SCn")] + [("N%02d" % i, "n") for i in range(1, 19)] + [("SC1", "SC")]
    assert len(staff) == 20
    inp = make_inputs(staff, ["2027-01-01"])
    res = solved(inp)
    assert names(inp, res, 0) == ["A1", "N01", "N02", "SC1", "GAP", "N03", "GAP"]
    assert res.cut_short[0] is False and res.nodes[0] < 100


def test_backtracking_undoes_first_senior():
    # A1 is the only senior able to do Night; A2 (senior+coag, cannot Night). The first leaf puts
    # A1 in Day 1 and leaves GAP:S for Night; search must reach A2 -> Day 1, A1 -> Night 1.
    staff = [("A1", "SC"), ("A2", "SCn")] + [("B%d" % i, "") for i in range(1, 6)]
    inp = make_inputs(staff, ["2027-01-01"])
    res = solved(inp)
    assert names(inp, res, 0) == ["A2", "B1", "B2", "A1", "B3", "B4", "B5"]


def test_rest_after_night():
    staff = [("A1", "SC"), ("A2", "SC")] + [("B%d" % i, "") for i in range(1, 9)]
    inp = make_inputs(staff, ["2027-02-06", "2027-02-07"])
    res = solved(inp)
    assert names(inp, res, 0) == ["A1", "B1", "B2", "A2", "B3", "B4", "B5"]
    assert names(inp, res, 1) == ["A1", "B6", "B7", "A2", "B8", "B4", "B5"]
    for k in (3, 4):
        s = res.grid[0][k]
        for j, slot in enumerate(inp.slots):
            if res.grid[1][j] == s:
                assert slot.kind in (model.NIGHT, model.SBY_N)


def test_rule_switches():
    inp = make_inputs([("B%d" % i, "") for i in range(1, 8)], ["2027-01-01"], rules_off=(1, 2))
    res = solved(inp)
    assert names(inp, res, 0) == ["B1", "B2", "B3", "B4", "B5", "B6", "B7"]


def test_approved_cell_kept():
    inp = make_inputs(BASE, ["2027-01-01", "2027-03-10"])
    res = solved(inp)
    res.grid[1] = [EMPTY_CELL] * 7
    res.grid[1][0] = 7  # B6
    res.approved[1][0] = True
    solve_year(inp, res, 1)
    assert names(inp, res, 1) == ["B6", "A1", "B4", "A2", "B5", "B1", "B2"]
    assert res.grid[1].count(7) == 1
    assert res.active_count == [2, 2, 1, 1, 1, 1, 1, 1]
    assert res.standby_count == [0, 0, 1, 1, 0, 1, 1, 0]


def test_forward_night_block():
    # B1 cannot do Day/Evening but is approved in Day 1 on day 2. Without the forward block B1
    # would take Night 2 on day 1; with it, B1 goes to Standby N.
    staff = [("A1", "SC"), ("A2", "SC"), ("B1", "de")] + [("B%d" % i, "") for i in range(2, 7)]
    inp = make_inputs(staff, ["2027-02-06", "2027-02-07"])
    res = new_result(inp)
    res.grid[1][0] = 2
    res.approved[1][0] = True
    solve_year(inp, res, 0)
    assert names(inp, res, 0) == ["A1", "B2", "B3", "A2", "B4", "B5", "B1"]
    assert res.grid[1][0] == 2


def test_node_cap(monkeypatch):
    staff = [("S1", "SC"), ("S2", "SC")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-01-01"], exempt=[("S1", 0)])
    monkeypatch.setattr(model, "NODE_CAP", 50)
    res = solved(inp)
    assert res.nodes[0] <= 51
    assert EMPTY_CELL not in res.grid[0]
    assert names(inp, res, 0) == ["S2", "B1", "B2", "B3", "GAP:S", "B4", "S1"]


def test_node_cap_small(monkeypatch):
    # The first leaf needs 7 nodes; with cap 8 the 9th node trips the cut-off.
    staff = [("S1", "SC"), ("S2", "SC")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-01-01"], exempt=[("S1", 0)])
    monkeypatch.setattr(model, "NODE_CAP", 8)
    res = solved(inp)
    assert res.cut_short[0] is True and res.nodes[0] == 9
    assert names(inp, res, 0) == ["S2", "B1", "B2", "B3", "GAP:S", "B4", "S1"]


def test_determinism():
    inp = make_inputs(BASE, ["2027-01-01", "2027-03-10", "2027-03-11"])
    a = solved(inp)
    b = solved(inp)
    assert a.grid == b.grid and a.nodes == b.nodes and a.unmet == b.unmet


def test_recount_leaves_result_unchanged():
    inp = make_inputs(BASE, ["2027-01-01", "2027-03-10"])
    res = solved(inp)
    before = copy.deepcopy(res)
    solve_year(inp, res, 1)
    assert res == before
    assert names(inp, res, 1) == ["A1", "B6", "B4", "A2", "B5", "B1", "B2"]
    solve_year(inp, res, len(inp.dates))
    assert res == before


def solved_exhaustive(inp, monkeypatch):
    """Solve the year with the early stop disabled: the bound is unreachable, so the search
    runs to completion (pruning against the best score still applies)."""
    with monkeypatch.context() as m:
        m.setattr(engine._Search, "lower_bound", lambda self: (-1, -1, -1))
        return solved(inp)


def test_penalty_prefers_unmet_coag_slot_over_fallback_plus_gap(monkeypatch):
    # Counts (1, 2, 0, 1, 1): slots Day, Evening 1, Evening 2, Standby D/E, Standby N.
    # P3, P4 ("en") can only do Day. No seniors. Hand trace, penalty = 2 per unmet + 1 per fallback.
    # Date 0: order P0..P4 (all ratios 0). Senior step -> GAP:S in Evening 2 (last empty active),
    #   (1,1,2). Coag step (a): P0 in Day first (first leaf worse), improved by P0 in Evening 1,
    #   then P3 Day, P1 Standby D/E, P2 Standby N -> (1,1,2) == bound (1,1,2): [P3,P0,GAP:S,P1,P2].
    # Date 1: eligible P0=1,P1=2,P2=1,P3=2,P4=2; active P0=1,P3=1; standby P1=1,P2=1 ->
    #   order P4,P1,P2,P3,P0. P0 on leave, P2 exempt (standby only), P3/P4 Day only.
    #   Senior -> GAP:S Evening 2 (1,1,2). Coag (a) none; (b) P2 -> Standby D/E (0,0,1) -> (1,1,3);
    #   P4 Day, P1 Evening 1, Standby N: nobody -> GAP (2,1,3). Bound: gapsA 2, activeA 1,
    #   reqG/reqA 1, p_lb 2+1 -> (2,1,3) == score, stop. The alternative GAP:C in Standby D/E with
    #   P2 in Standby N is (2,1,4), worse; the old scoring (fallback count) wrongly preferred it.
    staff = [("P0", "C"), ("P1", ""), ("P2", "C"), ("P3", "en"), ("P4", "en")]
    inp = make_inputs(staff, ["2027-02-10", "2027-02-11"], counts=(1, 2, 0, 1, 1),
                      exempt=[("P2", 1)], leave=[("P0", 1, model.M_ALL)])
    res = solved(inp)
    assert names(inp, res, 0) == ["P3", "P0", "GAP:S", "P1", "P2"]
    assert names(inp, res, 1) == ["P4", "P1", "GAP:S", "P2", "GAP"]
    ex = solved_exhaustive(inp, monkeypatch)
    assert ex.grid == res.grid and ex.unmet == res.unmet
    assert not any(ex.cut_short)


def test_early_stop_matches_exhaustive_search_random(monkeypatch):
    rng = random.Random(20261126)
    checked = 0
    for _ in range(200):
        n = rng.randint(4, 7)
        staff = []
        for i in range(n):
            flags = "".join(f for f in "SCden" if rng.random() < 0.25)
            staff.append(("P%d" % i, flags))
        counts = tuple(rng.choice([0, 1, 1, 2]) for _ in range(5))
        nd = rng.randint(1, 2)
        dates = ["2027-02-10", "2027-02-11"][:nd]
        exempt = [(nm, d) for nm, _ in staff for d in range(nd) if rng.random() < 0.15]
        leave = [(nm, d, rng.choice([1, 2, 4, 8, 3])) for nm, _ in staff for d in range(nd)
                 if rng.random() < 0.15]
        inp = make_inputs(staff, dates, counts=counts, exempt=exempt, leave=leave)
        res = solved(inp)
        ex = solved_exhaustive(inp, monkeypatch)
        if any(ex.cut_short):
            continue
        checked += 1
        assert res.grid == ex.grid and res.unmet == ex.unmet
    assert checked > 100
