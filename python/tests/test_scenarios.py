import csv
import os

from roster import gen_fixtures
from roster.model import GAP, GAP_SENIOR, GAP_COAG, EMPTY_CELL, NIGHT, SBY_N, from_serial, to_serial
from roster.review import build_review, check_rules, build_summary

FIXTURES = gen_fixtures.FIXTURES
EXPECTED_FILES = ["expected_roster.csv", "expected_review.csv", "expected_summary.csv"]


def solved(name):
    raw, inp, res, approval_rows, approvals = gen_fixtures.solve_scenario(
        os.path.join(FIXTURES, name))
    return inp, res, approvals


def slot_index(inp, name):
    for k in range(len(inp.slots)):
        if inp.slots[k].name == name:
            return k
    raise AssertionError("no slot " + name)


def staff_index(inp, name):
    for s in range(len(inp.staff)):
        if inp.staff[s] == name:
            return s
    raise AssertionError("no staff " + name)


def has_gap(res):
    for row in res.grid:
        for v in row:
            if v in (GAP, GAP_SENIOR, GAP_COAG, EMPTY_CELL):
                return True
    return False


def spread_hundredths(inp, res):
    _, spread = build_summary(inp, res)
    return int(spread.replace(".", ""))


def read_csv(path):
    with open(path, newline='', encoding='utf-8') as f:
        return list(csv.reader(f))


def test_normal_year():
    inp, res, approvals = solved("normal_year")
    assert len(inp.staff) == 30
    assert len(inp.dates) == 11
    assert not has_gap(res)
    assert build_review(inp, res) == []
    assert check_rules(inp, res, approvals) == []
    assert spread_hundredths(inp, res) <= 5000


def test_senior_exempt():
    inp, res, approvals = solved("senior_exempt")
    rows = build_review(inp, res)
    assert len(rows) == 1
    assert rows[0].rule == "No senior"
    assert "Tan Wei" in rows[0].candidates
    assert inp.senior[staff_index(inp, "Tan Wei")]
    assert inp.exempt[staff_index(inp, "Tan Wei")][0]


def test_coag_standby():
    inp, res, approvals = solved("coag_standby")
    assert not has_gap(res)
    assert len(build_review(inp, res)) == 0
    for s in range(len(inp.staff)):
        if inp.coag[s]:
            assert inp.exempt[s][0]
    for name in ("Standby D/E", "Standby N"):
        v = res.grid[0][slot_index(inp, name)]
        assert v >= 0 and inp.coag[v]


def test_leave_cannot():
    inp, res, approvals = solved("leave_cannot")
    assert check_rules(inp, res, approvals) == []
    raw = gen_fixtures.read_scenario(os.path.join(FIXTURES, "leave_cannot"))
    shifts = sorted(r[3] for r in raw.leave)
    assert shifts == ["All", "Day", "Evening", "Night"]
    flagged = 0
    for s in range(len(inp.staff)):
        if inp.cannot[s] != 0:
            flagged += 1
    assert flagged >= 3
    for d in range(len(inp.dates)):
        for k in range(len(inp.slots)):
            s = res.grid[d][k]
            if s < 0:
                continue
            lv = inp.leave[s][d]
            assert lv & 8 == 0 and lv & inp.slots[k].mask == 0
            assert inp.cannot[s] & inp.slots[k].mask == 0
    # each leave row actually bites: the named person is absent from the blocked cells
    for row in raw.leave:
        s = staff_index(inp, row[0])
        for d in range(len(inp.dates)):
            if to_serial(row[1]) <= inp.dates[d] <= to_serial(row[2]):
                for k in range(len(inp.slots)):
                    blocked = row[3] == "All" or (inp.slots[k].mask & {"Day": 1, "Evening": 2,
                                                                        "Night": 4}.get(row[3], 0)) != 0
                    if blocked:
                        assert res.grid[d][k] != s


def test_long_weekend():
    inp, res, approvals = solved("long_weekend")
    raw = gen_fixtures.read_scenario(os.path.join(FIXTURES, "long_weekend"))
    assert [r[2] for r in raw.holidays] == ["", "Y", "Y"]
    assert len(inp.dates) == 3
    last_active = []
    for row in raw.lastyear:
        if not row[3].startswith("Standby"):
            last_active.append(staff_index(inp, row[2]))
    assert len(last_active) == 5
    for d in range(3):
        for k in range(len(inp.slots)):
            if inp.slots[k].active:
                assert res.grid[d][k] not in last_active
    assert check_rules(inp, res, approvals) == []


def test_cny():
    inp, res, approvals = solved("cny")
    assert not has_gap(res)
    assert len(inp.dates) == 2 and inp.dates[1] == inp.dates[0] + 1
    assert inp.settings.rules[7]
    assert check_rules(inp, res, approvals) == []
    for k in range(len(inp.slots)):
        if inp.slots[k].kind == NIGHT:
            s = res.grid[0][k]
            for j in range(len(inp.slots)):
                if inp.slots[j].kind not in (NIGHT, SBY_N):
                    assert res.grid[1][j] != s


def test_rest_after_night():
    inp, res, approvals = solved("rest_after_night")
    assert len(inp.dates) == 3
    assert inp.dates[1] == inp.dates[0] + 1 and inp.dates[2] == inp.dates[1] + 1
    assert check_rules(inp, res, approvals) == []
    for d in range(2):
        for k in range(len(inp.slots)):
            if inp.slots[k].kind == NIGHT:
                s = res.grid[d][k]
                for j in range(len(inp.slots)):
                    if inp.slots[j].kind not in (NIGHT, SBY_N):
                        assert res.grid[d + 1][j] != s


def test_custom_counts():
    path = os.path.join(FIXTURES, "custom_counts", "expected_roster.csv")
    header = read_csv(path)[0]
    assert header[2:] == ["Day 1", "Day 2", "Day 3", "Evening 1", "Evening 2", "Night",
                          "Standby N 1", "Standby N 2"]
    inp, res, approvals = solved("custom_counts")
    assert inp.settings.counts == [3, 2, 1, 0, 2]
    assert not inp.settings.rules[1]


def test_approval():
    inp, res, approvals = solved("approval")
    assert len(inp.dates) == 3
    assert len(approvals) == 1
    a = approvals[0]
    assert res.grid[a.d][a.k] == a.s
    assert inp.staff[a.s] == "Tan Wei"
    assert inp.slots[a.k].name == "Night 2"
    assert res.approved[a.d][a.k]
    assert check_rules(inp, res, approvals) == []
    # without the approval the placement would break a rule
    assert check_rules(inp, res, []) != []


def test_no_scenario_is_cut_short():
    for name in gen_fixtures.scenario_names():
        inp, res, approvals = solved(name)
        assert not any(res.cut_short), name


def test_expected_files_are_current(tmp_path):
    gen_fixtures.generate(str(tmp_path))
    names = gen_fixtures.scenario_names()
    assert len(names) == 9
    for name in names:
        for fn in EXPECTED_FILES:
            with open(os.path.join(FIXTURES, name, fn), "rb") as f:
                committed = f.read()
            with open(os.path.join(str(tmp_path), name, fn), "rb") as f:
                fresh = f.read()
            assert committed == fresh, name + "/" + fn
    with open(os.path.join(FIXTURES, "testsheet.csv"), "rb") as f:
        committed = f.read()
    with open(os.path.join(str(tmp_path), "testsheet.csv"), "rb") as f:
        fresh = f.read()
    assert committed == fresh


def test_testsheet_lists_every_scenario():
    lines = read_csv(os.path.join(FIXTURES, "testsheet.csv"))
    assert len(lines) > 0
    width = len(lines[0])
    for line in lines:
        assert len(line) == width
    base = ["staff", "holidays", "lastyear", "leave", "settings",
            "expected_roster", "expected_review", "expected_summary"]
    for name in gen_fixtures.scenario_names():
        tables = []
        for line in lines:
            if line[0] == name and line[1] not in tables:
                tables.append(line[1])
        expected = list(base)
        if os.path.exists(os.path.join(FIXTURES, name, "approve.csv")):
            expected.insert(5, "approve")
        assert tables == expected, name
        for t in tables:
            assert [l[2] for l in lines if l[0] == name and l[1] == t][0] == "0"
