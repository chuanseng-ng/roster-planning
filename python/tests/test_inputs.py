from roster.csvio import RawTables, read_scenario, write_rows
from roster.inputs import build_inputs
from roster.validate import validate
from roster.model import to_serial

def tables(lastyear=(), leave=(), rules=None, holidays=None, staff=None):
    r = {"Rule 1": "Y", "Rule 2": "Y", "Rule 4": "Y", "Rule 5": "Y", "Rule 6": "Y", "Rule 7": "N"}
    if rules:
        r.update(rules)
    settings = [["Year", "2027"], ["Count Day", "2"], ["Count Evening", "1"], ["Count Night", "2"],
                ["Count Standby D/E", "1"], ["Count Standby N", "1"]]
    for k in ["Rule 1", "Rule 2", "Rule 4", "Rule 5", "Rule 6", "Rule 7"]:
        settings.append([k, r[k]])
    if staff is None:
        staff = [["Ann", "Dr", "Y", "Haem", "Night"], ["Bob", "Dr", "N", "", ""]]
    if holidays is None:
        holidays = [["New Year's Day", "2027-01-01", "N"], ["Good Friday", "2027-03-26", "N"],
                    ["Good Friday", "2027-03-27", "N"]]
    return RawTables(staff, holidays, [list(x) for x in lastyear], [list(x) for x in leave], settings)

LY = [[" good friday ", "2026-04-03", "Ann", "Day 1"], ["Good Friday", "2026-04-03", "Bob", "Standby N"]]
LV = [["Ann", "2027-03-26", "2027-03-27", "Day"], ["Ann", "2027-03-27", "2027-03-27", "Night"]]

def test_exemption_by_holiday_name_ignores_standby_and_case():
    raw = tables(lastyear=LY)
    assert validate(raw) == []
    inp = build_inputs(raw)
    assert inp.exempt == [[False, True, True], [False, False, False]]

def test_rule4_off_clears_exemptions():
    inp = build_inputs(tables(lastyear=LY, rules={"Rule 4": "N"}))
    assert inp.exempt == [[False, False, False], [False, False, False]]

def test_leave_masks_cover_range_and_combine():
    raw = tables(leave=LV)
    assert validate(raw) == []
    assert build_inputs(raw).leave == [[0, 1, 5], [0, 0, 0]]

def test_rule5_and_rule6_off_clear_masks():
    inp = build_inputs(tables(leave=LV, rules={"Rule 5": "N", "Rule 6": "N"}))
    assert inp.leave == [[0, 0, 0], [0, 0, 0]]
    assert inp.cannot == [0, 0]
    inp = build_inputs(tables(leave=LV))
    assert inp.cannot == [4, 0]

def test_role_list_detects_special_coag():
    staff = [["A", "", "N", "Haem, special coag ", ""], ["B", "", "N", "Haem, special  coag", ""],
             ["C", "", "N", "Special Coagulation", ""]]
    inp = build_inputs(tables(staff=staff))
    assert inp.coag == [True, False, False]

def test_dates_sorted_and_names_follow():
    hol = [["Zeta ", "2027-12-25", "N"], ["Alpha", "2027-01-01", "N"], ["Mid", "2027-06-01", "N"]]
    inp = build_inputs(tables(holidays=hol))
    assert inp.dates == [to_serial("2027-01-01"), to_serial("2027-06-01"), to_serial("2027-12-25")]
    assert inp.holiday == ["Alpha", "Mid", "Zeta"]
    assert inp.staff == ["Ann", "Bob"]
    assert inp.senior == [True, False]
    assert inp.settings.counts == [2, 1, 2, 1, 1]
    assert inp.settings.rules[7] is False

def test_read_scenario_roundtrip(tmp_path):
    raw = tables(lastyear=LY, leave=LV)
    write_rows(str(tmp_path / "staff.csv"), ["Name", "Title", "Senior", "Role", "Cannot do"], raw.staff)
    write_rows(str(tmp_path / "holidays.csv"), ["Holiday name", "Date", "Added"], raw.holidays)
    write_rows(str(tmp_path / "lastyear.csv"), ["Holiday name", "Date", "Staff name", "Slot"], raw.lastyear)
    write_rows(str(tmp_path / "leave.csv"), ["Staff name", "From date", "To date", "Shift"], raw.leave)
    write_rows(str(tmp_path / "settings.csv"), ["Key", "Value"], raw.settings)
    back = read_scenario(str(tmp_path))
    assert back == raw
    (tmp_path / "leave.csv").unlink()
    assert read_scenario(str(tmp_path)).leave == []
