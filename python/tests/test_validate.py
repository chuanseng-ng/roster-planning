from roster.csvio import RawTables
from roster.validate import validate

def good():
    settings = [["Year", "2027"], ["Count Day", "2"], ["Count Evening", "1"], ["Count Night", "2"],
                ["Count Standby D/E", "1"], ["Count Standby N", "1"], ["Rule 1", "Y"],
                ["Rule 2", "Y"], ["Rule 4", "Y"], ["Rule 5", "Y"], ["Rule 6", "Y"], ["Rule 7", "N"]]
    staff = [["Ann", "Dr", "Y", "Haem, special coag", "Night"],
             ["Bob", "Dr", "N", "", ""],
             ["Cy", "Dr", "N", "", "Day, Evening"]]
    holidays = [["New Year's Day", "2027-01-01", "N"], ["Good Friday", "2027-03-26", "N"]]
    lastyear = [["Good Friday", "2026-04-03", "Ann", "Day 1"]]
    leave = [["Bob", "2027-03-26", "2027-03-27", "Day"]]
    return RawTables(staff, holidays, lastyear, leave, settings)

def check(raw, msg):
    assert msg in validate(raw)

def test_good_tables_have_no_messages():
    assert validate(good()) == []

def test_missing_key():
    raw = good()
    raw.settings = raw.settings[1:]
    check(raw, "Settings: missing key Year")

def test_year_not_whole():
    raw = good()
    raw.settings[0][1] = "20x7"
    check(raw, "Settings: Year must be a whole number")

def test_count_range():
    raw = good()
    raw.settings[1][1] = "10"
    check(raw, "Settings: Count Day must be a whole number from 0 to 9")

def test_rule_yn():
    raw = good()
    raw.settings[6][1] = "maybe"
    check(raw, "Settings: Rule 1 must be Y or N")

def test_staff_name_blank():
    raw = good()
    raw.staff[1][0] = "  "
    check(raw, "Staff row 2: name is blank")

def test_staff_duplicate():
    raw = good()
    raw.staff[2][0] = " ann "
    check(raw, "Staff row 3: duplicate name ann")

def test_staff_ascii():
    raw = good()
    raw.staff[1][0] = "Zoë"
    check(raw, "Staff row 2: name must use plain ASCII characters")

def test_staff_senior():
    raw = good()
    raw.staff[1][2] = "Q"
    check(raw, "Staff row 2: Senior must be Y or N")

def test_staff_cannot_do():
    raw = good()
    raw.staff[1][4] = "Day, Lunch,"
    check(raw, "Staff row 2: unknown Cannot do value Lunch")

def test_holiday_name_blank():
    raw = good()
    raw.holidays[0][0] = ""
    check(raw, "Holidays row 1: holiday name is blank")

def test_holiday_ascii():
    raw = good()
    raw.holidays[0][0] = "Café Day"
    check(raw, "Holidays row 1: holiday name must use plain ASCII characters")

def test_holiday_invalid_date():
    raw = good()
    raw.holidays[0][1] = "2027-02-30"
    check(raw, "Holidays row 1: invalid date")

def test_holiday_outside_year():
    raw = good()
    raw.holidays[0][1] = "2026-12-31"
    check(raw, "Holidays row 1: date is outside year 2027")

def test_holiday_duplicate_date():
    raw = good()
    raw.holidays[1][1] = "2027-01-01"
    check(raw, "Holidays row 2: duplicate date 2027-01-01")

def test_lastyear_unknown_staff():
    raw = good()
    raw.lastyear[0][2] = "Zed"
    check(raw, "LastYear row 1: unknown staff name Zed")

def test_lastyear_unknown_holiday():
    raw = good()
    raw.lastyear[0][0] = "Boxing Day"
    check(raw, "LastYear row 1: holiday name Boxing Day is not on the Holidays sheet")

def test_lastyear_slot_blank():
    raw = good()
    raw.lastyear[0][3] = " "
    check(raw, "LastYear row 1: Slot is blank")

def test_leave_unknown_staff():
    raw = good()
    raw.leave[0][0] = "Zed"
    check(raw, "Leave row 1: unknown staff name Zed")

def test_leave_invalid_date():
    raw = good()
    raw.leave[0][1] = "27-03-26"
    check(raw, "Leave row 1: invalid date")

def test_leave_from_after_to():
    raw = good()
    raw.leave[0][1] = "2027-03-28"
    check(raw, "Leave row 1: From date is after To date")

def test_leave_unknown_shift():
    raw = good()
    raw.leave[0][3] = "Lunch"
    check(raw, "Leave row 1: unknown Shift Lunch")

def test_all_problems_reported_together():
    raw = good()
    raw.staff[1][2] = "Q"
    raw.leave[0][3] = "Lunch"
    assert validate(raw) == ["Staff row 2: Senior must be Y or N",
                             "Leave row 1: unknown Shift Lunch"]
