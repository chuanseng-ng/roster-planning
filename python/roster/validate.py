import datetime
from roster.model import SETTINGS_KEYS, norm

DIGITS = "0123456789"
COUNT_KEYS = ["Count Day", "Count Evening", "Count Night",
              "Count Standby D/E", "Count Standby N"]
CANNOT_WORDS = ["DAY", "EVENING", "NIGHT"]
SHIFT_WORDS = ["ALL", "DAY", "EVENING", "NIGHT"]

def is_valid_year(text):
    """Exactly 4 digits (after trimming spaces), value 1900 to 9999."""
    t = text.strip(' ')
    if len(t) != 4:
        return False
    for ch in t:
        if DIGITS.find(ch) < 0:
            return False
    return int(t) >= 1900

def is_plain_ascii(text):
    for ch in text:
        if ord(ch) < 32 or ord(ch) > 126:
            return False
    return True

def is_valid_date(text):
    """yyyy-mm-dd and a real calendar date."""
    t = text.strip(' ')
    if len(t) != 10:
        return False
    for i in range(10):
        if i == 4 or i == 7:
            if t[i] != '-':
                return False
        elif DIGITS.find(t[i]) < 0:
            return False
    y, m, d = int(t[0:4]), int(t[5:7]), int(t[8:10])
    if y < 1900:
        return False
    try:
        datetime.date(y, m, d)
    except ValueError:
        return False
    return True

def list_has(items, text):
    for item in items:
        if item == text:
            return True
    return False

def find_setting(raw, key):
    """Value text of the settings row for key, or None."""
    for row in raw.settings:
        if norm(row[0]) == norm(key):
            return row[1]
    return None

def validate(raw):
    msgs = []

    # Settings
    year_ok = False
    year = 0
    for key in SETTINGS_KEYS:
        value = find_setting(raw, key)
        if value is None:
            msgs.append("Settings: missing key " + key)
        elif key == "Year":
            if is_valid_year(value):
                year_ok = True
                year = int(value.strip(' '))
            else:
                msgs.append("Settings: Year must be a whole number")
        elif list_has(COUNT_KEYS, key):
            t = value.strip(' ')
            if not (len(t) == 1 and DIGITS.find(t) >= 0):
                msgs.append("Settings: " + key + " must be a whole number from 0 to 9")
        else:
            if norm(value) != "Y" and norm(value) != "N":
                msgs.append("Settings: " + key + " must be Y or N")

    # Staff
    staff_names = []
    seen = []
    r = 0
    for row in raw.staff:
        r += 1
        name = row[0].strip(' ')
        if name == "":
            msgs.append("Staff row " + str(r) + ": name is blank")
        else:
            if list_has(seen, norm(name)):
                msgs.append("Staff row " + str(r) + ": duplicate name " + name)
            else:
                seen.append(norm(name))
                staff_names.append(norm(name))
            if not is_plain_ascii(name):
                msgs.append("Staff row " + str(r) + ": name must use plain ASCII characters")
            if norm(name) == "GAP":
                msgs.append("Staff row " + str(r) + ": name cannot be GAP")
        if norm(row[2]) != "Y" and norm(row[2]) != "N":
            msgs.append("Staff row " + str(r) + ": Senior must be Y or N")
        for piece in row[4].split(","):
            word = piece.strip(' ')
            if word != "" and not list_has(CANNOT_WORDS, norm(word)):
                msgs.append("Staff row " + str(r) + ": unknown Cannot do value " + word)

    # Holidays
    holiday_names = []
    seen_dates = []
    r = 0
    for row in raw.holidays:
        r += 1
        name = row[0].strip(' ')
        if name == "":
            msgs.append("Holidays row " + str(r) + ": holiday name is blank")
        else:
            holiday_names.append(norm(name))
            if not is_plain_ascii(name):
                msgs.append("Holidays row " + str(r) + ": holiday name must use plain ASCII characters")
        date_text = row[1].strip(' ')
        if not is_valid_date(date_text):
            msgs.append("Holidays row " + str(r) + ": invalid date")
        else:
            if year_ok and int(date_text[0:4]) != year:
                msgs.append("Holidays row " + str(r) + ": date is outside year " + str(year))
            if list_has(seen_dates, date_text):
                msgs.append("Holidays row " + str(r) + ": duplicate date " + date_text)
            else:
                seen_dates.append(date_text)

    # LastYear
    r = 0
    for row in raw.lastyear:
        r += 1
        if not list_has(staff_names, norm(row[2])):
            msgs.append("LastYear row " + str(r) + ": unknown staff name " + row[2].strip(' '))
        if not list_has(holiday_names, norm(row[0])):
            msgs.append("LastYear row " + str(r) + ": holiday name " + row[0].strip(' ')
                        + " is not on the Holidays sheet")
        if row[3].strip(' ') == "":
            msgs.append("LastYear row " + str(r) + ": Slot is blank")

    # Leave
    r = 0
    for row in raw.leave:
        r += 1
        if not list_has(staff_names, norm(row[0])):
            msgs.append("Leave row " + str(r) + ": unknown staff name " + row[0].strip(' '))
        from_ok = is_valid_date(row[1])
        to_ok = is_valid_date(row[2])
        if not from_ok or not to_ok:
            msgs.append("Leave row " + str(r) + ": invalid date")
        else:
            if row[1].strip(' ') > row[2].strip(' '):
                msgs.append("Leave row " + str(r) + ": From date is after To date")
        if not list_has(SHIFT_WORDS, norm(row[3])):
            msgs.append("Leave row " + str(r) + ": unknown Shift " + row[3].strip(' '))

    return msgs
