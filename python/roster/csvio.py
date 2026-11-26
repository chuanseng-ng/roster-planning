from dataclasses import dataclass
import csv
import os

STAFF_HEADER = ["Name", "Title", "Senior", "Role", "Cannot do"]
HOLIDAYS_HEADER = ["Holiday name", "Date", "Added"]
LASTYEAR_HEADER = ["Holiday name", "Date", "Staff name", "Slot"]
LEAVE_HEADER = ["Staff name", "From date", "To date", "Shift"]
SETTINGS_HEADER = ["Key", "Value"]

@dataclass
class RawTables:
    staff: list      # rows of [Name, Title, Senior, Role, Cannot do]
    holidays: list   # rows of [Holiday name, Date, Added]
    lastyear: list   # rows of [Holiday name, Date, Staff name, Slot]
    leave: list      # rows of [Staff name, From date, To date, Shift]
    settings: list   # rows of [Key, Value]

def read_table(path, width, required):
    """Rows below the header, padded to width, blank rows dropped."""
    rows = []
    if not os.path.exists(path):
        if required:
            raise FileNotFoundError(path)
        return rows
    with open(path, newline='', encoding='utf-8-sig') as f:
        first = True
        for rec in csv.reader(f):
            if first:
                first = False
                continue
            row = []
            for i in range(width):
                if i < len(rec):
                    row.append(rec[i])
                else:
                    row.append("")
            blank = True
            for cell in row:
                if cell.strip(' ') != "":
                    blank = False
            if not blank:
                rows.append(row)
    return rows

def read_scenario(folder):
    return RawTables(
        read_table(os.path.join(folder, "staff.csv"), 5, True),
        read_table(os.path.join(folder, "holidays.csv"), 3, True),
        read_table(os.path.join(folder, "lastyear.csv"), 4, False),
        read_table(os.path.join(folder, "leave.csv"), 4, False),
        read_table(os.path.join(folder, "settings.csv"), 2, True))

def write_rows(path, header, rows):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow(header)
        for row in rows:
            w.writerow(row)
