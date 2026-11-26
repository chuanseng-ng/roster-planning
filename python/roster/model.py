from dataclasses import dataclass
import datetime

DAY, EVE, NIGHT, SBY_DE, SBY_N = 0, 1, 2, 3, 4
KIND_NAMES = ["Day", "Evening", "Night", "Standby D/E", "Standby N"]
KIND_GROUP = [0, 0, 1, 0, 1]
KIND_ACTIVE = [True, True, True, False, False]
KIND_MASK = [1, 2, 4, 3, 4]          # shifts whose leave / cannot-do block this kind
M_DAY, M_EVE, M_NIGHT, M_ALL = 1, 2, 4, 8
GROUP_NAMES = ["Day + evening", "Night"]
GAP, GAP_SENIOR, GAP_COAG, EMPTY_CELL = -1, -2, -3, -9
NODE_CAP = 200000
RULE_TEXT = {GAP: "No eligible staff", GAP_SENIOR: "No senior", GAP_COAG: "No special coag"}
CUT_SHORT_TEXT = "Search cut short"
GAP_TEXT = "GAP"
SWITCHABLE_RULES = [1, 2, 4, 5, 6, 7]
SETTINGS_KEYS = ["Year", "Count Day", "Count Evening", "Count Night",
                 "Count Standby D/E", "Count Standby N",
                 "Rule 1", "Rule 2", "Rule 4", "Rule 5", "Rule 6", "Rule 7"]

@dataclass
class Slot:
    kind: int
    group: int
    active: bool
    name: str
    mask: int

def build_slots(counts):
    """counts = [day, evening, night, standby_de, standby_n]. Columns in kind order.
    A slot is named '<kind name> <i>' when its count is above 1, else '<kind name>'."""
    slots = []
    for kind in range(5):
        n = counts[kind]
        for i in range(1, n + 1):
            name = KIND_NAMES[kind] + (" " + str(i) if n > 1 else "")
            slots.append(Slot(kind, KIND_GROUP[kind], KIND_ACTIVE[kind], name, KIND_MASK[kind]))
    return slots

def to_serial(text):
    """'yyyy-mm-dd' -> Excel serial (whole number)."""
    y, m, d = int(text[0:4]), int(text[5:7]), int(text[8:10])
    return datetime.date(y, m, d).toordinal() - 693594

def from_serial(n):
    return datetime.date.fromordinal(n + 693594).isoformat()

def norm(text):
    return text.strip(' ').upper()

@dataclass
class Settings:
    year: int
    counts: list            # 5 ints
    rules: dict             # {1: bool, 2: bool, 4: bool, 5: bool, 6: bool, 7: bool}

@dataclass
class Inputs:
    settings: Settings
    slots: list             # list[Slot]
    staff: list             # names as entered, sheet order
    senior: list            # bool per staff
    coag: list              # bool per staff
    cannot: list            # mask per staff (already 0 when rule 6 is off)
    dates: list             # serials, ascending, unique
    holiday: list           # holiday name per date, as entered
    exempt: list            # exempt[s][d] bool (already False when rule 4 is off)
    leave: list             # leave[s][d] mask (already 0 when rule 5 is off)

@dataclass
class YearResult:
    grid: list              # grid[d][k]: staff index, or GAP / GAP_SENIOR / GAP_COAG / EMPTY_CELL
    approved: list          # approved[d][k] bool
    active_count: list      # per staff
    standby_count: list     # per staff
    eligible: list          # eligible dates per staff (fixed)
    cut_short: list         # bool per date
    unmet: list             # unmet[d]: list of (group, code) with no slot
    nodes: list             # node count per date (for tests and parity)
