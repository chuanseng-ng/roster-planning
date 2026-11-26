# PH Roster Planner Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an Excel VBA workbook that generates the public-holiday active and standby roster, with a Python twin of the engine that is tested here and supplies the expected results for the workbook's self-test.

**Architecture:** A pure roster engine (arrays in, arrays out) is written first in Python with tests, then ported function by function to VBA. Scenario fixtures (CSV) are shared: Python generates expected results, a PowerShell script builds the `.xlsm` on Windows from the VBA text modules and fixtures, and the VBA self-test compares against the Python expectations.

**Tech Stack:** Python 3.13, pytest (dev only); Excel VBA (Windows desktop Excel); PowerShell 5+ with Excel COM (build only).

**Specs:** `docs/source-spec.md` and `docs/superpowers/specs/2026-11-26-ph-roster-planner-design.md` ("the design"). The design wins where they differ; design section 11 wins over its sections 5 and 7.

**Roles:** Sonnet agents implement one task each. An Opus agent reviews each task before the next starts. Fable coordinates and resolves questions. An implementer who finds the plan or design ambiguous stops and asks; it does not guess.

## Global Constraints

- Nothing in this repo can run VBA or PowerShell. Python tasks are verified by `python -m pytest python/tests -q` from the repo root. VBA tasks are verified by `python tools/check_vba.py src/vba` (exit code 0) plus review against the Python source.
- No floating point anywhere in engine, fairness, or summary code, in either language. Ratios are compared as `a1*e2 < a2*e1`.
- Python engine code never iterates a `dict` or `set` to decide order; it iterates lists in sheet order.
- Python uses `s.strip(' ')` (never bare `strip()`) and `upper()`; VBA uses `UCase$(Trim$(s))` and `StrComp(a, b, vbBinaryCompare)`. No `Option Compare Text`.
- VBA files: ASCII only, CRLF line endings, first line `Attribute VB_Name = "<module>"`, then `Option Explicit`. No `Integer` type (use `Long`). Every scalar parameter is `ByVal`. One variable per `Dim` with its own `As`.
- VBA `And`/`Or` do not short-circuit: never combine a bounds check and an array access in one condition; nest the `If`s.
- VBA dictionaries are keyed only by `CLng(dateSerial)` or `UCase$(Trim$(text))`.
- Dates are whole-number serials: Excel serial in VBA (`CLng(dateValue)`), `date.toordinal() - 693594` in Python. Date text is always `yyyy-mm-dd`; VBA parses it with `DateSerial(CLng(Left$(t,4)), CLng(Mid$(t,6,2)), CLng(Right$(t,2)))`.
- Constants (identical names and values in both languages): kinds `DAY=0, EVE=1, NIGHT=2, SBY_DE=3, SBY_N=4`; masks `M_DAY=1, M_EVE=2, M_NIGHT=4, M_ALL=8`; cell codes `GAP=-1, GAP_SENIOR=-2, GAP_COAG=-3, EMPTY_CELL=-9`; `NODE_CAP=200000`.
- User-visible strings (exact): kind names `Day`, `Evening`, `Night`, `Standby D/E`, `Standby N`; group names `Day + evening`, `Night`; rule texts `No senior`, `No special coag`, `No eligible staff`, `Search cut short`; roster GAP cell text `GAP`.
- Commit after each task with a message describing the task. Do not push; the coordinator pushes.

## File Structure

| Path | Responsibility |
| --- | --- |
| `python/roster/model.py` | Constants, `Slot`, `Settings`, `Inputs`, `YearResult`, `build_slots`, date helpers |
| `python/roster/csvio.py` | Read a scenario folder into `RawTables`; write result CSVs |
| `python/roster/validate.py` | `validate(raw) -> list[str]` |
| `python/roster/inputs.py` | `build_inputs(raw) -> Inputs` |
| `python/roster/rules.py` | `blocking_rules`, `forward_night_block` |
| `python/roster/fairness.py` | `eligible_dates`, `ratio_less`, `candidate_order` |
| `python/roster/engine.py` | `new_result`, `max_matching`, `lower_bound`, `solve_date`, `solve_year` |
| `python/roster/review.py` | `build_review`, `check_rules`, `build_summary` |
| `python/roster/ops.py` | `expand_weekends`, `roster_to_lastyear`, `approval_broken_rules`, `apply_approval` |
| `python/roster/gen_fixtures.py` | Regenerates `fixtures/*/expected_*.csv` and `fixtures/testsheet.csv` |
| `python/tests/` | `helpers.py` and one test file per module |
| `tools/check_vba.py` | Static checks on VBA sources |
| `fixtures/<scenario>/` | Scenario inputs and expected outputs |
| `src/vba/mod*.bas` | VBA modules listed in design section 10 |
| `build/build.ps1` | Builds `dist/PHRosterPlanner.xlsm` on Windows |
| `README.md` | How to build and use on Windows |

Dependencies: Task 2 is independent. Tasks 3 and 4 both follow 1 and may run in parallel. 6 follows 5; 7 follows 6. 11 and 12 follow 10 (12 also needs 11's `BuildReview`). Everything else is sequential.

---

### Task 1: Project scaffold and model

**Files:**
- Create: `python/roster/__init__.py` (empty), `python/roster/model.py`, `python/tests/__init__.py` (empty), `python/tests/helpers.py`, `python/tests/test_model.py`, `pytest.ini`, `.gitattributes`, `.gitignore`

**Interfaces — Produces:**

```python
# python/roster/model.py
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
```

```python
# python/tests/helpers.py
from roster.model import Settings, Inputs, build_slots, to_serial

def make_inputs(staff, dates, counts=(2, 1, 2, 1, 1), rules_off=(), exempt=(), leave=(),
                holiday=None, year=2027):
    """staff: list of (name, flags). flags letters: S senior, C special coag,
    d / e / n cannot do Day / Evening / Night.
    dates: list of 'yyyy-mm-dd' (ascending). holiday: list of names (default 'H<i>').
    exempt: iterable of (name, date_index). leave: iterable of (name, date_index, mask)."""
    names = [n for n, _ in staff]
    rules = {r: (r not in rules_off) for r in (1, 2, 4, 5, 6, 7)}
    cannot = [(1 if 'd' in f else 0) | (2 if 'e' in f else 0) | (4 if 'n' in f else 0)
              for _, f in staff]
    if not rules[6]:
        cannot = [0] * len(staff)
    ex = [[False] * len(dates) for _ in staff]
    lv = [[0] * len(dates) for _ in staff]
    if rules[4]:
        for n, d in exempt:
            ex[names.index(n)][d] = True
    if rules[5]:
        for n, d, m in leave:
            lv[names.index(n)][d] |= m
    return Inputs(Settings(year, list(counts), rules), build_slots(list(counts)), names,
                  ['S' in f for _, f in staff], ['C' in f for _, f in staff], cannot,
                  [to_serial(t) for t in dates],
                  holiday or ['H%d' % i for i in range(len(dates))], ex, lv)
```

Only rules 1, 2 and 7 are checked by switch inside the engine; rules 4, 5, 6 are switched off by clearing their data in `Inputs`.

- [ ] **Step 1:** Create `pytest.ini`:

```ini
[pytest]
pythonpath = python
testpaths = python/tests
```

`.gitattributes`:

```
src/vba/*.bas text eol=crlf
build/*.ps1 text eol=crlf
*.csv text eol=lf
```

`.gitignore`:

```
__pycache__/
.pytest_cache/
dist/
```

- [ ] **Step 2:** Write `python/tests/test_model.py`:

```python
from roster.model import build_slots, to_serial, from_serial, norm

def test_default_layout_names():
    names = [s.name for s in build_slots([2, 1, 2, 1, 1])]
    assert names == ["Day 1", "Day 2", "Evening", "Night 1", "Night 2", "Standby D/E", "Standby N"]

def test_zero_count_drops_column_and_single_is_unnumbered():
    slots = build_slots([1, 0, 3, 2, 0])
    assert [s.name for s in slots] == ["Day", "Night 1", "Night 2", "Night 3",
                                       "Standby D/E 1", "Standby D/E 2"]
    assert [s.group for s in slots] == [0, 1, 1, 1, 0, 0]
    assert [s.active for s in slots] == [True, True, True, True, False, False]
    assert [s.mask for s in slots] == [1, 4, 4, 4, 3, 3]

def test_serial_matches_excel():
    assert to_serial("2027-01-01") == 46388
    assert from_serial(46388) == "2027-01-01"
    assert to_serial("1900-03-01") == 61

def test_norm_trims_spaces_only_and_uppercases():
    assert norm("  Special Coag ") == "SPECIAL COAG"
    assert norm("\tx") == "\tX"
```

- [ ] **Step 3:** Run `python -m pytest python/tests -q`. Expected: import error (module missing).
- [ ] **Step 4:** Create `model.py` and `helpers.py` exactly as in Interfaces.
- [ ] **Step 5:** Run `python -m pytest python/tests -q`. Expected: 4 passed.
- [ ] **Step 6:** Commit: `git add -A && git commit -m "Add project scaffold and roster model"`.

---

### Task 2: VBA static checker

**Files:**
- Create: `tools/check_vba.py`, `python/tests/test_check_vba.py`

**Interfaces — Produces:** `check_text(name: str, data: bytes) -> list[str]` (one message per problem, formatted `<name>:<line>: <text>`), and a CLI `python tools/check_vba.py <dir>` that checks every `*.bas`, prints messages, and exits 1 if any (0 if none, also 0 for an empty folder).

Checks, in this order per file:

1. Any byte above 127: `non-ASCII character`.
2. Any `\n` not preceded by `\r`: `line ending is not CRLF` (report the first only).
3. Line 1 is not `Attribute VB_Name = "<name>"`: `missing or wrong Attribute VB_Name`.
4. No line equal to `Option Explicit`: `missing Option Explicit` (line 1).
5. A line equal to `Option Compare Text`: `Option Compare Text is not allowed`.
6. `As Integer`: `use Long, not Integer`.
7. A `Dim`/`Private`/`Public` variable declaration line containing a comma where any declared name lacks ` As `: `one typed variable per declaration`.
8. A procedure parameter list (`Sub`/`Function` line, joined across ` _` continuations) containing a parameter that starts with none of `ByVal`, `ByRef`, `Optional ByVal`: `parameter needs ByVal or ByRef`.
9. Block balance: `Sub`/`End Sub`, `Function`/`End Function`, `For`/`Next`, `Do`/`Loop`, `With`/`End With`, `Select Case`/`End Select`, and multi-line `If ... Then` (code line ends with `Then`) / `End If`. Report `unbalanced <keyword>` with the line of the unclosed opener or the stray closer.
10. A variable or parameter named `Date`, `Day`, `Year`, `Name`, `Weekday` or `Month`: `reserved name used as identifier`.

For checks 5 to 10, first strip from each line everything after an unquoted `'` and replace string literals with `""`.

- [ ] **Step 1:** Write `python/tests/test_check_vba.py` with one test per check, starting with:

```python
import importlib.util, pathlib
spec = importlib.util.spec_from_file_location(
    "check_vba", pathlib.Path(__file__).resolve().parents[2] / "tools" / "check_vba.py")
check_vba = importlib.util.module_from_spec(spec); spec.loader.exec_module(check_vba)

def src(*lines, name="modX"):
    head = ['Attribute VB_Name = "%s"' % name, "Option Explicit"]
    return "\r\n".join(head + list(lines) + [""]).encode("ascii")

def test_clean_module_has_no_findings():
    data = src("Public Sub A(ByVal n As Long)", "    Dim i As Long",
               "    For i = 1 To n", "        If i > 1 Then", "            i = i",
               "        End If", "    Next i", "End Sub")
    assert check_vba.check_text("modX", data) == []

def test_integer_is_reported_with_line():
    msgs = check_vba.check_text("modX", src("Private x As Integer"))
    assert msgs == ["modX:3: use Long, not Integer"]
```

Add in the same style: non-ASCII byte; bare `\n`; wrong `VB_Name`; missing `Option Explicit`; `Option Compare Text`; `Dim a, b As Long`; parameter without `ByVal`; `If x Then` with no `End If`; single-line `If x Then y = 1` (must NOT be reported); `For` without `Next`; a variable named `Name`; the word `Integer` inside a comment and inside a string (must NOT be reported).

- [ ] **Step 2:** Run the tests; expected: failures (module missing).
- [ ] **Step 3:** Implement `tools/check_vba.py`.
- [ ] **Step 4:** Run `python -m pytest python/tests -q`; expected: all pass. Run `mkdir -p src/vba && python tools/check_vba.py src/vba; echo $?`; expected `0`.
- [ ] **Step 5:** Commit: `git add -A && git commit -m "Add VBA static checker"`.

---

### Task 3: CSV loading, validation, building inputs

**Files:**
- Create: `python/roster/csvio.py`, `python/roster/validate.py`, `python/roster/inputs.py`, `python/tests/test_validate.py`, `python/tests/test_inputs.py`

**Interfaces — Consumes:** Task 1 model. **Produces:**

```python
# csvio.py
@dataclass
class RawTables:
    staff: list      # rows of [Name, Title, Senior, Role, Cannot do]   (strings, no header)
    holidays: list   # rows of [Holiday name, Date, Added]
    lastyear: list   # rows of [Holiday name, Date, Staff name, Slot]
    leave: list      # rows of [Staff name, From date, To date, Shift]
    settings: list   # rows of [Key, Value]
def read_scenario(folder) -> RawTables      # staff.csv, holidays.csv, lastyear.csv, leave.csv, settings.csv
def write_rows(path, header, rows) -> None  # csv module, lineterminator '\n', UTF-8
# validate.py
def validate(raw) -> list[str]              # empty list = OK
# inputs.py
def build_inputs(raw) -> Inputs             # call only when validate(raw) == []
```

CSV files have a header row exactly equal to the column names above; dates are `yyyy-mm-dd`; a missing `lastyear.csv` or `leave.csv` means an empty table. Rows that are entirely blank are dropped on read. Short rows are padded with `""`. Every cell is kept as a string.

Validation messages (exact formats; `<r>` is the 1-based data row number; report every problem, tables in the order Settings, Staff, Holidays, LastYear, Leave):

- `Settings: missing key <key>`
- `Settings: Year must be a whole number`
- `Settings: <key> must be a whole number from 0 to 9` (count keys)
- `Settings: <key> must be Y or N` (rule keys)
- `Staff row <r>: name is blank`
- `Staff row <r>: duplicate name <name>`
- `Staff row <r>: name must use plain ASCII characters`
- `Staff row <r>: Senior must be Y or N`
- `Staff row <r>: unknown Cannot do value <word>`
- `Holidays row <r>: holiday name is blank`
- `Holidays row <r>: holiday name must use plain ASCII characters`
- `Holidays row <r>: invalid date`
- `Holidays row <r>: date is outside year <year>`
- `Holidays row <r>: duplicate date <yyyy-mm-dd>`
- `LastYear row <r>: unknown staff name <name>`
- `LastYear row <r>: holiday name <name> is not on the Holidays sheet`
- `LastYear row <r>: Slot is blank`
- `Leave row <r>: unknown staff name <name>`
- `Leave row <r>: invalid date`
- `Leave row <r>: From date is after To date`
- `Leave row <r>: unknown Shift <word>`

A date is valid when it matches `dddd-dd-dd` and is a real calendar date. Duplicate, lookup, Y/N and word comparisons use `norm()`. Cannot do is split on commas; empty pieces are ignored. The "outside year" check is skipped when Year itself is invalid.

`build_inputs` rules:

- `settings.rules[r]` = value is `Y`. `counts` in key order.
- Staff in sheet order. `senior` = `Y`. `coag` = some comma-separated piece of Role has `norm(piece) == "SPECIAL COAG"`. `cannot` = OR of masks for `DAY`/`EVENING`/`NIGHT`; forced to 0 when rule 6 is off.
- Dates: Holidays rows sorted ascending by serial; `holiday[d]` is the name as entered (trimmed of spaces).
- `exempt[s][d]` = rule 4 on and LastYear has a row with that staff, `norm(holiday name) == norm(holiday[d])`, and `norm(Slot)` not starting with `STANDBY`.
- `leave[s][d]` = rule 5 on: OR, over that staff's Leave rows covering the date, of `M_ALL`/`M_DAY`/`M_EVE`/`M_NIGHT` for `ALL`/`DAY`/`EVENING`/`NIGHT`.

- [ ] **Step 1:** Write `test_validate.py`: one test per message above, each starting from a small valid `RawTables` built by a local `good()` helper (3 staff, 2 holidays in 2027, 1 lastyear row, 1 leave row, full settings) and changing one cell; assert the exact message is in the result. Plus `test_good_tables_have_no_messages` and `test_all_problems_reported_together` (two unrelated faults give exactly two messages).
- [ ] **Step 2:** Write `test_inputs.py` with these tests, each with literal expected lists:
  - `test_exemption_by_holiday_name_ignores_standby_and_case`: lastyear `[" good friday ", "2026-04-03", "Ann", "Day 1"]` and `["Good Friday", "2026-04-03", "Bob", "Standby N"]`; holidays New Year's Day 2027-01-01, Good Friday 2027-03-26 and 2027-03-27 → `exempt[Ann] == [False, True, True]`, `exempt[Bob] == [False, False, False]`.
  - `test_rule4_off_clears_exemptions`.
  - `test_leave_masks_cover_range_and_combine`: Ann has `Day` 2027-03-26..2027-03-27 and `Night` 2027-03-27..2027-03-27 → `leave[Ann] == [0, 1, 5]`.
  - `test_rule5_and_rule6_off_clear_masks`.
  - `test_role_list_detects_special_coag`: `"Haem, special coag "` is a match; `"Haem, special  coag"` (two spaces) and `"Special Coagulation"` are not.
  - `test_dates_sorted_and_names_follow`: holidays entered out of order.
  - `test_read_scenario_roundtrip(tmp_path)`: write CSVs with `write_rows`, read back, compare.
- [ ] **Step 3:** Run tests; expected: failures. Implement the three modules. Run again; expected: all pass.
- [ ] **Step 4:** Commit: `git add -A && git commit -m "Add CSV loading, validation and input building"`.

---

### Task 4: Rules and fairness

**Files:**
- Create: `python/roster/rules.py`, `python/roster/fairness.py`, `python/tests/test_rules.py`, `python/tests/test_fairness.py`

**Interfaces — Consumes:** Task 1 model and `make_inputs`. **Produces:**

```python
# rules.py
def blocking_rules(inp, grid, s, d, k) -> list[int]:
    """Rules among 3,4,5,6,7 that stop staff s taking slot k on date index d, ascending.
    The cell (d, k) itself is ignored, so this works for an already-filled cell.
    3: s is in another cell of grid[d].
    4: slot k is active and inp.exempt[s][d].
    5: inp.leave[s][d] has M_ALL, or shares a bit with the slot mask.
    6: inp.cannot[s] shares a bit with the slot mask.
    7: rule 7 is on, d > 0, inp.dates[d-1] == inp.dates[d] - 1, s is in an active NIGHT cell of
       grid[d-1], and slot k's kind is not NIGHT and not SBY_N."""
def forward_night_block(inp, grid, approved, s, d, k) -> bool:
    """True when rule 7 is on, slot k's kind is NIGHT, date d+1 exists and is the next calendar
    day, and s sits in an approved cell of grid[d+1] whose kind is not NIGHT and not SBY_N."""
# fairness.py
def eligible_dates(inp) -> list[int]:
    """Per staff: number of d with not inp.exempt[s][d] and not (inp.leave[s][d] & M_ALL)."""
def ratio_less(a1, e1, a2, e2) -> bool:   # is a1/e1 strictly lower than a2/e2; requires e1, e2 > 0
    return a1 * e2 < a2 * e1
def candidate_order(inp, active_count, standby_count, eligible) -> list[int]:
    """Staff indexes sorted by: (eligible == 0) last; then ratio ascending (cross-multiplied);
    then standby_count ascending; then norm(name) ascending. Implement as an insertion sort with
    an explicit before(x, y) comparison so the VBA port is line for line."""
```

- [ ] **Step 1:** Write `test_rules.py`. Use `make_inputs` and hand-built grids (`[[EMPTY_CELL] * 7 for _ in dates]`). Tests, each with explicit asserts:
  - rule 3: staff in another cell of the date → `[3]`; in the same cell only → `[]`.
  - rule 4: exempt blocks Day, Evening, Night; does not block either standby.
  - rule 5: `M_DAY` leave blocks Day and Standby D/E only; `M_EVE` blocks Evening and Standby D/E; `M_NIGHT` blocks Night and Standby N; `M_ALL` blocks all seven.
  - rule 6: flags `d`, `e`, `n` block the same patterns.
  - rule 7: active Night on 2027-02-06, next date 2027-02-07 → blocks Day, Evening, Standby D/E; allows Night and Standby N. Standby N on the first date blocks nothing. A two-day gap between dates blocks nothing. `rules_off=(7,)` blocks nothing.
  - several rules at once come back ascending, e.g. `[4, 6]`.
  - `forward_night_block`: true only for an approved next-day Day / Evening / Standby D/E cell; false when that cell is not approved, when the next-day slot is Night or Standby N, when dates are not consecutive, when slot k is not a Night slot, or when rule 7 is off.
- [ ] **Step 2:** Write `test_fairness.py`:

```python
from roster.fairness import eligible_dates, ratio_less, candidate_order
from roster.model import M_ALL, M_DAY
from tests.helpers import make_inputs

def test_eligible_dates_excludes_exempt_and_full_day_leave_only():
    inp = make_inputs([("Ann", ""), ("Bob", "")], ["2027-01-01", "2027-02-06", "2027-02-07"],
                      exempt=[("Ann", 0)], leave=[("Ann", 1, M_ALL), ("Bob", 1, M_DAY)])
    assert eligible_dates(inp) == [1, 3]

def test_ratio_less_uses_cross_multiplication():
    assert ratio_less(1, 3, 1, 2) and not ratio_less(1, 2, 1, 3) and not ratio_less(2, 4, 1, 2)

def test_order_ratio_then_standby_then_name_zero_eligible_last():
    inp = make_inputs([("Dan", ""), ("bob", ""), ("Ann", ""), ("Cat", ""), ("Eve", "")],
                      ["2027-01-01"])
    #            Dan bob Ann Cat Eve
    active =    [1,  0,  0,  0,  0]
    standby =   [0,  1,  0,  0,  0]
    eligible =  [4,  4,  4,  0,  4]
    assert [inp.staff[i] for i in candidate_order(inp, active, standby, eligible)] == \
        ["Ann", "Eve", "bob", "Dan", "Cat"]
```

Add: two zero-eligible staff ordered by standby then name; equal ratios with different denominators (1/2 vs 2/4) fall through to standby count.

- [ ] **Step 3:** Run; expect failures. Implement. Run; expect pass.
- [ ] **Step 4:** Commit: `git add -A && git commit -m "Add eligibility rules and fairness ordering"`.

---

### Task 5: Engine

**Files:**
- Create: `python/roster/engine.py`, `python/tests/test_engine.py`

**Interfaces — Consumes:** Tasks 1 and 4. **Produces:**

```python
def new_result(inp) -> YearResult      # grid all EMPTY_CELL, approved all False, counts 0,
                                       # eligible = eligible_dates(inp), cut_short False, unmet [], nodes 0
def max_matching(n_left, adj) -> int   # adj[i] = list of right-side ints for left node i; Kuhn's algorithm
def solve_date(inp, res, d) -> None    # fills grid[d]; sets cut_short[d], unmet[d], nodes[d]; does NOT touch counts
def solve_year(inp, res, from_d) -> None
```

`solve_year(inp, res, from_d)`:

1. For every `d >= from_d` and slot `k`: if not `approved[d][k]`, set `grid[d][k] = EMPTY_CELL`. Reset `cut_short[d] = False`, `unmet[d] = []`, `nodes[d] = 0`.
2. Recount `active_count` and `standby_count` from scratch over dates `< from_d` (cells holding a staff index; active slot → active, else standby).
3. For `d` from `from_d` upward: `solve_date(inp, res, d)`, then add that date's filled cells (approved ones included) to the counts.

`solve_date(inp, res, d)` — implement exactly this. Keep the search state in one small class or module-level variables (not closures over many locals), so the VBA port can mirror it with module-level variables:

```
row   = res.grid[d]                      (approved cells hold staff, the rest EMPTY_CELL)
order = candidate_order(inp, res.active_count, res.standby_count, res.eligible)
used[s]    = s appears in row
elig[s][k] = (blocking_rules(inp, res.grid, s, d, k) minus rule 3) is empty
             and not forward_night_block(inp, res.grid, res.approved, s, d, k)
steps = []
for g in (0, 1):
    if the layout has at least one active slot in group g:
        if rules[1]: steps.append(("senior", g))
        if rules[2]: steps.append(("coag", g))
for k in column order: if row[k] == EMPTY_CELL: steps.append(("slot", k))
bound = lower_bound(...)                  (see below; computed before any placement)
score = (0, 0, 0); best = None; best_score = None; nodes = 0; stop = False; cut = False
unmet = []

rec(i):
    if i == len(steps):
        best = copy of row; best_unmet = copy of unmet; best_score = score
        if best_score == bound: stop = True
        return
    (kind, x) = steps[i]
    if kind == "slot":
        if row[x] != EMPTY_CELL: rec(i + 1); return
        for s in order:
            if not used[s] and elig[s][x]:
                try_place(i, s, x, (0, 0, 0));  if stop: return
        try_gap(i, x, GAP, None);  return
    if kind == "senior":
        g = x
        if some active slot k of g has row[k] >= 0 and senior[row[k]]: rec(i + 1); return
        for s in order:
            if senior[s] and not used[s]:
                for k in active slots of g, column order:
                    if row[k] == EMPTY_CELL and elig[s][k]:
                        try_place(i, s, k, (0, 0, 0));  if stop: return
        k = last active slot of g with row[k] == EMPTY_CELL, or None
        try_gap(i, k, GAP_SENIOR, g);  return
    if kind == "coag":
        g = x
        if some active slot k of g has row[k] >= 0 and coag[row[k]]: rec(i + 1); return
        if some standby slot k of g has row[k] >= 0 and approved[d][k] and coag[row[k]]: rec(i + 1); return
        for s in order:                                            # (a) active
            if coag[s] and not used[s]:
                for k in active slots of g, column order:
                    if row[k] == EMPTY_CELL and elig[s][k]:
                        try_place(i, s, k, (0, 0, 0));  if stop: return
        k = first standby slot of g with row[k] == EMPTY_CELL, or None
        if k is not None:                                          # (b) fallback
            for s in order:
                if coag[s] and not used[s] and elig[s][k]:
                    try_place(i, s, k, (0, 0, 1));  if stop: return
        if k is None: k = last active slot of g with row[k] == EMPTY_CELL, or None   # (c) unmet
        try_gap(i, k, GAP_COAG, g);  return

try_place(i, s, k, delta):
    new = score + delta
    if best_score is not None and new >= best_score: return        # tuple comparison
    nodes += 1
    if nodes > NODE_CAP: stop = True; cut = True; return
    row[k] = s; used[s] = True; old = score; score = new
    rec(i + 1)
    row[k] = EMPTY_CELL; used[s] = False; score = old

try_gap(i, k, code, g):
    pen = 0 if code == GAP else 2                     # unmet senior / coag requirement
    if k is None:            delta = (1, 1, pen)
    elif slot k is active:   delta = (1, 1, pen)
    else:                    delta = (1, 0, pen)
    new = score + delta
    if best_score is not None and new >= best_score: return
    nodes += 1
    if nodes > NODE_CAP: stop = True; cut = True; return
    if k is None: unmet.append((g, code))  else: row[k] = code
    old = score; score = new
    rec(i + 1)
    if k is None: unmet.pop()  else: row[k] = EMPTY_CELL
    score = old

rec(0)
res.grid[d] = best; res.unmet[d] = best_unmet; res.cut_short[d] = cut; res.nodes[d] = nodes
```

Notes: `stop` is checked after every `try_place` inside loops (as shown); `try_gap` is always the last option so needs no check after it. `best` is never `None` at the end because the first leaf is reached before any pruning. `NODE_CAP` must be read from the module at call time (`engine.NODE_CAP` via `from roster import model` then `model.NODE_CAP`, or a module-level name in `engine`) so a test can lower it; state which in a comment.

`lower_bound`, computed on the row before the search starts, where `E(s, k)` means `not used[s] and elig[s][k]`:

```
empty   = [k for k in columns if row[k] == EMPTY_CELL]
gapsA   = len(empty) - max_matching(len(empty), [[s for s in staff order if E(s, k)] for k in empty])
emptyA  = [k for k in empty if slot k is active]
activeA = len(emptyA) - max_matching(len(emptyA), [[s for s in staff order if E(s, k)] for k in emptyA])
reqG = reqA = pen = 0
for each ("senior", g) step:
    if no active slot of g already holds a senior:
        if there is no senior s and empty active slot k of g with E(s, k): reqG += 1; reqA += 1; pen += 2
for each ("coag", g) step:
    if not already satisfied (the same two skip tests as in rec):
        if there is no coag s and empty active slot k of g with E(s, k):
            k1 = first empty standby slot of g, or None
            if k1 is not None and some coag s has E(s, k1): pen += 1
            else:
                pen += 2
                reqG += 1
                if k1 is None: reqA += 1
bound = (max(gapsA, reqG), max(activeA, reqA), pen)
```

- [ ] **Step 1:** Write `python/tests/test_engine.py`, starting with:

```python
from roster.engine import new_result, solve_year, max_matching
from roster.model import GAP, GAP_SENIOR, GAP_COAG, EMPTY_CELL
from tests.helpers import make_inputs

def names(inp, res, d):
    out = []
    for v in res.grid[d]:
        out.append(inp.staff[v] if v >= 0 else {GAP: "GAP", GAP_SENIOR: "GAP:S", GAP_COAG: "GAP:C"}[v])
    return out

def solved(inp):
    res = new_result(inp); solve_year(inp, res, 0); return res

BASE = [("A1", "SC"), ("A2", "SC"), ("B1", ""), ("B2", ""), ("B3", ""), ("B4", ""), ("B5", ""), ("B6", "")]

def test_single_date_fills_in_candidate_order():
    inp = make_inputs(BASE, ["2027-01-01"]); res = solved(inp)
    assert names(inp, res, 0) == ["A1", "B1", "B2", "A2", "B3", "B4", "B5"]
    assert res.nodes[0] == 7 and not res.cut_short[0] and res.unmet[0] == []
    assert res.active_count == [1, 1, 1, 1, 1, 0, 0, 0]
    assert res.standby_count == [0, 0, 0, 0, 0, 1, 1, 0]

def test_second_date_prefers_lowest_ratio_then_fewest_standby():
    inp = make_inputs(BASE, ["2027-01-01", "2027-03-10"]); res = solved(inp)
    assert names(inp, res, 1) == ["A1", "B6", "B4", "A2", "B5", "B1", "B2"]

def test_only_senior_for_one_group_reserves_gap_and_uses_coag_standby():
    staff = [("S1", "SC"), ("S2", "SC")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-01-01"], exempt=[("S1", 0)]); res = solved(inp)
    assert names(inp, res, 0) == ["S2", "B1", "B2", "B3", "GAP:S", "B4", "S1"]
    assert not res.cut_short[0]

def test_no_coag_at_all_reserves_both_standby_slots():
    staff = [("S1", "S"), ("S2", "S")] + [("B%d" % i, "") for i in range(1, 7)]
    inp = make_inputs(staff, ["2027-01-01"]); res = solved(inp)
    assert names(inp, res, 0) == ["S1", "B1", "B2", "S2", "B3", "GAP:C", "GAP:C"]

def test_matching_values():
    assert max_matching(3, [[0], [0, 1], [1]]) == 2
    assert max_matching(2, [[], []]) == 0
    assert max_matching(3, [[0, 1], [0], [2]]) == 3
```

Then add these tests. For each, design the smallest staff list that produces the stated outcome, trace the algorithm above by hand, and write the literal expected roster in the test before implementing. If the implementation later disagrees with a hand trace, stop and report; do not edit the expectation to match.

  - **Coag fallback in both groups:** two coag staff, both exempt on the date, plus two seniors and six others → each standby cell holds a coag person, no GAP, `nodes[0] == 7`.
  - **Coag unmet with no standby slots:** `counts=(2, 1, 2, 0, 0)`, seniors but no coag → `Evening` and `Night 2` are `GAP:C`.
  - **Unmet without a slot:** `counts=(0, 0, 1, 0, 0)`, staff with no senior and no coag → roster `["GAP:S"]` and `res.unmet[0] == [(1, GAP_COAG)]`.
  - **Plain shortage stops at the bound:** 20 staff of whom 19 have flag `n`; the remaining one is `SC` → `Night 2` and `Standby N` are `GAP`, `cut_short[0] is False`, `nodes[0] < 100`.
  - **Backtracking:** a case where the first senior tried for group DE must be undone for a complete roster to exist (for example the lowest-ordered senior is the only person able to do Night) → no GAP in the result.
  - **Rest after night:** dates `2027-02-06` and `2027-02-07`, 10 staff → everyone in an active Night cell on day 1 is, on day 2, in a Night cell, in Standby N, or absent.
  - **Rule switches:** `rules_off=(1, 2)` with no seniors and no coag → no GAP.
  - **Approved cell kept:** after a full solve of two dates, set `res.grid[1] = [EMPTY_CELL] * 7`, put B6 in `grid[1][0]` with `approved[1][0] = True`, call `solve_year(inp, res, 1)` → B6 still in that cell, appears once on the date, and `active_count` includes it.
  - **Forward night block:** consecutive dates; B1 approved in `Day 1` on day 2; `solve_year(inp, res, 0)` → B1 is in no Night cell on day 1.
  - **Node cap:** lower the cap to 50 with `monkeypatch` on a case whose optimum is above its bound (the "only senior for one group" case above works) and assert `nodes[0] <= 51` and no `EMPTY_CELL` remains; if that case finishes below 50 nodes, lower the cap to 5 and assert `cut_short[0] is True` and `nodes[0] == 6`.
  - **Determinism:** solving the same inputs twice gives identical grids and node counts.
  - **Recount:** after a full solve, `solve_year(inp, res, 1)` leaves grid and counts unchanged.

- [ ] **Step 2:** Run; expect failures. Implement `engine.py`. Run; expect all pass.
- [ ] **Step 3:** Commit: `git add -A && git commit -m "Add roster search engine"`.

---

### Task 6: Review rows, rule checker, summary

**Files:**
- Create: `python/roster/review.py`, `python/tests/test_review.py`

**Interfaces — Consumes:** Tasks 1, 4, 5. **Produces:**

```python
@dataclass
class ReviewRow:
    date: int          # serial
    group: str         # "Day + evening", "Night" or ""
    slot: str          # slot name or ""
    rule: str          # RULE_TEXT value or CUT_SHORT_TEXT
    reason: str
    candidates: str
@dataclass
class Approval:
    d: int             # date index
    k: int             # slot index
    s: int             # staff index
    rules: list        # ascending rule numbers
def build_review(inp, res) -> list[ReviewRow]
def check_rules(inp, res, approvals) -> list[str]
def ratio_text(a, e) -> str                 # '' when e == 0; else 4 decimals, half up
def build_summary(inp, res) -> tuple[list[list[str]], str]   # rows, spread text
```

`build_review` order: dates ascending; within a date: a cut-short row first (if any), then GAP cells in column order, then `unmet[d]` entries in stored order.

- GAP cell row: `group = GROUP_NAMES[slot.group]`, `slot = slot.name`, `rule = RULE_TEXT[code]`.
- People considered: seniors for `GAP_SENIOR`, coag for `GAP_COAG`, everyone for `GAP`; staff sheet order.
- Reason: `"; "`-joined `"<name>: <why>"`, where why comes from the first rule in `blocking_rules(inp, res.grid, s, d, k)`: 3 `already rostered`, 4 `exempt`, 5 `on leave`, 6 `cannot do <kind name of the slot>`, 7 `rest after night`; no rule → `not placed`. If nobody is considered: `no staff with this requirement`.
- Candidates: people considered who are not in `grid[d]` and whose `blocking_rules` is exactly one rule from 4 to 7: `"<name> (rule 4 exemption)"`, `(rule 5 leave)`, `(rule 6 cannot do)`, `(rule 7 rest after night)`, joined with `"; "`.
- Unmet-without-slot row: group name, slot `""`, rule text, reason `no free slot in this group`, candidates `""`.
- Cut-short row: group `""`, slot `""`, rule `CUT_SHORT_TEXT`, reason `best roster found so far was kept`, candidates `""`.

`check_rules` returns, in date then column order, `"<yyyy-mm-dd> <slot name>: <staff> breaks rule <n>"` for each filled cell and each rule in `blocking_rules` not listed in an `Approval` for that exact cell. Then, per date and per group that has active slots, has no negative cell in any of its slots, and has no `unmet` entry for that group: `"<yyyy-mm-dd> <group name>: rule 1 not met"` when rule 1 is on and no active cell holds a senior, and `"<yyyy-mm-dd> <group name>: rule 2 not met"` when rule 2 is on and neither an active cell nor a standby cell of the group holds a coag person. Each is excused when any `Approval` for a cell of that group on that date lists that rule.

`ratio_text(a, e)`: `q = (a * 20000 + e) // (2 * e)`; text is `str(q // 10000) + "." + ("%04d" % (q % 10000))`.

`build_summary`: rows in staff order `[name, str(active), str(eligible), ratio_text, str(standby)]`. Spread text: with `q` computed as above for each staff with `eligible > 0`, format `max(q) - min(q)` the same way; `"0.0000"` when there is no such person.

- [ ] **Step 1:** Write tests with literal expected strings:
  - staff `S1` (`S`, exempt), `S2` (`S`), `C1` (`C`), `C2` (`C`), `B1`..`B5`, one date → one row, slot `Night 2`, rule `No senior`, reason `S1: exempt; S2: already rostered`, candidates `S1 (rule 4 exemption)` (S1 has zero eligible dates, sorts last and is not rostered);
  - the "only senior for one group" case from Task 5, where S1 ends up in Standby N → reason `S1: already rostered; S2: already rostered` and empty candidates;
  - the plain-shortage case → reason contains `cannot do Night`;
  - a person blocked by two rules is in Reason but not in Candidates;
  - unmet-without-slot row and cut-short row (set `res.cut_short[0] = True` by hand);
  - `check_rules` returns `[]` on engine output for three different inputs;
  - hand-edit a solved grid to create each of rule 3, 4, 5, 6, 7, 1 and 2 violations and assert the exact message; an `Approval` excuses exactly its recorded rule and nothing else;
  - `ratio_text(1, 3) == "0.3333"`, `(2, 3) → "0.6667"`, `(1, 8) → "0.1250"`, `(1, 16) → "0.0625"`, `(3, 2) → "1.5000"`, `(0, 0) → ""`;
  - summary rows and spread for the two-date `BASE` solve (spread `0.5000`).
- [ ] **Step 2:** Run; fail. Implement. Run; pass.
- [ ] **Step 3:** Commit: `git add -A && git commit -m "Add review rows, rule checker and summary"`.

---

### Task 7: Long weekends, new year, approvals

**Files:**
- Create: `python/roster/ops.py`, `python/tests/test_ops.py`

**Interfaces — Consumes:** Tasks 1, 4, 5, 6 (`Approval`, `check_rules`). **Produces:**

```python
def expand_weekends(rows, year) -> tuple[list, list]:
    """rows: Holidays rows [name, 'yyyy-mm-dd', added]. Returns (rows sorted by date, skipped)
    where skipped lists 'yyyy-mm-dd' dates not added because they fall outside `year`.
    Only rows whose added cell is not Y (norm) are expanded. Weekday via date.isoweekday():
    Fri(5): +1, +2.  Mon(1): -2, -1.  Sat(6): +1, +2.  Sun(7): -1, +1.
    A date already present (original, or added earlier in this call) is not added again.
    New rows are [same name, date, 'Y']. The sort is stable by date."""
def roster_to_lastyear(inp, res) -> list:
    """Rows [holiday name, 'yyyy-mm-dd', staff name, slot name] for every filled active cell,
    dates ascending then column order."""
def approval_broken_rules(inp, res, d, k, s) -> list[int]:
    """Ascending, unique: blocking_rules(inp, res.grid, s, d, k) minus rule 3; plus 7 if
    forward_night_block(inp, res.grid, res.approved, s, d, k); plus 1 if grid[d][k] == GAP_SENIOR
    and not senior[s]; plus 2 if grid[d][k] == GAP_COAG and not coag[s]."""
def apply_approval(inp, res, approvals, d, k, s) -> tuple[bool, str]:
    """Fails with (False, 'slot is no longer a GAP') when grid[d][k] >= 0, and with
    (False, '<name> already has a slot on this date') when s is in grid[d]. Otherwise:
    append Approval(d, k, s, approval_broken_rules(...)), set the cell and its approved flag,
    add the person to the counts by calling solve_year(inp, res, d + 1) (which recounts), then
    for each other approval a with a.d > d: extra = blocking_rules(inp, res.grid, a.s, a.d, a.k)
    minus rule 3 minus a.rules; if extra is not empty, merge it into a.rules (ascending).
    Returns (True, msg) where msg is '' or
    'approval for <name> on <yyyy-mm-dd> now also breaks rule <n>[, <n>]' joined by '; '."""
```

When `d` is the last date, `solve_year(inp, res, d + 1)` solves nothing but still recounts; make sure it accepts `from_d == len(dates)`.

- [ ] **Step 1:** Write tests with literal expectations:
  - Friday 2027-03-26 → adds 27th and 28th; Monday 2027-08-09 → adds 7th and 8th; CNY Saturday 2027-02-06 plus Sunday 2027-02-07 → adds only Monday 2027-02-08, once; Friday 2027-01-01 → adds 2nd and 3rd; Saturday 2027-12-25 → adds 26th and 27th; a synthetic Friday 2027-12-31 → adds nothing and `skipped == ["2028-01-01", "2028-01-02"]`;
  - running `expand_weekends` on its own output changes nothing; rows marked `Y` are not expanded;
  - `roster_to_lastyear` excludes standby and GAP cells;
  - `apply_approval`: approving exempt S1 into the `GAP:S` cell of the Task 5 case (extended to three dates) records rules `[4]`, leaves the other cells of that date unchanged, re-solves later dates, and `check_rules(inp, res, approvals) == []`;
  - approving a non-senior into `GAP:S` records `[1]`;
  - both failure messages;
  - an approval already recorded on a later date stays in its cell after an earlier-date approval is applied.
- [ ] **Step 2:** Run; fail. Implement. Run; pass.
- [ ] **Step 3:** Commit: `git add -A && git commit -m "Add long weekend, new year and approval operations"`.

---

### Task 8: Fixtures and scenario tests

**Files:**
- Create: `fixtures/<scenario>/{staff,holidays,lastyear,leave,settings}.csv` for the 9 scenarios below, `fixtures/approval/approve.csv`, `python/roster/gen_fixtures.py`, `python/tests/test_scenarios.py`
- Generated and committed: `fixtures/<scenario>/expected_{roster,review,summary}.csv`, `fixtures/testsheet.csv`

**Interfaces — Consumes:** Tasks 1 to 7. **Produces:**

- `python python/roster/gen_fixtures.py [out_root]` (default `fixtures/`; inputs are always read from `fixtures/`). It must add `python/` to `sys.path` itself so it runs from the repo root. For each folder in `fixtures/` sorted by name: read, validate (must be empty, else exit 1 with the messages), build inputs, `solve_year(inp, res, 0)`, apply `approve.csv` rows in file order if present (columns `Date,Slot,Staff name`, resolved to indexes; a failed approval is an error), then write:
  - `expected_roster.csv`: header `Holiday,Date,<slot names…>`; one row per date; cells are staff names or `GAP`.
  - `expected_review.csv`: header `Date,Group,Slot,Rule not met,Reason,Candidates`.
  - `expected_summary.csv`: header `Staff name,Active slots,Eligible dates,Ratio,Standby slots`, staff rows, then a final row `Spread,,,<spread>,`.
  - `<out_root>/testsheet.csv`: no header; columns `Scenario,Table,RowNo,V1..Vn`, every line padded with empty strings to the widest line. Tables per scenario in this order: `staff`, `holidays`, `lastyear`, `leave`, `settings`, `approve` (if any), `expected_roster`, `expected_review`, `expected_summary`. `RowNo` 0 is the header row; data rows count from 1.
- Scenario folders. Each test asserts the stated property, not only equality with generated files:

| Folder | Content | Asserted property |
| --- | --- | --- |
| `normal_year` | 30 fictional staff (6 seniors, 5 special coag, 2 of them both, a few Cannot do), the 2027 holidays below, LastYear with 5 active names and 2 standby names per holiday date, 4 leave rows, default settings | no GAP, no Review rows, `check_rules == []`, spread ≤ `0.5000` |
| `senior_exempt` | 10 staff, one date; one of the two seniors is exempt, so only one group can have a senior | exactly one Review row, rule `No senior`, the exempt senior named in Candidates |
| `coag_standby` | 10 staff, one date; all special coag staff exempt | no GAP; both standby cells hold special coag staff |
| `leave_cannot` | 12 staff, two dates, one leave row of each Shift type and several Cannot do flags | `check_rules == []`; each named person is absent from the cells their leave or Cannot do blocks |
| `long_weekend` | Good Friday 2027-03-26 plus rows for the 27th and 28th marked Added, 14 staff, LastYear rows for Good Friday | nobody active on Good Friday last year is active on any of the 3 dates |
| `cny` | Chinese New Year 2027-02-06 and 2027-02-07, 14 staff | no GAP; rule 7 holds across the two dates |
| `rest_after_night` | three consecutive dates, 9 staff | `check_rules == []`; everyone on active Night on a date is on Night, Standby N or absent the next date |
| `custom_counts` | counts 3,2,1,0,2; Rule 1 = N; 12 staff, one date | roster header slot columns are `Day 1,Day 2,Day 3,Evening 1,Evening 2,Night,Standby N 1,Standby N 2` |
| `approval` | `senior_exempt` staff over three dates, plus `approve.csv` placing the exempt senior in the GAP | the approved name is in that cell; `check_rules == []` given that approval |

- `normal_year/holidays.csv` (sample dates, not verified against the official list):

```
Holiday name,Date,Added
New Year's Day,2027-01-01,
Chinese New Year,2027-02-06,
Chinese New Year,2027-02-07,
Hari Raya Puasa,2027-03-10,
Good Friday,2027-03-26,
Labour Day,2027-05-01,
Hari Raya Haji,2027-05-17,
Vesak Day,2027-05-20,
National Day,2027-08-09,
Deepavali,2027-10-28,
Christmas Day,2027-12-25,
```

  `gen_fixtures` does not expand weekends; `normal_year` ships unexpanded so the user can try the button.
- Staff names are short fictional ASCII names (for example `Tan Wei`, `Lim Hui`); no real people.

- [ ] **Step 1:** Write the input CSVs for all 9 scenarios.
- [ ] **Step 2:** Write `gen_fixtures.py` and run it. Read each `expected_roster.csv` and `expected_review.csv` against the "Asserted property" column; adjust the input CSVs (never the engine) until each scenario shows its property. If a scenario cannot be made to show its property, stop and report.
- [ ] **Step 3:** Write `test_scenarios.py`: one test per scenario asserting its property; `test_expected_files_are_current(tmp_path)` runs the generator into `tmp_path` and compares every generated file byte for byte with the committed one; `test_no_scenario_is_cut_short`; `test_testsheet_lists_every_scenario`.
- [ ] **Step 4:** Run `python -m pytest python/tests -q`; expect all pass.
- [ ] **Step 5:** Commit: `git add -A && git commit -m "Add scenario fixtures and expected results"`.

---

### VBA conventions for Tasks 9 to 13

Every VBA task follows these in addition to Global Constraints.

- State lives in `Public` module-level variables in `modTypes`, arrays 0-based, mirroring the Python fields: `gYear As Long`, `gCounts(0 To 4) As Long`, `gRuleOn(1 To 7) As Boolean`, `gSlotCount As Long`, `gSlotKind() As Long`, `gSlotGroup() As Long`, `gSlotActive() As Boolean`, `gSlotName() As String`, `gSlotMask() As Long`, `gStaffCount As Long`, `gStaff() As String`, `gSenior() As Boolean`, `gCoag() As Boolean`, `gCannot() As Long`, `gDateCount As Long`, `gDates() As Long`, `gHoliday() As String`, `gExempt() As Boolean` `(s, d)`, `gLeave() As Long` `(s, d)`, `gGrid() As Long` `(d, k)`, `gApproved() As Boolean` `(d, k)`, `gActiveCount() As Long`, `gStandbyCount() As Long`, `gEligible() As Long`, `gCutShort() As Boolean`, `gNodes() As Long`, `gUnmetCount() As Long` `(d)`, `gUnmetGroup() As Long` `(d, 0 To 3)`, `gUnmetCode() As Long` `(d, 0 To 3)`; approvals `gApprCount As Long`, `gApprD() As Long`, `gApprK() As Long`, `gApprS() As Long`, `gApprRules() As String` (comma-separated ascending, e.g. `"4,7"`).
- Raw tables are `Public gRawStaff() As String` and so on: 2-D, `(1 To rows, 1 To cols)`, each with a companion row count (`gRawStaffRows As Long`) because a table can have zero rows. Never call `UBound` on a table whose count is 0. Dates in raw tables are `yyyy-mm-dd` text.
- Each Python function is ported to one VBA procedure with the same name in PascalCase (`blocking_rules` → `BlockingRules`). Port line by line; keep every loop order identical. Short lists of rule numbers are returned as a comma-separated `String`; other lists are filled into a `ByRef` array with a `ByRef` count.
- The search uses module-level `Private` variables in `modEngine` for the row, `used`, `elig`, order, steps, the three score parts, best copies, node count, stop flag and cut flag. Undo runs on every return path. The stop flag is checked after every recursive call.
- Reading a sheet range: use `ws.Range(...).Value2` as an array only when the range has 2 or more cells; handle the single-cell and no-data cases separately.
- Never `Select` or `Activate`. Long messages go through `modMain.ShowMessages(title, text)`: `MsgBox` for up to 900 characters; otherwise write the lines to the Review sheet from cell I2 downward under a header `Messages` in I1 and tell the user to look there.
- After writing each module run `python tools/check_vba.py src/vba` and fix every finding.

---

### Task 9: VBA types, sheet I/O, validation, inputs

**Files:**
- Create: `src/vba/modTypes.bas`, `src/vba/modIO.bas`, `src/vba/modValidate.bas`

**Interfaces — Consumes:** Python `model.py`, `csvio.py`, `validate.py`, `inputs.py` (port source). **Produces:**

- `modTypes`: all constants from Global Constraints as `Public Const <NAME> As Long = <value>`; the `Public` state variables above; `Public Sub BuildSlots()` (from `gCounts`); `Public Function Norm(ByVal t As String) As String`; `Public Function SerialFromText(ByVal t As String) As Long`; `Public Function TextFromSerial(ByVal n As Long) As String` (`Format$(CDate(n), "yyyy-mm-dd")`); `Public Function IsDateText(ByVal t As String) As Boolean`; `Public Function KindName(ByVal kind As Long) As String`; `Public Function GroupName(ByVal g As Long) As String`; `Public Function RuleText(ByVal code As Long) As String`; `Public Function StaffIndex(ByVal t As String) As Long` (−1 when not found; compares `Norm`).
- `modIO`:
  - `Public Sub ReadSheetsToRaw()` — fills the raw tables from sheets Staff, Holidays, LastYear, Leave (data from row 2; last row found by scanning the used range; fully blank rows dropped; date cells converted with `TextFromSerial(CLng(v))` when `IsNumeric`, else `Trim$` of the text) and Settings (Key in A, Value in B, rows 2 to 13).
  - `Public Sub BuildInputs()` — port of `build_inputs`: raw tables → state, then `BuildSlots`.
  - `Public Sub WriteRoster()` — header row 1 `Holiday, Date, <slot names>`; clear contents, fills and notes of everything below and beside first; GAP cells get text `GAP` and fill `RGB(255, 199, 206)`; approved cells get fill `RGB(255, 235, 156)` and a note `Approved exception: rule <list>` (or `Approved exception: no rule broken`); dates written as serial with `NumberFormat = "yyyy-mm-dd"`.
  - `Public Sub WriteSummary()` — rows per staff plus the `Spread` row; uses `modReview.RatioText` and `modReview.SpreadText` (written in Task 11); Ratio column `NumberFormat = "@"`.
  - `Public Function LoadResultFromSheets() As String` — rebuilds `gGrid`, `gApproved`, `gCutShort`, `gUnmet*` and the approval arrays from Roster, Approvals and Review, then recounts `gActiveCount`, `gStandbyCount` and calls `modRules.ComputeEligible`. Roster cell with a staff name → index; `GAP` → code from the Review row with the same date and slot (`No senior` → `GAP_SENIOR`, `No special coag` → `GAP_COAG`, otherwise `GAP`); Review rows with an empty Slot → cut-short flag when the rule is `Search cut short`, else an unmet entry; Approvals rows → approval arrays and `gApproved`. Returns `""` on success, or a message when Roster dates or slot headers do not match Holidays and Settings or a Roster name is not on Staff (the caller then stops and tells the user to run Generate roster again).
  - `Public Sub WriteApprovals()` — rewrites the Approvals sheet (`Date, Slot, Staff name, Rules broken`) from the approval arrays.
- `modValidate`: `Public Function ValidateRaw() As String` — port of `validate`; messages joined with `vbLf`; exactly the same texts and order.

- [ ] **Step 1:** Write `modTypes.bas`. Run the checker.
- [ ] **Step 2:** Write `modValidate.bas` as a port of `python/roster/validate.py`. Run the checker.
- [ ] **Step 3:** Write `modIO.bas`. Run the checker.
- [ ] **Step 4:** Compare each procedure with its Python source; list in the task report every place where the VBA differs in structure from the Python, and why.
- [ ] **Step 5:** Commit: `git add -A && git commit -m "Add VBA types, sheet I/O and validation"`.

---

### Task 10: VBA rules and engine

**Files:**
- Create: `src/vba/modRules.bas`, `src/vba/modEngine.bas`

**Interfaces — Consumes:** `modTypes` state; Python `rules.py`, `fairness.py`, `engine.py` (port source). **Produces:**

- `modRules`: `Public Function BlockingRules(ByVal s As Long, ByVal d As Long, ByVal k As Long) As String` (comma-separated ascending, `""` when none); `Public Function HasRule(ByVal ruleList As String, ByVal r As Long) As Boolean`; `Public Function WithoutRule(ByVal ruleList As String, ByVal r As Long) As String`; `Public Function ForwardNightBlock(ByVal s As Long, ByVal d As Long, ByVal k As Long) As Boolean`; `Public Sub ComputeEligible()` (fills `gEligible`); `Public Function RatioLess(ByVal a1 As Long, ByVal e1 As Long, ByVal a2 As Long, ByVal e2 As Long) As Boolean`; `Public Sub CandidateOrder(ByRef order() As Long)`.
- `modEngine`: `Public Sub NewResult()`; `Public Function MaxMatching(ByVal nLeft As Long, ByRef adjCount() As Long, ByRef adj() As Long) As Long`; `Public Sub SolveDate(ByVal d As Long)`; `Public Sub SolveYear(ByVal fromD As Long)`.

All products `a1 * e2` stay far below the `Long` limit; keep every operand `Long`.

- [ ] **Step 1:** Port `rules.py` and `fairness.py` to `modRules.bas`. Run the checker.
- [ ] **Step 2:** Port `engine.py` to `modEngine.bas`, keeping the same procedure split (`Rec`, `TryPlace`, `TryGap`, `LowerBound`, `MaxMatching`) and the same order of every loop and test. Run the checker.
- [ ] **Step 3:** Desk-check: trace `test_single_date_fills_in_candidate_order` and `test_only_senior_for_one_group_reserves_gap_and_uses_coag_standby` through the VBA by hand. Record in the task report the final roster and node count for each; they must equal the Python values (print `res.nodes[0]` from Python for the second).
- [ ] **Step 4:** Commit: `git add -A && git commit -m "Add VBA rules and search engine"`.

---

### Task 11: VBA review, checker, summary

**Files:**
- Create: `src/vba/modReview.bas`

**Interfaces — Consumes:** `modTypes`, `modRules`; Python `review.py` (port source). **Produces:** `Public gRevCount As Long` and `Public gRev() As String` `(1 To n, 1 To 6)` (date text, group, slot, rule, reason, candidates); `Public Sub BuildReview()`; `Public Sub WriteReview()` (sheet Review, header row 1 `Date, Group, Slot, Rule not met, Reason, Candidates, Approve`; old rows cleared including column G; date written as serial with `yyyy-mm-dd` format); `Public Function CheckRules() As String` (messages joined with `vbLf`); `Public Function RatioText(ByVal a As Long, ByVal e As Long) As String`; `Public Function SpreadText() As String`.

`RatioText` uses whole-number arithmetic only: `q = (a * 20000 + e) \ (2 * e)`; text `CStr(q \ 10000) & "." & Format$(q Mod 10000, "0000")`.

- [ ] **Step 1:** Port `review.py`. Run the checker.
- [ ] **Step 2:** Desk-check `RatioText` for (1,3), (2,3), (1,16), (3,2), (0,0) against the Python test values; record results in the task report.
- [ ] **Step 3:** Commit: `git add -A && git commit -m "Add VBA review, rule checker and summary"`.

---

### Task 12: VBA holidays, new year, approvals, entry points

**Files:**
- Create: `src/vba/modHolidays.bas`, `src/vba/modNewYear.bas`, `src/vba/modApprovals.bas`, `src/vba/modMain.bas`

**Interfaces — Consumes:** Tasks 9 to 11 modules; Python `ops.py` (port source). **Produces:**

- `modMain` button entry points, all `Public Sub` with no parameters: `GenerateRoster`, `ApplyApprovals`, `ExpandLongWeekends`, `StartNewYear`, `RunSelfTest` (calls `modSelfTest.RunSelfTestInteractive`, written in Task 13). Also `Public Sub ShowMessages(ByVal title As String, ByVal text As String)`.
- `modApprovals`: `Public Function ApprovalBrokenRules(ByVal d As Long, ByVal k As Long, ByVal s As Long) As String`; `Public Function ApplyApprovalCore(ByVal d As Long, ByVal k As Long, ByVal s As Long, ByRef note As String) As Boolean` (port of `apply_approval`; no dialogs).
- `modHolidays`: `Public Sub ExpandHolidaySheet()`. `modNewYear`: `Public Sub DoStartNewYear()`.

Behaviour:

- `GenerateRoster`: `ReadSheetsToRaw`; `ValidateRaw` → if not empty, `ShowMessages` and stop. If the Approvals sheet has data rows: `MsgBox(..., vbYesNo)` "Existing approvals will be discarded. Continue?"; stop on No. Then `BuildInputs`, `NewResult` (which also clears approvals), `SolveYear 0`, `BuildReview`, `WriteRoster`, `WriteReview`, `WriteSummary`, `WriteApprovals`, `CheckRules`; show findings if any, else `Roster generated. <n> item(s) to review.`
- `ExpandLongWeekends` → `ExpandHolidaySheet`: port of `expand_weekends` over the Holidays sheet. Use `Weekday(d, vbMonday)` (Monday = 1 … Sunday = 7, the same numbers as Python `isoweekday`). Needs a whole-number Year in Settings, else message and stop. Appends rows, sorts the data rows by the date column, reports skipped dates.
- `StartNewYear` → `DoStartNewYear`: confirm with `vbYesNo`; `ReadSheetsToRaw`, validate, `BuildInputs`, `LoadResultFromSheets` (stop on message); write LastYear from the port of `roster_to_lastyear` (replacing its data rows); clear the data rows (not headers) of Holidays, Leave, Roster, Review, Summary, Approvals; Year = Year + 1.
- `ApplyApprovals`: `ReadSheetsToRaw`, validate, `BuildInputs`, `LoadResultFromSheets` (stop on message). Collect Review rows with a non-empty Approve cell; rows with an empty Slot are reported as not approvable. Sort the rest by date, then slot index. For each: resolve the name (unknown → report, skip); if the cell is no longer a GAP or the person already has a slot that date → report, skip; compute `ApprovalBrokenRules`; `MsgBox(vbYesNo)` "`<name>` into `<slot>` on `<date>` breaks rule `<list>`. Approve?" (or "breaks no rule" when empty); on Yes call `ApplyApprovalCore`. At the end: `BuildReview`, `WriteRoster`, `WriteReview`, `WriteSummary`, `WriteApprovals`, `CheckRules`, and one message listing applied rows, skipped rows with reasons, and any "now also breaks" notes.

- [ ] **Step 1:** Write `modHolidays.bas`; checker.
- [ ] **Step 2:** Write `modNewYear.bas`; checker.
- [ ] **Step 3:** Write `modApprovals.bas`; checker.
- [ ] **Step 4:** Write `modMain.bas`; checker.
- [ ] **Step 5:** Commit: `git add -A && git commit -m "Add VBA holiday, new year, approval macros and entry points"`.

---

### Task 13: VBA self-test

**Files:**
- Create: `src/vba/modSelfTest.bas`

**Interfaces — Consumes:** Test sheet layout from Task 8 (`Scenario, Table, RowNo, V1..Vn`; all cells text; no header row), Tasks 9 to 12 modules. **Produces:** `Public Function SelfTestSilent() As String` (returns `PASS: <n> scenarios` or lines `FAIL <scenario>: <what differed>`; never shows a dialog; has an error handler that returns `FAIL <scenario>: runtime error <number> <description>`), and `Public Sub RunSelfTestInteractive()` (shows the result via `ShowMessages`).

Per scenario, in sheet order: fill the raw tables from the `staff`, `holidays`, `lastyear`, `leave`, `settings` rows with RowNo ≥ 1; `ValidateRaw` must be `""`; `BuildInputs`; `NewResult`; `SolveYear 0`; for each `approve` row call `ApplyApprovalCore` (must return True); `BuildReview`; then compare:

- every `expected_roster` cell with the roster built from state (holiday, date text, staff name or `GAP`), reporting the first mismatch as `roster <date> <slot>: got <x>, expected <y>`;
- `expected_review` row count and columns Date, Slot, Rule; first mismatch reported;
- `expected_summary` rows including the `Spread` row;
- `CheckRules` must be `""`.

The self-test reads no sheet other than Test and writes none. It leaves test data in the globals, which is fine because every button macro rebuilds state from the sheets first.

- [ ] **Step 1:** Write `modSelfTest.bas`; checker.
- [ ] **Step 2:** Commit: `git add -A && git commit -m "Add VBA self-test"`.

---

### Task 14: Build script and README

**Files:**
- Create: `build/build.ps1`, `README.md`

**Interfaces — Consumes:** `src/vba/*.bas`, `fixtures/normal_year/*.csv`, `fixtures/testsheet.csv`, the entry-point names from Task 12, `modSelfTest.SelfTestSilent`.

`build.ps1` steps (PowerShell 5.1 compatible; `$ErrorActionPreference = 'Stop'`; all paths resolved from `$PSScriptRoot`):

1. Read `AccessVBOM` under `HKCU:\Software\Microsoft\Office\16.0\Excel\Security`. If it is not 1, print how to enable "Trust access to the VBA project object model" (File > Options > Trust Center > Trust Center Settings > Macro Settings) and exit 1.
2. Start Excel: `New-Object -ComObject Excel.Application`; `Visible = $false`; `DisplayAlerts = $false`. Everything after this is inside `try { } finally { }`. The `finally` closes the workbook without saving if it is still open, calls `Quit()`, releases every COM object held in a variable with `[Runtime.InteropServices.Marshal]::ReleaseComObject`, then `[GC]::Collect(); [GC]::WaitForPendingFinalizers()`.
3. `Workbooks.Add()`. Create sheets in this order: Settings, Staff, Holidays, LastYear, Leave, Roster, Review, Summary, Approvals, Test (rename the first existing sheet, add the rest after the last one, delete leftover default sheets at the end).
4. Write bold header rows (design section 2). Settings: `Key` / `Value` header, then the 12 keys with values from `fixtures/normal_year/settings.csv`.
5. Load `staff.csv`, `holidays.csv`, `lastyear.csv`, `leave.csv` from `fixtures/normal_year` with `Import-Csv`. Never open a CSV with Excel. Text columns: set `NumberFormat = '@'` before writing. Date columns: `[datetime]::ParseExact($t, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).ToOADate()` written to `.Value2`, with `NumberFormat = 'yyyy-mm-dd'`. Address cells with `$ws.Cells.Item($r, $c)`. Put the text `Sample dates - verify against MOM` in Holidays cell E1.
6. Test sheet: set the whole sheet `NumberFormat = '@'`, then write every line of `fixtures/testsheet.csv` as text (parse with `ConvertFrom-Csv -Header C1..Cn`, where n is the number of columns in the first line; all lines have the same width). Write one 2-D array per sheet in a single assignment.
7. Import every `src/vba/*.bas` with `$wb.VBProject.VBComponents.Import(<absolute path>)`.
8. Buttons with `$ws.Shapes.AddFormControl(0, left, top, width, height)`; set the caption with `.TextFrame.Characters().Text` and the macro with `.OnAction`. Settings: `Generate roster` → `modMain.GenerateRoster`, `Start new year` → `modMain.StartNewYear`, `Run self-test` → `modMain.RunSelfTest`. Holidays: `Expand long weekends` → `modMain.ExpandLongWeekends`. Review: `Apply approvals` → `modMain.ApplyApprovals`. Place buttons clear of the data (Settings column D; Holidays column E from row 3; Review row 1 to the right of column G is not available, so place it at column K).
9. Freeze the top row and autofit columns on each visible sheet. Hide Approvals and Test (`Visible = 0`).
10. Ensure `dist/` exists; delete an existing `dist/PHRosterPlanner.xlsm`; `$wb.SaveAs(<absolute path>, 52)`.
11. `$result = $xl.Run('modSelfTest.SelfTestSilent')`; print it; exit 0 when it starts with `PASS`, else 2. If `Run` itself throws (for example a VBA compile error), print the error and the hint `Open dist\PHRosterPlanner.xlsm, press Alt+F11, then Debug > Compile VBAProject to find the line`, and exit 3. The saved workbook is kept in all three cases.

`README.md` sections: what this is (two sentences); requirements (Windows, desktop Excel, the trust setting); build (`powershell -ExecutionPolicy Bypass -File build\build.ps1`) and what the three exit results mean; first run (enable macros; replace the sample Staff, Holidays, LastYear and Leave data; Expand long weekends; Generate roster; work through Review; Apply approvals); starting a new year; developing (`python -m pytest python/tests -q`, `python tools/check_vba.py src/vba`, `python python/roster/gen_fixtures.py`); known limits (design section 12).

- [ ] **Step 1:** Write `build/build.ps1`.
- [ ] **Step 2:** Check each numbered step above against the script; confirm every COM object stored in a variable is released in `finally`.
- [ ] **Step 3:** Write `README.md`.
- [ ] **Step 4:** Commit: `git add -A && git commit -m "Add Windows build script and README"`.

---

### Task 15: Parity guard and final review

**Files:**
- Create: `python/tests/test_parity.py`

**Interfaces — Consumes:** everything.

- [ ] **Step 1:** Write `test_parity.py`:
  - `check_vba.check_text` returns no findings for every file in `src/vba`.
  - Every constant in Global Constraints appears in `modTypes.bas` as `Public Const <NAME> As Long = <value>` with the Python value.
  - Every user-visible string in Global Constraints appears inside double quotes in some `.bas` file.
  - Every validation message's fixed text (the parts outside `<...>`) appears in `modValidate.bas`.
  - The five `OnAction` targets in `build.ps1` each match a `Public Sub <name>()` in `modMain.bas`; `SelfTestSilent` exists in `modSelfTest.bas`.
  - The set of sheet names created in `build.ps1` equals the set of names used as `Worksheets("<name>")` across the `.bas` files.
  - Every module listed in design section 10 exists.
- [ ] **Step 2:** Run `python -m pytest python/tests -q`; fix any mismatch in the VBA or the script (never by weakening the test). Expected: all pass.
- [ ] **Step 3:** Commit: `git add -A && git commit -m "Add Python/VBA parity guard"`.
- [ ] **Step 4 (coordinator):** Opus whole-repo review of the VBA against the Python, focused on: conditions that rely on short-circuiting, single-cell `Value2`, unsized arrays, undo on every search exit path, loop-order parity, and COM usage in `build.ps1`. Fix findings, rerun tests, push.

## After the plan

The user runs `build\build.ps1` on Windows and reports the printed result. Compile errors and self-test mismatches from that first run are fixed in a follow-up round; this is expected (design section 12).
