# PH Roster Planner — Implementation Design

Date: 2026-11-26. Source spec: `docs/source-spec.md` (dated 19 Nov 2026). This document adds the
decisions the source spec leaves open. Where the two differ, this document wins.

## 1. Delivery and verification

- The repo holds VBA source as text (`src/vba/*.bas`), a Python twin of the roster engine
  (`python/roster/`), shared scenario fixtures (`fixtures/`), and a Windows build script
  (`build/build.ps1`).
- `build.ps1` is run by the user on Windows. It drives Excel over COM to create the sheets,
  formatting and buttons, import the VBA modules, load sample data and the hidden Test sheet from
  `fixtures/`, and save `dist/PHRosterPlanner.xlsm`. It needs Excel's "Trust access to the VBA
  project object model" setting.
- The development container is Linux with no Excel. Python code and tests run there. VBA and
  `build.ps1` cannot be executed there; they are verified by code review and static checks, then
  by the user running the build and the "Run self-test" macro on Windows.
- The Python twin is the executable reference. It generates the expected results stored in
  `fixtures/`. The VBA self-test runs the same scenarios in memory and compares cell by cell.
- Parity rules for both engines: identical step order, identical candidate order, ratios compared
  by integer cross-multiplication (never floating point), names compared as upper-cased text with
  binary comparison, identical search-node counting.

## 2. Sheets

Input: Staff, Holidays, LastYear, Leave, Settings. Output: Roster, Review, Summary.
Hidden: Test (self-test scenarios and expectations), Approvals (applied exceptions).

| Sheet | Columns |
| --- | --- |
| Staff | Name, Title, Senior (Y/N), Role (comma-separated list), Cannot do (comma-separated: Day, Evening, Night) |
| Holidays | Holiday name, Date, Added (Y on rows created by the expand button) |
| LastYear | Holiday name, Date, Staff name, Slot |
| Leave | Staff name, From date, To date, Shift (All / Day / Evening / Night) |
| Settings | Year; counts for Day, Evening, Night, Standby D/E, Standby N; Y/N switch for rules 1, 2, 4, 5, 6, 7 |
| Roster | Holiday, Date, then one column per slot: Day 1..n, Evening 1..n, Night 1..n, Standby D/E 1..n, Standby N 1..n |
| Review | Date, Group, Slot, Rule not met, Reason, Candidates, Approve |
| Summary | Staff name, Active slots, Eligible dates, Ratio, Standby slots; plus one line with the ratio spread |
| Approvals | Date, Slot, Staff name, Rules broken |

- A slot column is numbered only when its count is above 1 (so the default layout is Day 1, Day 2,
  Evening, Night 1, Night 2, Standby D/E, Standby N, as in the source spec).
- Buttons: "Expand long weekends" on Holidays; "Generate roster", "Start new year" and
  "Run self-test" on Settings; "Apply approvals" on Review.
- Shipped input sheets are pre-filled with about 30 fictional staff, a 2027 Singapore holiday
  list labelled "sample — verify against MOM", a fictional LastYear and a few Leave rows.

## 3. Settings

- Year: whole number.
- Slot counts: each a whole number 0 to 9. Defaults 2, 1, 2, 1, 1.
- Rule switches: Y/N for rules 1, 2, 4, 5, 6, 7. Rule 3 is always on.
- A group with zero active slots skips rules 1 and 2.
- A switched-off rule is ignored everywhere, including eligibility counting (rule 4 off: nobody
  is exempt; rule 5 off: leave is ignored).

## 4. Rules, made precise

Groups: DE = Day, Evening and Standby D/E slots. N = Night and Standby N slots. Active slots are
Day, Evening and Night slots.

1. Senior: each group with active slots has at least one Senior = Y person in an active slot.
2. Special coag: a person is special coag when their Role list contains "Special Coag" (case and
   spaces ignored). Each group with active slots has a special coag person in an active slot, or
   failing that, in one of that group's standby slots. One person may satisfy rules 1 and 2.
   The engine prefers an active special coag and uses the standby form only when it must (see 5).
3. A person fills at most one slot per date.
4. Exemption: a person with an active slot in LastYear under a holiday name cannot take an active
   slot on any date with that holiday name this year. Standby is allowed. LastYear rows whose Slot
   starts with "Standby" do not count.
5. Leave: a Leave row covers a date when From ≤ date ≤ To. Shift All blocks every slot. Day blocks
   Day slots and Standby D/E. Evening blocks Evening slots and Standby D/E. Night blocks Night
   slots and Standby N.
6. Cannot do: same blocking pattern as rule 5, taken from the Staff sheet, on every date.
7. Rest after night: a person in an active Night slot on date D may, on D+1 (if D+1 is a roster
   date), only take a Night slot or Standby N. Standby N on D triggers nothing.

Fairness:

- Eligible date for a person: a roster date where the person is not exempt (rule 4) and has no
  Shift All leave covering it (rule 5).
- Ratio = active slots so far this year ÷ eligible dates for the whole year. The denominator is
  fixed before the first date is filled.
- Candidate order for every slot on a date: lowest ratio, then fewest standby slots so far, then
  name. People with zero eligible dates sort last. The order is computed once per date, from the
  state before that date.
- An approved exception in an active slot counts as an active slot.

## 5. Engine

Dates are filled in chronological order. Each date is solved by one depth-first search.

Steps for a date, in order:

1. Senior step for DE (if rule 1 is on and DE has active slots).
2. Special coag step for DE (if rule 2 is on and DE has active slots).
3. Senior step for N, then special coag step for N, under the same conditions.
4. One step per still-empty slot, in Roster column order.

Options tried at each step, in order:

- Senior step: skipped if a senior is already in an active slot of the group. Otherwise, for each
  senior in candidate order, for each empty active slot of the group in column order that the
  person is eligible for: place them. Last option, "unmet": the last empty active slot of the
  group becomes a reserved GAP.
- Special coag step: skipped if a special coag person is already in an active slot of the group.
  Otherwise: (a) each special coag person × each empty, unreserved active slot of the group;
  (b) "fallback": each special coag person eligible for the group's first standby slot, placed
  there; (c) "unmet": the group's first standby slot becomes a reserved GAP, or, if the group has
  no standby slots, its last empty active slot.
- Slot step: each eligible person in candidate order; last option: plain GAP.

Eligibility of a person for a slot applies rules 3 to 7, subject to the switches.

Score of an assignment: (GAP count, GAPs in active slots, fallback count), compared left to
right, lower is better. The search keeps the first assignment found with the best score. A branch
is cut as soon as its running score is not better than the best so far.

Stop conditions:

- Lower bound: one active GAP for each group with no senior individually eligible for any of its
  active slots; one fallback for each group with no special coag person individually eligible for
  any of its active slots (or one GAP if none is eligible for standby either). The search stops
  when the best score equals the lower bound. In a normal year this is the first leaf reached.
- Node cap: 200,000 placements per date. If reached, the best assignment so far is kept and the
  Review sheet says the search was cut short for that date.

After each date, active and standby counts are updated.

## 6. Validation

"Generate roster" stops with one message listing every problem found:

- Staff: blank or duplicate names, Senior not Y/N, unknown words in Cannot do.
- Holidays: blank names, invalid dates, dates outside the Settings year, duplicate dates.
- LastYear: unknown staff names, holiday names not present in Holidays, blank Slot.
- Leave: unknown staff names, invalid dates, From after To, unknown Shift.
- Settings: Year not a whole number, counts outside 0 to 9, switches not Y/N.

## 7. Review, approvals and re-runs

- One Review row per GAP slot. "Rule not met" is "No senior", "No special coag", or "No eligible
  staff". Reason lists the relevant people and what blocks each (exempt, on leave, cannot do,
  rest after night, already rostered that date).
- Candidates lists people who would fit the slot if exactly one of rules 4 to 7 were relaxed,
  with that rule named. For a reserved senior or special coag GAP, only seniors or special coag
  people are listed.
- "Apply approvals" handles Review rows that have a name in Approve, in date order:
  - The name must be on the Staff sheet and must not already hold a slot on that date (rule 3 is
    never broken). Otherwise the row is skipped and reported.
  - The macro lists the rules the person would break and asks for confirmation.
  - On confirmation the person is placed, the cell is coloured as an approved exception, a cell
    note lists the rules broken, and a row is added to Approvals.
  - That date's other slots and all earlier dates stay as they are. All later dates are
    regenerated. Approvals already recorded on later dates are kept: those people are placed
    first and the rest of the date is generated around them.
  - If a later pending row's slot is no longer a GAP after regeneration, it is skipped and
    reported.
- "Generate roster" when Approvals has rows: warns that approvals will be discarded and asks
  before continuing.
- GAP cells and approved exceptions have distinct fill colours on Roster.

## 8. Other macros

- Expand long weekends: for each Holidays row not marked Added: Friday adds the following Saturday
  and Sunday; Monday adds the preceding Saturday and Sunday; Saturday adds the Sunday and Monday
  after it; Sunday adds the Saturday before and the Monday after. Added rows take the same holiday
  name and Added = Y. A date already on the sheet is not added again. A date outside the Settings
  year is not added and is reported.
- Start new year: after confirmation, replaces LastYear with the Roster's active assignments
  (approved exceptions included; standby and GAPs excluded), clears Holidays, Leave, Roster,
  Review, Summary and Approvals, and adds 1 to Year.
- Rule checker: re-reads Roster and confirms rules 1 to 7 for every filled slot. Approved cells
  are excused only for the rules recorded in Approvals. Runs after every generate and apply; any
  finding is shown to the user.
- Run self-test: for each scenario on the Test sheet, builds the inputs in memory (user sheets are
  not touched), runs the engine, compares Roster and Summary with the stored expectation and the
  Review rows by date, slot and rule, runs the rule checker, and shows pass/fail per scenario.

## 9. Scenarios

Stored under `fixtures/<scenario>/` as CSV, one file per sheet, plus expected Roster, Review and
Summary generated by the Python twin.

- Normal year (the shipped sample data).
- Only senior exempt: expect a reserved GAP and the senior as candidate.
- No special coag available for active: expect special coag in standby.
- Leave and Cannot do blocks.
- Long weekend with 3 dates.
- CNY with 2 dates.
- Rest after night across consecutive dates.
- Changed slot counts and a switched-off rule.
- Approval applied, later dates regenerated.

The fairness spread (highest ratio minus lowest, among people with eligible dates) is reported on
Summary and checked in the normal-year scenario.

## 10. Code layout

| Path | Purpose |
| --- | --- |
| `src/vba/modTypes.bas` | Constants and data types |
| `src/vba/modRules.bas` | Eligibility for rules 3 to 7; rule checker |
| `src/vba/modEngine.bas` | Per-date search, year loop, fairness order |
| `src/vba/modReview.bas` | Review reasons and candidates |
| `src/vba/modIO.bas` | Read input sheets into arrays; write Roster, Review, Summary |
| `src/vba/modValidate.bas` | Input validation |
| `src/vba/modHolidays.bas` | Expand long weekends |
| `src/vba/modApprovals.bas` | Apply approvals and regenerate |
| `src/vba/modNewYear.bas` | Start new year |
| `src/vba/modSelfTest.bas` | Self-test |
| `src/vba/modMain.bas` | Button entry points |
| `python/roster/` | Twin of types, rules, engine, review, validation, holidays, approvals; CSV loader; fixture generator |
| `python/tests/` | pytest suite |
| `tools/check_vba.py` | Static checks on VBA sources (ASCII only, CRLF, Option Explicit, balanced blocks) |
| `build/build.ps1` | Builds the workbook on Windows |
| `fixtures/` | Scenario inputs and expectations |

VBA sources are stored with CRLF line endings and ASCII characters only so Excel imports them
unchanged.

## 11. Limits

- VBA and `build.ps1` are not executed before the user's first Windows build. Expect at least one
  round of fixes after that first run.
- The 2027 sample holiday dates are not checked against the official list.
