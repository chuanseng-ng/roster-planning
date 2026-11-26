# PH Roster Planner (Excel VBA) — Design Spec

Nov 19, 2026 · @Chris

## Purpose and scope

An Excel workbook with a VBA macro generates the active and standby roster for every Singapore public-holiday date in a year. It flags every rule it cannot meet for the user to review instead of overriding anything.

- In scope: holiday dates entered by the user, long-weekend expansion, active and standby slots, exemption from last year, leave and shift limits, fairness balancing, gap report.
- Out of scope: normal (non-holiday) rostering, fetching holiday dates online, pay or claims.
- Platform: Windows desktop Excel, macro-enabled workbook (.xlsm).

## Shift structure and staffing per day

Every holiday date needs 7 people: 5 active and 2 standby, split into two coverage groups.

| Group | Active slots | Standby slot |
| --- | --- | --- |
| Day + evening | Day 1, Day 2, Evening | Standby D/E |
| Night | Night 1, Night 2 | Standby N |

- Each person fills at most one slot per date.
- Multi-day holidays (CNY, long weekends) are rostered date by date; the same person may work more than one date of the same holiday.

## Workbook sheets

Five input sheets feed one macro, which writes three output sheets.

| Sheet | Type | Columns |
| --- | --- | --- |
| Staff | Input | Name, Title, Senior (Y/N), Role (e.g. Special Coag), Cannot do (Day / Evening / Night, any combination) |
| Holidays | Input | Holiday name, Date. A button expands long weekends: a Fri or Mon holiday adds that Sat and Sun; a Sat or Sun holiday adds the other weekend day and the Monday. User can edit or delete added rows |
| LastYear | Input | Holiday name, Date, Staff name, Slot. Entered by hand for the first run; later filled from the previous Roster |
| Leave | Input | Staff name, From date, To date, Shift (All / Day / Evening / Night) |
| Settings | Input | Year, slot counts, rule switches |
| Roster | Output | One row per date: Holiday, Date, Day 1, Day 2, Evening, Night 1, Night 2, Standby D/E, Standby N |
| Review | Output | Date, Group, Rule not met, Reason, Candidates the user could approve |
| Summary | Output | Staff name, Active slots, Eligible dates, Ratio, Standby slots |

## Rules and constraints

Hard rules are never broken by the macro; when one cannot be met, the slot is left as GAP and reported on the Review sheet.

**Hard rules (per date)**

1. Senior: each group (day + evening, night) has at least one senior in an active slot.
2. Special coag: each group has a special coag staff in an active slot; if none, that group's standby must be special coag.
3. One slot per person per date.
4. Exemption: anyone active on any date of a holiday last year (matched by holiday name) cannot be active on any date of that holiday this year. They may still be standby. Standby last year does not count.
5. Leave: a person on leave for a date and shift is not placed in that shift or its standby. Shift "All" blocks the whole date.
6. Cannot do: a person is never placed in a shift listed in their Cannot do column, including that shift's standby (Night blocks Standby N; Day or Evening blocks Standby D/E).
7. Rest after night: anyone on an active night slot on one date can only be given a night slot (Night 1, Night 2 or Standby N) on the next calendar day, if it is also a holiday date. Standby N does not count as a night shift for this rule.

**Soft rule (fairness)**

- Ratio = active slots this year ÷ eligible dates this year. A date is eligible if the person is not exempt and not on full-day leave.
- When several people could fill a slot, the one with the lowest ratio so far goes first. Standby does not count toward the ratio.
- Fairness covers this year only.

## Algorithm

The macro fills one date at a time with a backtracking search, ordered so that fairness is decided by who gets picked first.

1. Validate inputs: unknown names, duplicate staff, dates outside the year, holiday names in LastYear that do not match Holidays. Stop with a message list if any are found.
2. Build the exemption set: for each holiday name, everyone active on it last year.
3. Sort dates chronologically.
4. For each date, fill the hardest slots first: the senior and special coag requirements in each group, then the remaining slots.
5. For each slot, list eligible candidates (rules 3–7) and sort them by lowest ratio, then fewest standby slots, then name.
6. Try candidates in order. If a later slot on the same date cannot be filled, undo and try the next candidate (backtrack).
7. If no full assignment exists, keep the best partial one, mark the unfilled slots GAP, and write the reasons to Review.
8. Update each person's ratio, then move to the next date.
9. Write Roster, Review and Summary.

Size check: 7 slots and a few dozen staff per date keeps the search to well under a second per date.

## Gap flagging and review

The macro never breaks a hard rule on its own; the user decides every exception on the Review sheet.

- Each Review row names the date, group, the rule not met and why (e.g. "No senior for night: Tan exempt, Lim on leave, Wong cannot do night").
- Candidates column lists people who would fit if one rule were relaxed, with the rule they would break (usually the exemption).
- The user types a name into an Approve column and runs "Apply approvals". The macro places that person, marks the cell as an approved exception, and re-runs later dates so ratios stay correct.
- GAP cells and approved exceptions are coloured on the Roster sheet.

## Testing

The workbook ships with a hidden Test sheet and a "Run self-test" macro that checks every rule on sample data.

- Rule checks: after any run, a checker re-reads Roster and confirms rules 1–7 for every filled slot.
- Scenarios: normal year; only senior exempt (expect GAP + candidate); no special coag active (expect coag standby); leave and Cannot do blocks; long weekend with 3 dates; CNY with 2 dates.
- Fairness check: spread between highest and lowest ratio, reported on Summary.

## Open points to confirm

None. All questions resolved on 19 Nov 2026.
