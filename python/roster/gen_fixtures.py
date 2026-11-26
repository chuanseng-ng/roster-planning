"""Generate expected results and the Test-sheet table from the scenario folders in fixtures/.

Usage (from the repo root): python python/roster/gen_fixtures.py [out_root]
Inputs are always read from fixtures/; outputs go to out_root (default fixtures/).
"""
import csv
import os
import sys

_PYTHON_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PYTHON_DIR not in sys.path:
    sys.path.insert(0, _PYTHON_DIR)

from roster.csvio import (read_scenario, read_table, write_rows, STAFF_HEADER, HOLIDAYS_HEADER,
                          LASTYEAR_HEADER, LEAVE_HEADER, SETTINGS_HEADER)
from roster.engine import new_result, solve_year
from roster.inputs import build_inputs
from roster.model import GAP_TEXT, from_serial, to_serial, norm
from roster.ops import apply_approval
from roster.review import build_review, build_summary
from roster.validate import validate

REPO_ROOT = os.path.dirname(_PYTHON_DIR)
FIXTURES = os.path.join(REPO_ROOT, "fixtures")
APPROVE_HEADER = ["Date", "Slot", "Staff name"]
REVIEW_HEADER = ["Date", "Group", "Slot", "Rule not met", "Reason", "Candidates"]
SUMMARY_HEADER = ["Staff name", "Active slots", "Eligible dates", "Ratio", "Standby slots"]


class FixtureError(Exception):
    pass


def scenario_names(root=FIXTURES):
    out = []
    for name in sorted(os.listdir(root)):
        if os.path.isdir(os.path.join(root, name)):
            out.append(name)
    return out


def read_approvals(folder):
    return read_table(os.path.join(folder, "approve.csv"), 3, False)


def _index(items, text, what):
    for i in range(len(items)):
        if norm(items[i]) == norm(text):
            return i
    raise FixtureError("approval: unknown " + what + " " + text)


def solve_scenario(folder):
    """Returns (raw, inp, res, approval_rows, approvals). Raises FixtureError on bad input."""
    raw = read_scenario(folder)
    msgs = validate(raw)
    if len(msgs) > 0:
        raise FixtureError(folder + ": " + "; ".join(msgs))
    inp = build_inputs(raw)
    res = new_result(inp)
    solve_year(inp, res, 0)
    approval_rows = read_approvals(folder)
    approvals = []
    for row in approval_rows:
        serial = to_serial(row[0].strip(' '))
        d = -1
        for i in range(len(inp.dates)):
            if inp.dates[i] == serial:
                d = i
        if d < 0:
            raise FixtureError("approval: date " + row[0] + " is not a holiday date")
        slot_names = []
        for sl in inp.slots:
            slot_names.append(sl.name)
        k = _index(slot_names, row[1], "slot")
        s = _index(inp.staff, row[2], "staff name")
        ok, msg = apply_approval(inp, res, approvals, d, k, s)
        if not ok:
            raise FixtureError("approval " + row[0] + " " + row[1] + " " + row[2] + ": " + msg)
    return raw, inp, res, approval_rows, approvals


def roster_rows(inp, res):
    header = ["Holiday", "Date"]
    for sl in inp.slots:
        header.append(sl.name)
    rows = []
    for d in range(len(inp.dates)):
        row = [inp.holiday[d], from_serial(inp.dates[d])]
        for k in range(len(inp.slots)):
            v = res.grid[d][k]
            if v >= 0:
                row.append(inp.staff[v])
            else:
                row.append(GAP_TEXT)
        rows.append(row)
    return header, rows


def review_rows(inp, res):
    rows = []
    for r in build_review(inp, res):
        rows.append([from_serial(r.date), r.group, r.slot, r.rule, r.reason, r.candidates])
    return rows


def summary_rows(inp, res):
    rows, spread = build_summary(inp, res)
    rows.append(["Spread", "", "", spread, ""])
    return rows


def generate(out_root=FIXTURES, fixtures_root=FIXTURES):
    """Write every expected file and testsheet.csv under out_root. Returns {scenario: res}."""
    lines = []
    results = {}
    for name in scenario_names(fixtures_root):
        folder = os.path.join(fixtures_root, name)
        raw, inp, res, approval_rows, approvals = solve_scenario(folder)
        results[name] = (inp, res)
        r_header, r_rows = roster_rows(inp, res)
        rv_rows = review_rows(inp, res)
        s_rows = summary_rows(inp, res)
        target = os.path.join(out_root, name)
        os.makedirs(target, exist_ok=True)
        write_rows(os.path.join(target, "expected_roster.csv"), r_header, r_rows)
        write_rows(os.path.join(target, "expected_review.csv"), REVIEW_HEADER, rv_rows)
        write_rows(os.path.join(target, "expected_summary.csv"), SUMMARY_HEADER, s_rows)
        tables = [("staff", STAFF_HEADER, raw.staff), ("holidays", HOLIDAYS_HEADER, raw.holidays),
                  ("lastyear", LASTYEAR_HEADER, raw.lastyear), ("leave", LEAVE_HEADER, raw.leave),
                  ("settings", SETTINGS_HEADER, raw.settings)]
        if os.path.exists(os.path.join(folder, "approve.csv")):
            tables.append(("approve", APPROVE_HEADER, approval_rows))
        tables.append(("expected_roster", r_header, r_rows))
        tables.append(("expected_review", REVIEW_HEADER, rv_rows))
        tables.append(("expected_summary", SUMMARY_HEADER, s_rows))
        for table, header, rows in tables:
            lines.append([name, table, "0"] + list(header))
            n = 0
            for row in rows:
                n += 1
                lines.append([name, table, str(n)] + list(row))
    width = 0
    for line in lines:
        if len(line) > width:
            width = len(line)
    os.makedirs(out_root, exist_ok=True)
    with open(os.path.join(out_root, "testsheet.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        for line in lines:
            w.writerow(line + [""] * (width - len(line)))
    return results


def main(argv):
    out_root = FIXTURES
    if len(argv) > 1:
        out_root = argv[1]
    try:
        generate(out_root)
    except FixtureError as e:
        print(str(e))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
