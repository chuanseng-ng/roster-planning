import datetime

from roster.engine import solve_year
from roster.model import GAP_SENIOR, GAP_COAG, from_serial, norm
from roster.review import Approval
from roster.rules import blocking_rules, forward_night_block

# Day offsets to add, by isoweekday: Mon=1 ... Sun=7.
_OFFSETS = {1: (-2, -1), 5: (1, 2), 6: (1, 2), 7: (-1, 1)}


def _contains(items, x):
    for v in items:
        if v == x:
            return True
    return False


def expand_weekends(rows, year):
    out = []
    for r in rows:
        out.append(list(r))
    present = []
    for r in rows:
        present.append(r[1])
    skipped = []
    for r in rows:
        if norm(r[2]) == "Y":
            continue
        base = datetime.date(int(r[1][0:4]), int(r[1][5:7]), int(r[1][8:10]))
        wd = base.isoweekday()
        if wd not in _OFFSETS:
            continue
        for off in _OFFSETS[wd]:
            cand = base + datetime.timedelta(days=off)
            text = cand.isoformat()
            if _contains(present, text) or _contains(skipped, text):
                continue
            if cand.year != year:
                skipped.append(text)
                continue
            present.append(text)
            out.append([r[0], text, "Y"])
    # stable insertion sort by date text (yyyy-mm-dd sorts chronologically)
    for i in range(1, len(out)):
        row = out[i]
        j = i - 1
        while j >= 0 and out[j][1] > row[1]:
            out[j + 1] = out[j]
            j -= 1
        out[j + 1] = row
    return out, skipped


def roster_to_lastyear(inp, res):
    rows = []
    for d in range(len(inp.dates)):
        for k in range(len(inp.slots)):
            v = res.grid[d][k]
            if v >= 0 and inp.slots[k].active:
                rows.append([inp.holiday[d], from_serial(inp.dates[d]), inp.staff[v],
                             inp.slots[k].name])
    return rows


def _add_rule(rules, r):
    """Insert r into the ascending list rules unless already present."""
    if _contains(rules, r):
        return
    rules.append(r)
    i = len(rules) - 1
    while i > 0 and rules[i - 1] > rules[i]:
        t = rules[i - 1]
        rules[i - 1] = rules[i]
        rules[i] = t
        i -= 1


def approval_broken_rules(inp, res, d, k, s):
    out = []
    for r in blocking_rules(inp, res.grid, s, d, k):
        if r != 3:
            _add_rule(out, r)
    if forward_night_block(inp, res.grid, res.approved, s, d, k):
        _add_rule(out, 7)
    if res.grid[d][k] == GAP_SENIOR and not inp.senior[s]:
        _add_rule(out, 1)
    if res.grid[d][k] == GAP_COAG and not inp.coag[s]:
        _add_rule(out, 2)
    return out


def apply_approval(inp, res, approvals, d, k, s):
    if res.grid[d][k] >= 0:
        return False, "slot is no longer a GAP"
    for j in range(len(inp.slots)):
        if res.grid[d][j] == s:
            return False, inp.staff[s] + " already has a slot on this date"
    approvals.append(Approval(d, k, s, approval_broken_rules(inp, res, d, k, s)))
    res.grid[d][k] = s
    res.approved[d][k] = True
    solve_year(inp, res, d + 1)
    notes = []
    for a in approvals:
        if a.d > d:
            extra = []
            for r in blocking_rules(inp, res.grid, a.s, a.d, a.k):
                if r != 3 and not _contains(a.rules, r):
                    extra.append(r)
            if len(extra) > 0:
                for r in extra:
                    _add_rule(a.rules, r)
                text = ""
                for i in range(len(extra)):
                    if i > 0:
                        text += ", "
                    text += str(extra[i])
                notes.append("approval for " + inp.staff[a.s] + " on "
                             + from_serial(inp.dates[a.d]) + " now also breaks rule " + text)
    msg = ""
    for i in range(len(notes)):
        if i > 0:
            msg += "; "
        msg += notes[i]
    return True, msg
