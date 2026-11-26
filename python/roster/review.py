from dataclasses import dataclass

from roster.model import (GAP, GAP_SENIOR, GAP_COAG, GROUP_NAMES, KIND_NAMES, RULE_TEXT,
                          CUT_SHORT_TEXT, from_serial)
from roster.rules import blocking_rules


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


def _join(parts):
    out = ""
    for i in range(len(parts)):
        if i > 0:
            out += "; "
        out += parts[i]
    return out


def _why(rule, slot):
    if rule == 3:
        return "already rostered"
    if rule == 4:
        return "exempt"
    if rule == 5:
        return "on leave"
    if rule == 6:
        return "cannot do " + KIND_NAMES[slot.kind]
    return "rest after night"


def _cand_text(rule):
    if rule == 4:
        return "rule 4 exemption"
    if rule == 5:
        return "rule 5 leave"
    if rule == 6:
        return "rule 6 cannot do"
    return "rule 7 rest after night"


def _gap_row(inp, res, d, k):
    code = res.grid[d][k]
    slot = inp.slots[k]
    reasons = []
    cands = []
    for s in range(len(inp.staff)):
        considered = True
        if code == GAP_SENIOR:
            considered = inp.senior[s]
        elif code == GAP_COAG:
            considered = inp.coag[s]
        if not considered:
            continue
        rules = blocking_rules(inp, res.grid, s, d, k)
        if len(rules) == 0:
            reasons.append(inp.staff[s] + ": not placed")
        else:
            reasons.append(inp.staff[s] + ": " + _why(rules[0], slot))
        placed = False
        for j in range(len(inp.slots)):
            if res.grid[d][j] == s:
                placed = True
        if not placed and len(rules) == 1 and rules[0] >= 4 and rules[0] <= 7:
            cands.append(inp.staff[s] + " (" + _cand_text(rules[0]) + ")")
    if len(reasons) == 0:
        reason = "no staff with this requirement"
    else:
        reason = _join(reasons)
    return ReviewRow(inp.dates[d], GROUP_NAMES[slot.group], slot.name, RULE_TEXT[code],
                     reason, _join(cands))


def build_review(inp, res):
    rows = []
    for d in range(len(inp.dates)):
        if res.cut_short[d]:
            rows.append(ReviewRow(inp.dates[d], "", "", CUT_SHORT_TEXT,
                                  "best roster found so far was kept", ""))
        for k in range(len(inp.slots)):
            v = res.grid[d][k]
            if v == GAP or v == GAP_SENIOR or v == GAP_COAG:
                rows.append(_gap_row(inp, res, d, k))
        for i in range(len(res.unmet[d])):
            g = res.unmet[d][i][0]
            code = res.unmet[d][i][1]
            rows.append(ReviewRow(inp.dates[d], GROUP_NAMES[g], "", RULE_TEXT[code],
                                  "no free slot in this group", ""))
    return rows


def _approved(approvals, d, k, rule):
    """Is there an approval for exactly cell (d, k) listing the rule."""
    for a in approvals:
        if a.d == d and a.k == k:
            for r in a.rules:
                if r == rule:
                    return True
    return False


def _group_approved(inp, approvals, d, g, rule):
    """Is there an approval for any cell of group g on date index d listing the rule."""
    for a in approvals:
        if a.d == d and inp.slots[a.k].group == g:
            for r in a.rules:
                if r == rule:
                    return True
    return False


def check_rules(inp, res, approvals):
    out = []
    nk = len(inp.slots)
    for d in range(len(inp.dates)):
        for k in range(nk):
            s = res.grid[d][k]
            if s >= 0:
                rules = blocking_rules(inp, res.grid, s, d, k)
                for r in rules:
                    if not _approved(approvals, d, k, r):
                        out.append(from_serial(inp.dates[d]) + " " + inp.slots[k].name + ": "
                                   + inp.staff[s] + " breaks rule " + str(r))
    for d in range(len(inp.dates)):
        for g in (0, 1):
            n_active = 0
            negative = False
            has_senior = False
            has_coag = False
            for k in range(nk):
                if inp.slots[k].group != g:
                    continue
                v = res.grid[d][k]
                if v < 0:
                    negative = True
                    continue
                if inp.slots[k].active:
                    if inp.senior[v]:
                        has_senior = True
                if inp.coag[v]:
                    has_coag = True
            for k in range(nk):
                if inp.slots[k].group == g and inp.slots[k].active:
                    n_active += 1
            if n_active == 0 or negative:
                continue
            unmet = False
            for i in range(len(res.unmet[d])):
                if res.unmet[d][i][0] == g:
                    unmet = True
            if unmet:
                continue
            label = from_serial(inp.dates[d]) + " " + GROUP_NAMES[g] + ": "
            if inp.settings.rules[1] and not has_senior:
                if not _group_approved(inp, approvals, d, g, 1):
                    out.append(label + "rule 1 not met")
            if inp.settings.rules[2] and not has_coag:
                if not _group_approved(inp, approvals, d, g, 2):
                    out.append(label + "rule 2 not met")
    return out


def _q(a, e):
    return (a * 20000 + e) // (2 * e)


def _q_text(q):
    return str(q // 10000) + "." + ("%04d" % (q % 10000))


def ratio_text(a, e):
    if e == 0:
        return ""
    return _q_text(_q(a, e))


def build_summary(inp, res):
    rows = []
    have = False
    qmax = 0
    qmin = 0
    for s in range(len(inp.staff)):
        a = res.active_count[s]
        e = res.eligible[s]
        rows.append([inp.staff[s], str(a), str(e), ratio_text(a, e), str(res.standby_count[s])])
        if e > 0:
            q = _q(a, e)
            if not have:
                qmax = q
                qmin = q
                have = True
            else:
                if q > qmax:
                    qmax = q
                if q < qmin:
                    qmin = q
    if not have:
        return rows, "0.0000"
    return rows, _q_text(qmax - qmin)
