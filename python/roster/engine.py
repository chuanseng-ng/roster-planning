from roster import model
from roster.model import GAP, GAP_SENIOR, GAP_COAG, EMPTY_CELL
from roster.fairness import eligible_dates, candidate_order
from roster.rules import blocking_rules, forward_night_block

# NODE_CAP is read at call time as model.NODE_CAP (not imported by name) so tests can lower it.


def new_result(inp):
    nd = len(inp.dates)
    nk = len(inp.slots)
    ns = len(inp.staff)
    grid = []
    approved = []
    for d in range(nd):
        grid.append([EMPTY_CELL] * nk)
        approved.append([False] * nk)
    return model.YearResult(grid, approved, [0] * ns, [0] * ns, eligible_dates(inp),
                            [False] * nd, [[] for _ in range(nd)], [0] * nd)


def max_matching(n_left, adj):
    """Kuhn's algorithm. adj[i] = right-side ints for left node i."""
    n_right = 0
    for i in range(n_left):
        for r in adj[i]:
            if r + 1 > n_right:
                n_right = r + 1
    match_right = [-1] * n_right

    def augment(u, seen):
        for r in adj[u]:
            if not seen[r]:
                seen[r] = True
                if match_right[r] < 0 or augment(match_right[r], seen):
                    match_right[r] = u
                    return True
        return False

    total = 0
    for u in range(n_left):
        if augment(u, [False] * n_right):
            total += 1
    return total


def _add(a, b):
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


class _Search:
    """Search state for one date. Mirrors module-level variables in the VBA port."""

    def __init__(self, inp, res, d):
        self.inp = inp
        self.res = res
        self.d = d
        nk = len(inp.slots)
        ns = len(inp.staff)
        self.nk = nk
        self.row = res.grid[d]
        self.order = candidate_order(inp, res.active_count, res.standby_count, res.eligible)
        self.used = [False] * ns
        for k in range(nk):
            if self.row[k] >= 0:
                self.used[self.row[k]] = True
        self.elig = []
        for s in range(ns):
            e = []
            for k in range(nk):
                rules = [r for r in blocking_rules(inp, res.grid, s, d, k) if r != 3]
                e.append(len(rules) == 0 and not forward_night_block(inp, res.grid, res.approved, s, d, k))
            self.elig.append(e)
        self.steps = []
        for g in (0, 1):
            if len(self.active_slots(g)) > 0:
                if inp.settings.rules[1]:
                    self.steps.append(("senior", g))
                if inp.settings.rules[2]:
                    self.steps.append(("coag", g))
        for k in range(nk):
            if self.row[k] == EMPTY_CELL:
                self.steps.append(("slot", k))
        self.bound = self.lower_bound()
        self.score = (0, 0, 0)
        self.best = None
        self.best_unmet = None
        self.best_score = None
        self.nodes = 0
        self.stop = False
        self.cut = False
        self.unmet = []

    # --- slot helpers -------------------------------------------------------
    def active_slots(self, g):
        return [k for k in range(self.nk) if self.inp.slots[k].group == g and self.inp.slots[k].active]

    def standby_slots(self, g):
        return [k for k in range(self.nk) if self.inp.slots[k].group == g and not self.inp.slots[k].active]

    def has_senior(self, g):
        for k in self.active_slots(g):
            if self.row[k] >= 0 and self.inp.senior[self.row[k]]:
                return True
        return False

    def has_coag(self, g):
        for k in self.active_slots(g):
            if self.row[k] >= 0 and self.inp.coag[self.row[k]]:
                return True
        for k in self.standby_slots(g):
            if self.row[k] >= 0 and self.res.approved[self.d][k] and self.inp.coag[self.row[k]]:
                return True
        return False

    def free(self, s, k):
        return (not self.used[s]) and self.elig[s][k]

    # --- lower bound --------------------------------------------------------
    def lower_bound(self):
        inp = self.inp
        order = self.order
        empty = [k for k in range(self.nk) if self.row[k] == EMPTY_CELL]
        adj = [[s for s in order if self.free(s, k)] for k in empty]
        gaps_a = len(empty) - max_matching(len(empty), adj)
        empty_a = [k for k in empty if inp.slots[k].active]
        adj = [[s for s in order if self.free(s, k)] for k in empty_a]
        active_a = len(empty_a) - max_matching(len(empty_a), adj)
        req_g = 0
        req_a = 0
        pen = 0
        for kind, g in self.steps:
            if kind == "senior":
                if not self.has_senior(g):
                    found = False
                    for s in order:
                        if inp.senior[s]:
                            for k in self.active_slots(g):
                                if self.row[k] == EMPTY_CELL and self.free(s, k):
                                    found = True
                    if not found:
                        req_g += 1
                        req_a += 1
                        pen += 2
            elif kind == "coag":
                if not self.has_coag(g):
                    found = False
                    for s in order:
                        if inp.coag[s]:
                            for k in self.active_slots(g):
                                if self.row[k] == EMPTY_CELL and self.free(s, k):
                                    found = True
                    if not found:
                        k1 = None
                        for k in self.standby_slots(g):
                            if self.row[k] == EMPTY_CELL:
                                k1 = k
                                break
                        ok = False
                        if k1 is not None:
                            for s in order:
                                if inp.coag[s] and self.free(s, k1):
                                    ok = True
                        if ok:
                            pen += 1
                        else:
                            pen += 2
                            req_g += 1
                            if k1 is None:
                                req_a += 1
        return (max(gaps_a, req_g), max(active_a, req_a), pen)

    # --- search -------------------------------------------------------------
    def rec(self, i):
        inp = self.inp
        row = self.row
        if i == len(self.steps):
            self.best = list(row)
            self.best_unmet = list(self.unmet)
            self.best_score = self.score
            if self.best_score == self.bound:
                self.stop = True
            return
        kind, x = self.steps[i]
        if kind == "slot":
            if row[x] != EMPTY_CELL:
                self.rec(i + 1)
                return
            for s in self.order:
                if not self.used[s] and self.elig[s][x]:
                    self.try_place(i, s, x, (0, 0, 0))
                    if self.stop:
                        return
            self.try_gap(i, x, GAP, None)
            return
        if kind == "senior":
            g = x
            if self.has_senior(g):
                self.rec(i + 1)
                return
            for s in self.order:
                if inp.senior[s] and not self.used[s]:
                    for k in self.active_slots(g):
                        if row[k] == EMPTY_CELL and self.elig[s][k]:
                            self.try_place(i, s, k, (0, 0, 0))
                            if self.stop:
                                return
            k = None
            for kk in self.active_slots(g):
                if row[kk] == EMPTY_CELL:
                    k = kk
            self.try_gap(i, k, GAP_SENIOR, g)
            return
        # kind == "coag"
        g = x
        if self.has_coag(g):
            self.rec(i + 1)
            return
        for s in self.order:                                   # (a) active
            if inp.coag[s] and not self.used[s]:
                for k in self.active_slots(g):
                    if row[k] == EMPTY_CELL and self.elig[s][k]:
                        self.try_place(i, s, k, (0, 0, 0))
                        if self.stop:
                            return
        k = None
        for kk in self.standby_slots(g):
            if row[kk] == EMPTY_CELL:
                k = kk
                break
        if k is not None:                                      # (b) fallback
            for s in self.order:
                if inp.coag[s] and not self.used[s] and self.elig[s][k]:
                    self.try_place(i, s, k, (0, 0, 1))
                    if self.stop:
                        return
        if k is None:                                          # (c) unmet
            for kk in self.active_slots(g):
                if row[kk] == EMPTY_CELL:
                    k = kk
        self.try_gap(i, k, GAP_COAG, g)

    def try_place(self, i, s, k, delta):
        new = _add(self.score, delta)
        if self.best_score is not None and new >= self.best_score:
            return
        self.nodes += 1
        # NODE_CAP must exceed the number of steps for a date (at most 49), so the first leaf is
        # always reached and `best` is never None at the real cap.
        if self.nodes > model.NODE_CAP:
            self.stop = True
            self.cut = True
            return
        self.row[k] = s
        self.used[s] = True
        old = self.score
        self.score = new
        self.rec(i + 1)
        self.row[k] = EMPTY_CELL
        self.used[s] = False
        self.score = old

    def try_gap(self, i, k, code, g):
        # Score = (gaps, active gaps, penalty); penalty is 2 per unmet requirement, 1 per fallback.
        pen = 2 if code != GAP else 0
        if k is None:
            delta = (1, 1, pen)
        elif self.inp.slots[k].active:
            delta = (1, 1, pen)
        else:
            delta = (1, 0, pen)
        new = _add(self.score, delta)
        if self.best_score is not None and new >= self.best_score:
            return
        self.nodes += 1
        if self.nodes > model.NODE_CAP:
            self.stop = True
            self.cut = True
            return
        if k is None:
            self.unmet.append((g, code))
        else:
            self.row[k] = code
        old = self.score
        self.score = new
        self.rec(i + 1)
        if k is None:
            self.unmet.pop()
        else:
            self.row[k] = EMPTY_CELL
        self.score = old


def solve_date(inp, res, d):
    st = _Search(inp, res, d)
    st.rec(0)
    res.grid[d] = st.best
    res.unmet[d] = st.best_unmet
    res.cut_short[d] = st.cut
    res.nodes[d] = st.nodes


def _count_date(inp, res, d):
    for k in range(len(inp.slots)):
        v = res.grid[d][k]
        if v >= 0:
            if inp.slots[k].active:
                res.active_count[v] += 1
            else:
                res.standby_count[v] += 1


def solve_year(inp, res, from_d):
    nd = len(inp.dates)
    nk = len(inp.slots)
    for d in range(from_d, nd):
        for k in range(nk):
            if not res.approved[d][k]:
                res.grid[d][k] = EMPTY_CELL
        res.cut_short[d] = False
        res.unmet[d] = []
        res.nodes[d] = 0
    for s in range(len(inp.staff)):
        res.active_count[s] = 0
        res.standby_count[s] = 0
    for d in range(from_d):
        _count_date(inp, res, d)
    for d in range(from_d, nd):
        solve_date(inp, res, d)
        _count_date(inp, res, d)
