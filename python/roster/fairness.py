from roster.model import M_ALL, norm


def eligible_dates(inp):
    """Per staff: number of dates that are not exempt and not full-day leave."""
    out = []
    for s in range(len(inp.staff)):
        n = 0
        for d in range(len(inp.dates)):
            if not inp.exempt[s][d] and (inp.leave[s][d] & M_ALL) == 0:
                n += 1
        out.append(n)
    return out


def ratio_less(a1, e1, a2, e2):
    """Is a1/e1 strictly lower than a2/e2; requires e1, e2 > 0."""
    return a1 * e2 < a2 * e1


def _before(inp, active_count, standby_count, eligible, x, y):
    """True when staff x must be tried strictly before staff y."""
    zx = eligible[x] == 0
    zy = eligible[y] == 0
    if zx != zy:
        return zy
    if not zx:
        if ratio_less(active_count[x], eligible[x], active_count[y], eligible[y]):
            return True
        if ratio_less(active_count[y], eligible[y], active_count[x], eligible[x]):
            return False
    if standby_count[x] != standby_count[y]:
        return standby_count[x] < standby_count[y]
    nx = norm(inp.staff[x])
    ny = norm(inp.staff[y])
    if nx != ny:
        return nx < ny
    return False


def candidate_order(inp, active_count, standby_count, eligible):
    """Staff indexes in the order they are tried: staff with no eligible dates last, then
    lowest active/eligible ratio, then fewest standby, then name. Stable insertion sort."""
    order = []
    for s in range(len(inp.staff)):
        order.append(s)
    for i in range(1, len(order)):
        cur = order[i]
        j = i - 1
        while j >= 0 and _before(inp, active_count, standby_count, eligible, cur, order[j]):
            order[j + 1] = order[j]
            j -= 1
        order[j + 1] = cur
    return order
