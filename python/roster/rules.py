from roster.model import NIGHT, SBY_N, M_ALL


def blocking_rules(inp, grid, s, d, k):
    """Rules among 3,4,5,6,7 that stop staff s taking slot k on date index d, ascending.
    The cell (d, k) itself is ignored."""
    out = []
    slot = inp.slots[k]
    # Rule 3: already in another cell of this date.
    for j in range(len(inp.slots)):
        if j != k and grid[d][j] == s:
            out.append(3)
            break
    # Rule 4: exempt, for active slots only.
    if slot.active and inp.exempt[s][d]:
        out.append(4)
    # Rule 5: leave.
    lv = inp.leave[s][d]
    if (lv & M_ALL) != 0 or (lv & slot.mask) != 0:
        out.append(5)
    # Rule 6: cannot-do flags.
    if (inp.cannot[s] & slot.mask) != 0:
        out.append(6)
    # Rule 7: no day/evening work straight after a night.
    if inp.settings.rules[7] and d > 0:
        if inp.dates[d - 1] == inp.dates[d] - 1:
            if slot.kind != NIGHT and slot.kind != SBY_N:
                for j in range(len(inp.slots)):
                    if inp.slots[j].kind == NIGHT and grid[d - 1][j] == s:
                        out.append(7)
                        break
    return out


def forward_night_block(inp, grid, approved, s, d, k):
    """True when staff s on a night slot k of date d would collide with an approved
    day/evening/standby D/E cell of s on the next calendar day (rule 7)."""
    if not inp.settings.rules[7]:
        return False
    if inp.slots[k].kind != NIGHT:
        return False
    if d + 1 >= len(inp.dates):
        return False
    if inp.dates[d + 1] != inp.dates[d] + 1:
        return False
    for j in range(len(inp.slots)):
        kind = inp.slots[j].kind
        if kind != NIGHT and kind != SBY_N:
            if approved[d + 1][j] and grid[d + 1][j] == s:
                return True
    return False
