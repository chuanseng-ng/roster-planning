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
