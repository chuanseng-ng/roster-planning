from roster.model import (Settings, Inputs, build_slots, to_serial, norm, SETTINGS_KEYS,
                          M_DAY, M_EVE, M_NIGHT, M_ALL)
from roster.validate import find_setting

def piece_mask(word):
    w = norm(word)
    if w == "DAY":
        return M_DAY
    if w == "EVENING":
        return M_EVE
    if w == "NIGHT":
        return M_NIGHT
    if w == "ALL":
        return M_ALL
    return 0

def index_of(names, text):
    for i in range(len(names)):
        if names[i] == text:
            return i
    return -1

def build_inputs(raw):
    year = int(find_setting(raw, "Year").strip(' '))
    counts = []
    for key in SETTINGS_KEYS[1:6]:
        counts.append(int(find_setting(raw, key).strip(' ')))
    rules = {}
    for r in (1, 2, 4, 5, 6, 7):
        rules[r] = norm(find_setting(raw, "Rule " + str(r))) == "Y"
    settings = Settings(year, counts, rules)

    names = []
    keys = []
    senior = []
    coag = []
    cannot = []
    for row in raw.staff:
        names.append(row[0].strip(' '))
        keys.append(norm(row[0]))
        senior.append(norm(row[2]) == "Y")
        has_coag = False
        for piece in row[3].split(","):
            if norm(piece) == "SPECIAL COAG":
                has_coag = True
        coag.append(has_coag)
        mask = 0
        if rules[6]:
            for piece in row[4].split(","):
                w = norm(piece)
                if w == "DAY" or w == "EVENING" or w == "NIGHT":
                    mask |= piece_mask(w)
        cannot.append(mask)

    # holidays sorted ascending by serial (insertion sort, stable)
    pairs = []
    for row in raw.holidays:
        serial = to_serial(row[1].strip(' '))
        pos = len(pairs)
        while pos > 0 and pairs[pos - 1][0] > serial:
            pos -= 1
        pairs.insert(pos, (serial, row[0].strip(' ')))
    dates = []
    holiday = []
    for p in pairs:
        dates.append(p[0])
        holiday.append(p[1])

    exempt = []
    leave = []
    for s in range(len(names)):
        ex_row = []
        lv_row = []
        for d in range(len(dates)):
            ex = False
            if rules[4]:
                for ly in raw.lastyear:
                    if norm(ly[2]) == keys[s] and norm(ly[0]) == norm(holiday[d]):
                        if not norm(ly[3]).startswith("STANDBY"):
                            ex = True
            lv = 0
            if rules[5]:
                for lrow in raw.leave:
                    if norm(lrow[0]) == keys[s]:
                        if to_serial(lrow[1].strip(' ')) <= dates[d] <= to_serial(lrow[2].strip(' ')):
                            lv |= piece_mask(lrow[3])
            ex_row.append(ex)
            lv_row.append(lv)
        exempt.append(ex_row)
        leave.append(lv_row)

    return Inputs(settings, build_slots(counts), names, senior, coag, cannot,
                  dates, holiday, exempt, leave)
