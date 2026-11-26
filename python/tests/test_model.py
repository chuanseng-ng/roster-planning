from roster.model import build_slots, to_serial, from_serial, norm

def test_default_layout_names():
    names = [s.name for s in build_slots([2, 1, 2, 1, 1])]
    assert names == ["Day 1", "Day 2", "Evening", "Night 1", "Night 2", "Standby D/E", "Standby N"]

def test_zero_count_drops_column_and_single_is_unnumbered():
    slots = build_slots([1, 0, 3, 2, 0])
    assert [s.name for s in slots] == ["Day", "Night 1", "Night 2", "Night 3",
                                       "Standby D/E 1", "Standby D/E 2"]
    assert [s.group for s in slots] == [0, 1, 1, 1, 0, 0]
    assert [s.active for s in slots] == [True, True, True, True, False, False]
    assert [s.mask for s in slots] == [1, 4, 4, 4, 3, 3]

def test_serial_matches_excel():
    assert to_serial("2027-01-01") == 46388
    assert from_serial(46388) == "2027-01-01"
    assert to_serial("1900-03-01") == 61

def test_norm_trims_spaces_only_and_uppercases():
    assert norm("  Special Coag ") == "SPECIAL COAG"
    assert norm("\tx") == "\tX"
