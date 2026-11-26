"""Static parity guard: the VBA sources and build/build.ps1 versus the Python engine."""
import importlib.util
import pathlib
import re

import pytest

from roster import model

REPO = pathlib.Path(__file__).resolve().parents[2]
VBA_DIR = REPO / "src" / "vba"
BUILD_PS1 = REPO / "build" / "build.ps1"
DESIGN = REPO / "docs" / "superpowers" / "specs" / "2026-11-26-ph-roster-planner-design.md"
TASK3_BRIEF = REPO / ".superpowers" / "sdd" / "task-3-brief.md"

spec = importlib.util.spec_from_file_location("check_vba", REPO / "tools" / "check_vba.py")
check_vba = importlib.util.module_from_spec(spec)
spec.loader.exec_module(check_vba)

BAS_FILES = sorted(VBA_DIR.glob("*.bas"))


def text_of(path):
    return path.read_bytes().decode("ascii")


def all_vba_text():
    return "\n".join(text_of(p) for p in BAS_FILES)


def vba_code_lines(path):
    """(line number, line) pairs with comments removed (quote-aware)."""
    out = []
    for n, line in enumerate(text_of(path).splitlines(), 1):
        in_q = False
        cut = len(line)
        for i, ch in enumerate(line):
            if ch == '"':
                in_q = not in_q
            elif ch == "'" and not in_q:
                cut = i
                break
        out.append((n, line[:cut]))
    return out


PROC_RE = re.compile(r"^\s*Public\s+(?:Sub|Function)\s+(\w+)", re.IGNORECASE)


def public_procs():
    """{module stem: [proc names]} for Public Sub/Function declarations."""
    result = {}
    for p in BAS_FILES:
        names = []
        for _, line in vba_code_lines(p):
            m = PROC_RE.match(line)
            if m:
                names.append(m.group(1))
        result[p.stem] = names
    return result


def test_bas_files_exist():
    assert BAS_FILES


@pytest.mark.parametrize("path", BAS_FILES, ids=lambda p: p.name)
def test_check_vba_clean(path):
    assert check_vba.check_text(path.stem, path.read_bytes()) == []


# ---------------------------------------------------------------- constants
VBA_KIND_NAMES = ["K_DAY", "K_EVE", "K_NIGHT", "K_SBY_DE", "K_SBY_N"]
PY_KINDS = [model.DAY, model.EVE, model.NIGHT, model.SBY_DE, model.SBY_N]
OTHER_CONSTS = ["M_DAY", "M_EVE", "M_NIGHT", "M_ALL", "GAP", "GAP_SENIOR",
                "GAP_COAG", "EMPTY_CELL", "NODE_CAP"]
CONST_CASES = ([(n, v) for n, v in zip(VBA_KIND_NAMES, PY_KINDS)]
               + [(n, getattr(model, n)) for n in OTHER_CONSTS])


@pytest.mark.parametrize("name,value", CONST_CASES, ids=[c[0] for c in CONST_CASES])
def test_constant_declared_in_modtypes(name, value):
    text = text_of(VBA_DIR / "modTypes.bas")
    pat = r"^\s*Public\s+Const\s+%s\s+As\s+Long\s*=\s*%s\s*(?:'.*)?$" % (
        re.escape(name), re.escape(str(value)))
    assert re.search(pat, text, re.MULTILINE), \
        "modTypes.bas lacks 'Public Const %s As Long = %s'" % (name, value)


# ---------------------------------------------------------------- strings
USER_STRINGS = (list(model.KIND_NAMES) + list(model.GROUP_NAMES)
                + list(model.RULE_TEXT.values())
                + [model.CUT_SHORT_TEXT, model.GAP_TEXT])


@pytest.mark.parametrize("s", USER_STRINGS)
def test_user_string_in_double_quotes(s):
    assert '"' + s + '"' in all_vba_text(), "no VBA string literal \"%s\"" % s


# ---------------------------------------------------------------- validation messages
def task3_message_formats():
    """The backticked message formats from the Task 3 brief's validation list."""
    text = TASK3_BRIEF.read_text(encoding="utf-8")
    start = text.index("Validation messages")
    end = text.index("A date is valid when")
    formats = []
    for line in text[start:end].splitlines():
        m = re.match(r"^- `(.+)`", line)
        if m:
            formats.append(m.group(1))
    return formats


def fixed_fragments(fmt):
    return [f for f in re.split(r"<[^>]*>", fmt) if f]


def test_task3_brief_lists_21_messages():
    assert len(task3_message_formats()) == 21


def test_validation_fragments_in_modvalidate():
    text = text_of(VBA_DIR / "modValidate.bas")
    missing = []
    for fmt in task3_message_formats():
        for frag in fixed_fragments(fmt):
            if '"' + frag + '"' not in text and frag not in text:
                missing.append((fmt, frag))
    assert missing == []


def test_validation_fragments_are_inside_string_literals():
    """Stricter form: each fragment sits inside a double-quoted literal."""
    text = text_of(VBA_DIR / "modValidate.bas")
    literals = re.findall(r'"((?:[^"]|"")*)"', text)
    joined = "\n".join(l.replace('""', '"') for l in literals)
    missing = []
    for fmt in task3_message_formats():
        for frag in fixed_fragments(fmt):
            if frag not in joined:
                missing.append((fmt, frag))
    assert missing == []


# ---------------------------------------------------------------- build.ps1
def button_macros():
    return re.findall(r"^\s*Add-Button\b.*?'(mod\w+\.\w+)'\s*$",
                      BUILD_PS1.read_text(encoding="utf-8"), re.MULTILINE)


def test_exactly_five_distinct_button_macros():
    macros = button_macros()
    assert len(macros) == 5
    assert len(set(macros)) == 5


def test_button_macros_are_public_subs_in_modmain():
    text = text_of(VBA_DIR / "modMain.bas")
    for macro in button_macros():
        module, proc = macro.split(".")
        assert module == "modMain", macro
        assert re.search(r"^\s*Public\s+Sub\s+%s\s*\(\s*\)" % re.escape(proc),
                         text, re.MULTILINE | re.IGNORECASE), macro


def test_selftestsilent_is_public_function():
    ps1 = BUILD_PS1.read_text(encoding="utf-8")
    m = re.search(r"\$xl\.Run\('(mod\w+)\.(\w+)'\)", ps1)
    assert m, "build.ps1 has no $xl.Run('<module>.<proc>') call"
    assert (m.group(1), m.group(2)) == ("modSelfTest", "SelfTestSilent")
    text = text_of(VBA_DIR / "modSelfTest.bas")
    assert re.search(r"^\s*Public\s+Function\s+SelfTestSilent\b", text,
                     re.MULTILINE | re.IGNORECASE)


def test_sheet_names_match_between_build_and_vba():
    ps1 = BUILD_PS1.read_text(encoding="utf-8")
    m = re.search(r"^\$sheetNames\s*=\s*@\((.*?)\)\s*$", ps1, re.MULTILINE)
    assert m, "build.ps1 has no $sheetNames = @(...) line"
    built = set(re.findall(r"'([^']+)'", m.group(1)))
    assert len(built) >= 2
    text = all_vba_text()
    used = set(re.findall(r'Worksheets\("([^"]+)"\)', text))
    # sheets reached through a helper taking the sheet name as its first argument
    used |= set(re.findall(r'^\s*(?:ReadTable|ClearDataRows)\s+"([^"]+)"', text, re.MULTILINE))
    used |= set(re.findall(r'\bSheetHasData\("([^"]+)"', text))
    assert built == used, "only in build.ps1: %s; only in VBA: %s" % (
        sorted(built - used), sorted(used - built))


# ---------------------------------------------------------------- module list
def design_modules():
    text = DESIGN.read_text(encoding="utf-8")
    section = text[text.index("## 10. Code layout"):]
    nxt = re.search(r"^## ", section[5:], re.MULTILINE)
    if nxt:
        section = section[:nxt.start() + 5]
    return re.findall(r"`src/vba/(mod\w+)\.bas`", section)


def test_design_modules_exist_and_no_extras():
    listed = design_modules()
    assert len(listed) == 11
    assert len(set(listed)) == len(listed)
    on_disk = {p.stem for p in BAS_FILES}
    assert set(listed) == on_disk, "missing: %s; unlisted: %s" % (
        sorted(set(listed) - on_disk), sorted(on_disk - set(listed)))


@pytest.mark.parametrize("path", BAS_FILES, ids=lambda p: p.name)
def test_first_line_is_vb_name(path):
    first = text_of(path).split("\r\n", 1)[0]
    assert first == 'Attribute VB_Name = "%s"' % path.stem


# ---------------------------------------------------------------- procedure names
def test_public_procedure_names_unique():
    seen = {}
    dups = []
    for module, names in public_procs().items():
        for n in names:
            key = n.lower()
            if key in seen:
                dups.append("%s (%s and %s)" % (n, seen[key], module))
            else:
                seen[key] = module
    assert dups == []


PUBLIC_VAR_RE = re.compile(r"^\s*Public\s+(?:Const\s+)?(\w+)\b", re.IGNORECASE)


def public_members():
    """Public procedures plus module-level Public variables and constants."""
    members = {m: {n.lower() for n in names} for m, names in public_procs().items()}
    for p in BAS_FILES:
        for _, line in vba_code_lines(p):
            m = PUBLIC_VAR_RE.match(line)
            if m and m.group(1).lower() not in ("sub", "function", "type", "enum"):
                members[p.stem].add(m.group(1).lower())
    return members


def test_qualified_calls_refer_to_public_members():
    """mod<Name>.<X> must name a Public procedure (or Public variable/constant) of that module."""
    procs = public_members()
    bad = []
    for p in BAS_FILES:
        for n, line in vba_code_lines(p):
            # drop string literals so text like "modMain.X" is not treated as a call
            code = re.sub(r'"(?:[^"]|"")*"', '""', line)
            for mod, proc in re.findall(r"\b(mod[A-Z]\w*)\.(\w+)", code):
                if mod not in procs:
                    bad.append("%s:%d: unknown module %s" % (p.name, n, mod))
                elif proc.lower() not in procs[mod]:
                    bad.append("%s:%d: %s.%s is not a Public member" % (p.name, n, mod, proc))
    assert bad == []
