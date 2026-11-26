"""Static checker for the VBA sources (no VBA compiler is available here).

Library: check_text(name, data) -> list of "<name>:<line>: <text>".
CLI:     python tools/check_vba.py <dir>   checks every *.bas, exit 1 if any problem.
"""
import pathlib
import re
import sys

RESERVED = {"date", "day", "year", "name", "weekday", "month"}

DECL_RE = re.compile(
    r"^(?:Dim|Private|Public|Static|Global)\s+"
    r"(?!(?:Sub|Function|Const|Property|Enum|Type|Declare|Event)\b)(.*)$", re.I)
CONST_RE = re.compile(r"^(?:(?:Private|Public|Global)\s+)?Const\s+(.*)$", re.I)
PROC_RE = re.compile(
    r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function)\s+\w+\s*\(", re.I)
AS_RE = re.compile(r"\sAs\s", re.I)
PARAM_OK_RE = re.compile(r"^(?:ByVal|ByRef|Optional\s+ByVal|Optional\s+ByRef)\s", re.I)
MODIFIER_RE = re.compile(r"^(?:Optional\s+|ByVal\s+|ByRef\s+)*", re.I)


def proc_params(line):
    """Return the parameter text of a Sub/Function header (bracket-matched), else None."""
    m = PROC_RE.match(line)
    if not m:
        return None
    depth = 1
    for i in range(m.end(), len(line)):
        if line[i] == "(":
            depth += 1
        elif line[i] == ")":
            depth -= 1
            if depth == 0:
                return line[m.end():i]
    return None


def strip_line(line):
    """Drop an unquoted comment and replace string literals with ""."""
    out = []
    i, n = 0, len(line)
    while i < n:
        c = line[i]
        if c == '"':
            i += 1
            while i < n:
                if line[i] == '"':
                    if i + 1 < n and line[i + 1] == '"':
                        i += 2
                        continue
                    break
                i += 1
            out.append('""')
            i += 1
        elif c == "'":
            break
        else:
            out.append(c)
            i += 1
    return "".join(out).rstrip()


def logical_lines(stripped):
    """Join ' _' continuations. Returns [(first_line_number, text)]."""
    result = []
    buf, start = None, 0
    for no, text in enumerate(stripped, 1):
        t = text.strip()
        if buf is None:
            buf, start = "", no
        if t.endswith(" _") or t == "_":
            buf += t[:-1].rstrip() + " "
        else:
            result.append((start, (buf + t).strip()))
            buf = None
    if buf is not None:
        result.append((start, buf.strip()))
    return result


def split_top(text):
    """Split on commas that are not inside parentheses."""
    parts, depth, cur = [], 0, []
    for c in text:
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        if c == "," and depth == 0:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(c)
    parts.append("".join(cur))
    return [p.strip() for p in parts]


def first_identifier(part):
    part = MODIFIER_RE.sub("", part.strip())
    m = re.match(r"\w+", part)
    return m.group(0) if m else ""


def block_events(text):
    """Return (kind, 'open'|'close') for a logical line, or None."""
    if re.match(r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?Sub\s+\w", text, re.I):
        return "Sub", "open"
    if re.match(r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?Function\s+\w", text, re.I):
        return "Function", "open"
    m = re.match(r"^End\s+(Sub|Function|With|Select|If)\b", text, re.I)
    if m:
        k = m.group(1).capitalize()
        return ("Select Case" if k == "Select" else k), "close"
    if re.match(r"^For\b", text, re.I):
        return "For", "open"
    if re.match(r"^Next\b", text, re.I):
        return "For", "close"
    if re.match(r"^Do\b", text, re.I):
        return "Do", "open"
    if re.match(r"^Loop\b", text, re.I):
        return "Do", "close"
    if re.match(r"^With\b", text, re.I):
        return "With", "open"
    if re.match(r"^Select\s+Case\b", text, re.I):
        return "Select Case", "open"
    if re.match(r"^If\b", text, re.I) and re.search(r"\bThen$", text, re.I):
        return "If", "open"
    return None


def check_text(name, data):
    msgs = []

    def add(line, text):
        msgs.append("%s:%d: %s" % (name, line, text))

    text = data.decode("ascii", "replace")
    raw_lines = text.split("\n")
    if raw_lines and raw_lines[-1] == "":
        raw_lines.pop()
    raw_lines = [l[:-1] if l.endswith("\r") else l for l in raw_lines]

    # 1. non-ASCII
    for no, line in enumerate(data.split(b"\n"), 1):
        if any(b > 127 for b in line):
            add(no, "non-ASCII character")

    # 2. line endings (first only)
    idx = data.find(b"\n")
    while idx != -1:
        if idx == 0 or data[idx - 1:idx] != b"\r":
            add(data.count(b"\n", 0, idx) + 1, "line ending is not CRLF")
            break
        idx = data.find(b"\n", idx + 1)

    # 3. VB_Name
    if not raw_lines or raw_lines[0] != 'Attribute VB_Name = "%s"' % name:
        add(1, "missing or wrong Attribute VB_Name")

    # 4. Option Explicit
    if "Option Explicit" not in raw_lines:
        add(1, "missing Option Explicit")

    stripped = [strip_line(l) for l in raw_lines]
    logical = logical_lines(stripped)

    # 5. Option Compare Text
    for no, line in enumerate(stripped, 1):
        if re.fullmatch(r"\s*Option\s+Compare\s+Text\s*", line, re.I):
            add(no, "Option Compare Text is not allowed")

    # 6. As Integer
    for no, line in enumerate(stripped, 1):
        if re.search(r"\bAs\s+Integer\b", line, re.I):
            add(no, "use Long, not Integer")

    # 7. one typed variable per declaration
    for no, line in logical:
        m = DECL_RE.match(line)
        if m:
            parts = split_top(m.group(1))
            if len(parts) > 1 and any(not AS_RE.search(" " + p + " ") for p in parts):
                add(no, "one typed variable per declaration")

    # 8. parameter modifiers
    for no, line in logical:
        params = proc_params(line)
        if params is not None:
            params = params.strip()
            if params:
                for p in split_top(params):
                    if not PARAM_OK_RE.match(p + " "):
                        add(no, "parameter needs ByVal or ByRef")

    # 9. block balance
    stack = []
    for no, line in logical:
        ev = block_events(line)
        if ev is None:
            continue
        kind, what = ev
        if what == "open":
            stack.append((kind, no))
            continue
        for i in range(len(stack) - 1, -1, -1):
            if stack[i][0] == kind:
                for k, ln in stack[i + 1:]:
                    add(ln, "unbalanced " + k)
                del stack[i:]
                break
        else:
            add(no, "unbalanced " + kind)
    for k, ln in stack:
        add(ln, "unbalanced " + k)

    # 10. reserved identifiers
    for no, line in logical:
        names = []
        m = DECL_RE.match(line) or CONST_RE.match(line)
        if m:
            names = [first_identifier(p) for p in split_top(m.group(1))]
        else:
            params = proc_params(line)
            if params is not None and params.strip():
                names = [first_identifier(p) for p in split_top(params)]
        for n in names:
            if n.lower() in RESERVED:
                add(no, "reserved name used as identifier")

    return msgs


def main(argv):
    if len(argv) != 2:
        print("usage: check_vba.py <dir>", file=sys.stderr)
        return 2
    folder = pathlib.Path(argv[1])
    if not folder.is_dir():
        print("not a directory: %s" % folder, file=sys.stderr)
        return 2
    found = []
    for path in sorted(folder.glob("*.bas")):
        found.extend(check_text(path.stem, path.read_bytes()))
    for m in found:
        print(m)
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
