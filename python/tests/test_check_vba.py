import importlib.util, pathlib, subprocess, sys
REPO = pathlib.Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location(
    "check_vba", REPO / "tools" / "check_vba.py")
check_vba = importlib.util.module_from_spec(spec); spec.loader.exec_module(check_vba)

def src(*lines, name="modX"):
    head = ['Attribute VB_Name = "%s"' % name, "Option Explicit"]
    return "\r\n".join(head + list(lines) + [""]).encode("ascii")

def test_clean_module_has_no_findings():
    data = src("Public Sub A(ByVal n As Long)", "    Dim i As Long",
               "    For i = 1 To n", "        If i > 1 Then", "            i = i",
               "        End If", "    Next i", "End Sub")
    assert check_vba.check_text("modX", data) == []

def test_integer_is_reported_with_line():
    msgs = check_vba.check_text("modX", src("Private x As Integer"))
    assert msgs == ["modX:3: use Long, not Integer"]

def test_non_ascii_byte():
    data = src("Private x As Long ' café".replace("é", "?"))
    data = data.replace(b"?", b"\xe9")
    assert check_vba.check_text("modX", data) == ["modX:3: non-ASCII character"]

def test_bare_lf_reported_once():
    data = b'Attribute VB_Name = "modX"\r\nOption Explicit\nPrivate x As Long\nPrivate y As Long\r\n'
    assert check_vba.check_text("modX", data) == ["modX:2: line ending is not CRLF"]

def test_wrong_vb_name():
    msgs = check_vba.check_text("modX", src(name="modY"))
    assert msgs == ["modX:1: missing or wrong Attribute VB_Name"]

def test_missing_option_explicit():
    data = b'Attribute VB_Name = "modX"\r\nPrivate x As Long\r\n'
    assert check_vba.check_text("modX", data) == ["modX:1: missing Option Explicit"]

def test_option_compare_text():
    msgs = check_vba.check_text("modX", src("Option Compare Text"))
    assert msgs == ["modX:3: Option Compare Text is not allowed"]

def test_multi_declaration_without_types():
    msgs = check_vba.check_text("modX", src("Public Sub A()", "    Dim a, b As Long", "End Sub"))
    assert msgs == ["modX:4: one typed variable per declaration"]

def test_multi_declaration_all_typed_is_ok():
    msgs = check_vba.check_text("modX", src("Private a As Long, b As Long",
                                           "Private arr(1, 2) As Long"))
    assert msgs == []

def test_array_return_types_and_array_params_are_ok():
    for line in ("Private Function H() As Long()",
                 "Private Function K(ByVal a As Long) As String()",
                 "Public Sub S(ByRef b() As String, ByVal n As Long)"):
        end = "End Sub" if "Sub" in line else "End Function"
        assert check_vba.check_text("modX", src(line, end)) == [], line

def test_parameter_without_byval_with_array_return_type():
    msgs = check_vba.check_text("modX", src("Private Function M(a As Long) As Long()",
                                           "End Function"))
    assert msgs == ["modX:3: parameter needs ByVal or ByRef"]

def test_parameter_without_byval():
    msgs = check_vba.check_text("modX", src("Public Sub A(n As Long)", "End Sub"))
    assert msgs == ["modX:3: parameter needs ByVal or ByRef"]

def test_parameters_across_continuation_and_empty_list():
    msgs = check_vba.check_text("modX", src(
        "Public Function F(ByVal a As Long, _", "                  Optional ByVal b As Long = 0) As Long",
        "End Function", "Public Sub G()", "End Sub"))
    assert msgs == []

def test_if_without_end_if():
    msgs = check_vba.check_text("modX", src("Public Sub A()", "    If x Then", "End Sub"))
    assert msgs == ["modX:4: unbalanced If"]

def test_single_line_if_is_not_a_block():
    msgs = check_vba.check_text("modX", src("Public Sub A()", "    If x Then y = 1", "End Sub"))
    assert msgs == []

def test_for_without_next():
    msgs = check_vba.check_text("modX", src("Public Sub A()", "    For i = 1 To 3", "End Sub"))
    assert msgs == ["modX:4: unbalanced For"]

def test_stray_closer():
    msgs = check_vba.check_text("modX", src("Public Sub A()", "End Sub", "End If"))
    assert msgs == ["modX:5: unbalanced If"]

def test_exit_statements_and_else_open_nothing():
    msgs = check_vba.check_text("modX", src(
        "Public Sub A()", "    If x Then", "        Exit Sub", "    ElseIf y Then",
        "        Exit Sub", "    Else", "        Exit Sub", "    End If", "End Sub"))
    assert msgs == []

def test_reserved_name_variable():
    msgs = check_vba.check_text("modX", src("Public Sub A()", "    Dim Name As String", "End Sub"))
    assert msgs == ["modX:4: reserved name used as identifier"]

def test_reserved_name_parameter():
    msgs = check_vba.check_text("modX", src("Public Sub A(ByVal Date As Long)", "End Sub"))
    assert msgs == ["modX:3: reserved name used as identifier"]

def test_integer_in_comment_and_string_not_reported():
    msgs = check_vba.check_text("modX", src(
        "Private x As Long ' was As Integer",
        'Private s As String: s = "As Integer"',
        'Private t As String: t = "a ""quoted"" As Integer ' + "' not a comment" + '"'))
    assert msgs == []

def test_cli_exit_codes(tmp_path):
    script = str(REPO / "tools" / "check_vba.py")
    r = subprocess.run([sys.executable, script, str(tmp_path)], capture_output=True, text=True)
    assert r.returncode == 0
    (tmp_path / "modX.bas").write_bytes(src("Private x As Integer"))
    r = subprocess.run([sys.executable, script, str(tmp_path)], capture_output=True, text=True)
    assert r.returncode == 1
    assert "modX:3: use Long, not Integer" in r.stdout
