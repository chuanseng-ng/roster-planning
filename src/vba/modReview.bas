Attribute VB_Name = "modReview"
Option Explicit

' Review rows, rule checker and summary text. Port of python/roster/review.py.
' gRev columns: 1 date text, 2 group, 3 slot, 4 rule, 5 reason, 6 candidates.

Public gRevCount As Long
Public gRev() As String

' Appends part to acc, separated by "; ". Parts are never empty.
Private Sub AppendPart(ByRef acc As String, ByVal part As String)
    If Len(acc) = 0 Then
        acc = part
    Else
        acc = acc & "; " & part
    End If
End Sub

' Appends a checker message to acc, separated by a line feed.
Private Sub AppendMsg(ByRef acc As String, ByVal msg As String)
    If Len(acc) = 0 Then
        acc = msg
    Else
        acc = acc & vbLf & msg
    End If
End Sub

' Python _why. rule is the first blocking rule, k the slot index.
Private Function WhyText(ByVal rule As Long, ByVal k As Long) As String
    If rule = 3 Then
        WhyText = "already rostered"
    ElseIf rule = 4 Then
        WhyText = "exempt"
    ElseIf rule = 5 Then
        WhyText = "on leave"
    ElseIf rule = 6 Then
        WhyText = "cannot do " & KindName(gSlotKind(k))
    Else
        WhyText = "rest after night"
    End If
End Function

' Python _cand_text.
Private Function CandText(ByVal rule As Long) As String
    If rule = 4 Then
        CandText = "rule 4 exemption"
    ElseIf rule = 5 Then
        CandText = "rule 5 leave"
    ElseIf rule = 6 Then
        CandText = "rule 6 cannot do"
    Else
        CandText = "rule 7 rest after night"
    End If
End Function

' Stores one row in the temporary array tmp at row index n (1 based).
Private Sub PutRow(ByRef tmp() As String, ByVal n As Long, ByVal dateTxt As String, ByVal grp As String, _
                   ByVal slotTxt As String, ByVal ruleTxt As String, ByVal reasonTxt As String, _
                   ByVal candTxt As String)
    tmp(n, 1) = dateTxt
    tmp(n, 2) = grp
    tmp(n, 3) = slotTxt
    tmp(n, 4) = ruleTxt
    tmp(n, 5) = reasonTxt
    tmp(n, 6) = candTxt
End Sub

' Python _gap_row: fills row n of tmp for the GAP cell (d, k).
Private Sub GapRow(ByRef tmp() As String, ByVal n As Long, ByVal d As Long, ByVal k As Long)
    Dim code As Long
    Dim s As Long
    Dim j As Long
    Dim considered As Boolean
    Dim placed As Boolean
    Dim rules As String
    Dim first As Long
    Dim reasons As String
    Dim cands As String
    Dim reasonText As String

    code = gGrid(d, k)
    reasons = ""
    cands = ""
    For s = 0 To gStaffCount - 1
        considered = True
        If code = GAP_SENIOR Then
            considered = gSenior(s)
        ElseIf code = GAP_COAG Then
            considered = gCoag(s)
        End If
        If considered Then
            rules = BlockingRules(s, d, k)
            If Len(rules) = 0 Then
                AppendPart reasons, gStaff(s) & ": not placed"
            Else
                first = CLng(Left$(rules, 1))
                AppendPart reasons, gStaff(s) & ": " & WhyText(first, k)
            End If
            placed = False
            For j = 0 To gSlotCount - 1
                If gGrid(d, j) = s Then placed = True
            Next j
            If Not placed Then
                If Len(rules) = 1 Then
                    first = CLng(rules)
                    If first >= 4 And first <= 7 Then
                        AppendPart cands, gStaff(s) & " (" & CandText(first) & ")"
                    End If
                End If
            End If
        End If
    Next s
    If Len(reasons) = 0 Then
        reasonText = "no staff with this requirement"
    Else
        reasonText = reasons
    End If
    PutRow tmp, n, TextFromSerial(gDates(d)), GroupName(gSlotGroup(k)), gSlotName(k), RuleText(code), reasonText, cands
End Sub

' Python build_review. Fills gRev and gRevCount; gRev is unallocated when empty.
Public Sub BuildReview()
    Dim tmp() As String
    Dim bound As Long
    Dim n As Long
    Dim d As Long
    Dim k As Long
    Dim i As Long
    Dim v As Long
    Dim r As Long
    Dim c As Long

    gRevCount = 0
    Erase gRev
    If gDateCount < 1 Then Exit Sub

    bound = gDateCount * (gSlotCount + 5)
    ReDim tmp(1 To bound, 1 To 6)
    n = 0
    For d = 0 To gDateCount - 1
        If gCutShort(d) Then
            n = n + 1
            PutRow tmp, n, TextFromSerial(gDates(d)), "", "", "Search cut short", "best roster found so far was kept", ""
        End If
        For k = 0 To gSlotCount - 1
            v = gGrid(d, k)
            If v = GAP Or v = GAP_SENIOR Or v = GAP_COAG Then
                n = n + 1
                GapRow tmp, n, d, k
            End If
        Next k
        For i = 0 To gUnmetCount(d) - 1
            n = n + 1
            PutRow tmp, n, TextFromSerial(gDates(d)), GroupName(gUnmetGroup(d, i)), "", RuleText(gUnmetCode(d, i)), _
                   "no free slot in this group", ""
        Next i
    Next d

    gRevCount = n
    If n < 1 Then Exit Sub
    ReDim gRev(1 To n, 1 To 6)
    For r = 1 To n
        For c = 1 To 6
            gRev(r, c) = tmp(r, c)
        Next c
    Next r
End Sub

' Header row and rows of gRev written to the Review sheet; column G stays empty.
Public Sub WriteReview()
    Dim ws As Worksheet
    Dim r As Long
    Dim c As Long

    Set ws = ThisWorkbook.Worksheets("Review")
    ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 7)).ClearContents
    ws.Cells(1, 1).Value = "Date"
    ws.Cells(1, 2).Value = "Group"
    ws.Cells(1, 3).Value = "Slot"
    ws.Cells(1, 4).Value = "Rule not met"
    ws.Cells(1, 5).Value = "Reason"
    ws.Cells(1, 6).Value = "Candidates"
    ws.Cells(1, 7).Value = "Approve"
    If gRevCount < 1 Then Exit Sub

    ws.Range(ws.Cells(2, 1), ws.Cells(1 + gRevCount, 1)).NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(2, 2), ws.Cells(1 + gRevCount, 7)).NumberFormat = "@"
    For r = 1 To gRevCount
        ws.Cells(1 + r, 1).Value2 = SerialFromText(gRev(r, 1))
        For c = 2 To 6
            ws.Cells(1 + r, c).Value = gRev(r, c)
        Next c
    Next r
End Sub

' Python _approved: an approval for exactly cell (d, k) listing the rule.
Private Function CellApproved(ByVal d As Long, ByVal k As Long, ByVal rule As Long) As Boolean
    Dim a As Long

    CellApproved = False
    For a = 0 To gApprCount - 1
        If gApprD(a) = d Then
            If gApprK(a) = k Then
                If HasRule(Replace(gApprRules(a), " ", ""), rule) Then
                    CellApproved = True
                    Exit Function
                End If
            End If
        End If
    Next a
End Function

' Python _group_approved: an approval for any cell of group g on date index d listing the rule.
Private Function GroupApproved(ByVal d As Long, ByVal g As Long, ByVal rule As Long) As Boolean
    Dim a As Long

    GroupApproved = False
    For a = 0 To gApprCount - 1
        If gApprD(a) = d Then
            If gSlotGroup(gApprK(a)) = g Then
                If HasRule(Replace(gApprRules(a), " ", ""), rule) Then
                    GroupApproved = True
                    Exit Function
                End If
            End If
        End If
    Next a
End Function

' Python check_rules. Messages joined with vbLf; "" when there are none.
Public Function CheckRules() As String
    Dim out As String
    Dim d As Long
    Dim k As Long
    Dim g As Long
    Dim s As Long
    Dim v As Long
    Dim i As Long
    Dim rules As String
    Dim parts() As String
    Dim r As Long
    Dim nActive As Long
    Dim negative As Boolean
    Dim hasSenior As Boolean
    Dim hasCoag As Boolean
    Dim unmet As Boolean
    Dim label As String

    out = ""
    If gDateCount < 1 Then
        CheckRules = out
        Exit Function
    End If
    If gSlotCount < 1 Then
        CheckRules = out
        Exit Function
    End If

    For d = 0 To gDateCount - 1
        For k = 0 To gSlotCount - 1
            s = gGrid(d, k)
            If s >= 0 Then
                rules = BlockingRules(s, d, k)
                If Len(rules) > 0 Then
                    parts = Split(rules, ",")
                    For i = 0 To UBound(parts)
                        r = CLng(parts(i))
                        If Not CellApproved(d, k, r) Then
                            AppendMsg out, TextFromSerial(gDates(d)) & " " & gSlotName(k) & ": " & _
                                gStaff(s) & " breaks rule " & CStr(r)
                        End If
                    Next i
                End If
            End If
        Next k
    Next d

    For d = 0 To gDateCount - 1
        For g = 0 To 1
            nActive = 0
            negative = False
            hasSenior = False
            hasCoag = False
            For k = 0 To gSlotCount - 1
                If gSlotGroup(k) = g Then
                    v = gGrid(d, k)
                    If v < 0 Then
                        negative = True
                    Else
                        If gSlotActive(k) Then
                            If gSenior(v) Then hasSenior = True
                        End If
                        If gCoag(v) Then hasCoag = True
                    End If
                End If
            Next k
            For k = 0 To gSlotCount - 1
                If gSlotGroup(k) = g Then
                    If gSlotActive(k) Then nActive = nActive + 1
                End If
            Next k
            If nActive > 0 And Not negative Then
                unmet = False
                For i = 0 To gUnmetCount(d) - 1
                    If gUnmetGroup(d, i) = g Then unmet = True
                Next i
                If Not unmet Then
                    label = TextFromSerial(gDates(d)) & " " & GroupName(g) & ": "
                    If gRuleOn(1) And Not hasSenior Then
                        If Not GroupApproved(d, g, 1) Then
                            AppendMsg out, label & "rule 1 not met"
                        End If
                    End If
                    If gRuleOn(2) And Not hasCoag Then
                        If Not GroupApproved(d, g, 2) Then
                            AppendMsg out, label & "rule 2 not met"
                        End If
                    End If
                End If
            End If
        Next g
    Next d
    CheckRules = out
End Function

' Python _q: ratio a/e rounded to four decimals, as a whole number of 1/10000.
Private Function RatioQ(ByVal a As Long, ByVal e As Long) As Long
    RatioQ = (a * 20000 + e) \ (2 * e)
End Function

' Python _q_text.
Private Function QText(ByVal q As Long) As String
    QText = CStr(q \ 10000) & "." & Format$(q Mod 10000, "0000")
End Function

Public Function RatioText(ByVal a As Long, ByVal e As Long) As String
    If e = 0 Then
        RatioText = ""
    Else
        RatioText = QText(RatioQ(a, e))
    End If
End Function

' Python build_summary, spread part: highest minus lowest ratio over staff with eligible dates.
Public Function SpreadText() As String
    Dim s As Long
    Dim q As Long
    Dim qmax As Long
    Dim qmin As Long
    Dim have As Boolean

    have = False
    qmax = 0
    qmin = 0
    If gStaffCount >= 1 Then
        For s = 0 To gStaffCount - 1
            If gEligible(s) > 0 Then
                q = RatioQ(gActiveCount(s), gEligible(s))
                If Not have Then
                    qmax = q
                    qmin = q
                    have = True
                Else
                    If q > qmax Then qmax = q
                    If q < qmin Then qmin = q
                End If
            End If
        Next s
    End If
    If Not have Then
        SpreadText = "0.0000"
    Else
        SpreadText = QText(qmax - qmin)
    End If
End Function
