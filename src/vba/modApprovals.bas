Attribute VB_Name = "modApprovals"
Option Explicit

' Approvals of GAP cells. Port of ops.approval_broken_rules and ops.apply_approval,
' plus the Review-sheet driven macro DoApplyApprovals.

' Sets flags(1 To 7) for every rule number in the comma-separated list.
Private Sub FlagsFromList(ByVal ruleList As String, ByRef flags() As Boolean)
    Dim parts() As String
    Dim i As Long
    Dim p As String

    If Len(ruleList) = 0 Then Exit Sub
    parts = Split(ruleList, ",")
    For i = 0 To UBound(parts)
        p = Trim$(parts(i))
        If Len(p) = 1 Then
            If InStr(1, "1234567", p, vbBinaryCompare) > 0 Then flags(CLng(p)) = True
        End If
    Next i
End Sub

' Comma-separated ascending list of the set flags.
Private Function ListFromFlags(ByRef flags() As Boolean) As String
    Dim r As Long
    Dim out As String

    out = ""
    For r = 1 To 7
        If flags(r) Then
            If Len(out) = 0 Then
                out = CStr(r)
            Else
                out = out & "," & CStr(r)
            End If
        End If
    Next r
    ListFromFlags = out
End Function

' Python approval_broken_rules: rules that staff s taking slot k on date d would break.
Public Function ApprovalBrokenRules(ByVal d As Long, ByVal k As Long, ByVal s As Long) As String
    Dim flags(1 To 7) As Boolean
    Dim blk As String
    Dim parts() As String
    Dim i As Long
    Dim r As Long

    blk = modRules.BlockingRules(s, d, k)
    If Len(blk) > 0 Then
        parts = Split(blk, ",")
        For i = 0 To UBound(parts)
            r = CLng(parts(i))
            If r <> 3 Then
                If r >= 1 Then
                    If r <= 7 Then flags(r) = True
                End If
            End If
        Next i
    End If
    If modRules.ForwardNightBlock(s, d, k) Then flags(7) = True
    If gGrid(d, k) = GAP_SENIOR Then
        If Not gSenior(s) Then flags(1) = True
    End If
    If gGrid(d, k) = GAP_COAG Then
        If Not gCoag(s) Then flags(2) = True
    End If
    ApprovalBrokenRules = ListFromFlags(flags)
End Function

' Python apply_approval. Returns False with the reason in note, or True with
' note holding "now also breaks" remarks (possibly empty). No dialogs.
Public Function ApplyApprovalCore(ByVal d As Long, ByVal k As Long, ByVal s As Long, ByRef note As String) As Boolean
    Dim j As Long
    Dim a As Long
    Dim n As Long
    Dim i As Long
    Dim rules As String
    Dim blk As String
    Dim parts() As String
    Dim have(1 To 7) As Boolean
    Dim extra(1 To 7) As Boolean
    Dim extraCount As Long
    Dim extraText As String
    Dim r As Long

    note = ""
    ApplyApprovalCore = False
    If gGrid(d, k) >= 0 Then
        note = "slot is no longer a GAP"
        Exit Function
    End If
    For j = 0 To gSlotCount - 1
        If gGrid(d, j) = s Then
            note = gStaff(s) & " already has a slot on this date"
            Exit Function
        End If
    Next j

    rules = ApprovalBrokenRules(d, k, s)
    n = gApprCount
    If n = 0 Then
        ReDim gApprD(0 To 0)
        ReDim gApprK(0 To 0)
        ReDim gApprS(0 To 0)
        ReDim gApprRules(0 To 0)
    Else
        ReDim Preserve gApprD(0 To n)
        ReDim Preserve gApprK(0 To n)
        ReDim Preserve gApprS(0 To n)
        ReDim Preserve gApprRules(0 To n)
    End If
    gApprD(n) = d
    gApprK(n) = k
    gApprS(n) = s
    gApprRules(n) = rules
    gApprCount = n + 1

    gGrid(d, k) = s
    gApproved(d, k) = True
    modEngine.SolveYear d + 1

    For a = 0 To gApprCount - 1
        If gApprD(a) > d Then
            For i = 1 To 7
                have(i) = False
                extra(i) = False
            Next i
            FlagsFromList gApprRules(a), have
            extraCount = 0
            blk = modRules.BlockingRules(gApprS(a), gApprD(a), gApprK(a))
            If Len(blk) > 0 Then
                parts = Split(blk, ",")
                For i = 0 To UBound(parts)
                    r = CLng(parts(i))
                    If r <> 3 Then
                        If r >= 1 Then
                            If r <= 7 Then
                                If Not have(r) Then
                                    If Not extra(r) Then
                                        extra(r) = True
                                        extraCount = extraCount + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
            If extraCount > 0 Then
                extraText = ""
                For r = 1 To 7
                    If extra(r) Then
                        have(r) = True
                        If Len(extraText) > 0 Then extraText = extraText & ", "
                        extraText = extraText & CStr(r)
                    End If
                Next r
                gApprRules(a) = ListFromFlags(have)
                If Len(note) > 0 Then note = note & "; "
                note = note & "approval for " & gStaff(gApprS(a)) & " on " & TextFromSerial(gDates(gApprD(a))) & _
                       " now also breaks rule " & extraText
            End If
        End If
    Next a
    ApplyApprovalCore = True
End Function

' ---- Review sheet driven macro ----

Private Function CellString(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Dim v As Variant

    CellString = ""
    v = ws.Cells(r, c).Value2
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    CellString = Trim$(CStr(v))
End Function

Private Function CellSerialOf(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Long
    Dim v As Variant
    Dim t As String

    CellSerialOf = -1
    v = ws.Cells(r, c).Value2
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Then
        CellSerialOf = CLng(Int(CDbl(v)))
        Exit Function
    End If
    t = Trim$(CStr(v))
    If IsDateText(t) Then CellSerialOf = SerialFromText(t)
End Function

Private Function DateIdx(ByVal serial As Long) As Long
    Dim d As Long

    DateIdx = -1
    For d = 0 To gDateCount - 1
        If gDates(d) = serial Then
            DateIdx = d
            Exit Function
        End If
    Next d
End Function

Private Function SlotIdx(ByVal t As String) As Long
    Dim k As Long
    Dim key As String

    SlotIdx = -1
    key = Norm(t)
    For k = 0 To gSlotCount - 1
        If StrComp(Norm(gSlotName(k)), key, vbBinaryCompare) = 0 Then
            SlotIdx = k
            Exit Function
        End If
    Next k
End Function

Private Sub AddLine(ByRef acc As String, ByVal t As String)
    If Len(acc) = 0 Then
        acc = t
    Else
        acc = acc & vbLf & t
    End If
End Sub

Public Sub DoApplyApprovals()
    Dim msg As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim c As Long
    Dim lr As Long
    Dim r As Long
    Dim n As Long
    Dim cap As Long
    Dim itD() As Long
    Dim itK() As Long
    Dim itRow() As Long
    Dim itName() As String
    Dim ap As String
    Dim slotTxt As String
    Dim d As Long
    Dim k As Long
    Dim s As Long
    Dim i As Long
    Dim j As Long
    Dim tD As Long
    Dim tK As Long
    Dim tRow As Long
    Dim tName As String
    Dim lines As String
    Dim wanted As Long
    Dim applied As Long
    Dim rules As String
    Dim prompt As String
    Dim note As String
    Dim has As Boolean
    Dim chk As String

    modIO.ReadSheetsToRaw
    msg = modValidate.ValidateRaw()
    If Len(msg) > 0 Then
        modMain.ShowMessages "Input problems", msg
        Exit Sub
    End If
    modIO.BuildInputs
    msg = modIO.LoadResultFromSheets()
    If Len(msg) > 0 Then
        modMain.ShowMessages "Approvals", msg
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("Review")
    lastRow = 1
    For c = 1 To 7
        lr = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If lr > lastRow Then lastRow = lr
    Next c
    cap = lastRow
    ReDim itD(1 To cap)
    ReDim itK(1 To cap)
    ReDim itRow(1 To cap)
    ReDim itName(1 To cap)
    n = 0
    wanted = 0
    lines = ""
    For r = 2 To lastRow
        ap = CellString(ws, r, 7)
        If Len(ap) > 0 Then
            wanted = wanted + 1
            slotTxt = CellString(ws, r, 3)
            If Len(slotTxt) = 0 Then
                AddLine lines, "Skipped: Review row " & CStr(r) & " is not approvable (no slot)"
            Else
                d = DateIdx(CellSerialOf(ws, r, 1))
                k = SlotIdx(slotTxt)
                If d < 0 Then
                    AddLine lines, "Skipped: Review row " & CStr(r) & " has a date that is not on the roster"
                ElseIf k < 0 Then
                    AddLine lines, "Skipped: Review row " & CStr(r) & " has an unknown slot " & slotTxt
                Else
                    n = n + 1
                    itD(n) = d
                    itK(n) = k
                    itRow(n) = r
                    itName(n) = ap
                End If
            End If
        End If
    Next r
    If wanted = 0 Then
        modMain.ShowMessages "Approvals", "No entries in the Approve column of the Review sheet."
        Exit Sub
    End If

    ' stable insertion sort by date, then slot
    For i = 2 To n
        tD = itD(i)
        tK = itK(i)
        tRow = itRow(i)
        tName = itName(i)
        j = i - 1
        Do While j >= 1
            has = False
            If itD(j) > tD Then
                has = True
            ElseIf itD(j) = tD Then
                If itK(j) > tK Then has = True
            End If
            If has Then
                itD(j + 1) = itD(j)
                itK(j + 1) = itK(j)
                itRow(j + 1) = itRow(j)
                itName(j + 1) = itName(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        itD(j + 1) = tD
        itK(j + 1) = tK
        itRow(j + 1) = tRow
        itName(j + 1) = tName
    Next i

    applied = 0
    For i = 1 To n
        d = itD(i)
        k = itK(i)
        s = StaffIndex(itName(i))
        If s < 0 Then
            AddLine lines, "Skipped: Review row " & CStr(itRow(i)) & " name " & itName(i) & " is not on the Staff sheet"
        ElseIf gGrid(d, k) >= 0 Or gGrid(d, k) = EMPTY_CELL Then
            AddLine lines, "Skipped: " & gStaff(s) & " on " & TextFromSerial(gDates(d)) & " " & gSlotName(k) & _
                           ": slot is no longer a GAP"
        Else
            has = False
            For j = 0 To gSlotCount - 1
                If gGrid(d, j) = s Then has = True
            Next j
            If has Then
                AddLine lines, "Skipped: " & gStaff(s) & " on " & TextFromSerial(gDates(d)) & " " & gSlotName(k) & _
                               ": already has a slot on this date"
            Else
                rules = ApprovalBrokenRules(d, k, s)
                prompt = gStaff(s) & " into " & gSlotName(k) & " on " & TextFromSerial(gDates(d))
                If Len(rules) = 0 Then
                    prompt = prompt & " breaks no rule. Approve?"
                Else
                    prompt = prompt & " breaks rule " & Replace(rules, ",", ", ") & ". Approve?"
                End If
                If MsgBox(prompt, vbYesNo + vbQuestion, "PH Roster Planner") = vbYes Then
                    note = ""
                    If ApplyApprovalCore(d, k, s, note) Then
                        applied = applied + 1
                        AddLine lines, "Applied: " & gStaff(s) & " into " & gSlotName(k) & " on " & TextFromSerial(gDates(d))
                        If Len(note) > 0 Then AddLine lines, "Note: " & note
                    Else
                        AddLine lines, "Skipped: " & gStaff(s) & " on " & TextFromSerial(gDates(d)) & " " & _
                                       gSlotName(k) & ": " & note
                    End If
                Else
                    AddLine lines, "Skipped: " & gStaff(s) & " on " & TextFromSerial(gDates(d)) & " " & _
                                   gSlotName(k) & ": not confirmed"
                End If
            End If
        End If
    Next i

    If applied = 0 Then
        modMain.ShowMessages "Approvals", "Nothing was applied; the sheets are unchanged." & vbLf & lines
        Exit Sub
    End If

    modReview.BuildReview
    modIO.WriteRoster
    modReview.WriteReview
    modIO.WriteSummary
    modIO.WriteApprovals
    chk = modReview.CheckRules()
    If Len(chk) > 0 Then
        AddLine lines, "Rule check:"
        AddLine lines, chk
    End If
    modMain.ShowMessages "Approvals", lines
End Sub
