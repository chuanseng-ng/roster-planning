Attribute VB_Name = "modIO"
Option Explicit

' Reading sheets into raw tables, building inputs, writing and reloading output.
' Ports python/roster/csvio.py (read_table) and python/roster/inputs.py.

' ------------------------------------------------------------------
' Reading
' ------------------------------------------------------------------

' Text of one cell value. Date columns hold Excel serials: convert those to
' yyyy-mm-dd text. Everything else is Trim$ of its text. Errors and empty cells give "".
Private Function CellToText(ByVal v As Variant, ByVal isDateCol As Boolean) As String
    Dim dbl As Double

    CellToText = ""
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If isDateCol Then
        If VarType(v) = vbDouble Then
            dbl = CDbl(v)
            If dbl >= 0 And dbl <= 2958465 Then
                CellToText = TextFromSerial(CLng(Int(dbl)))
                Exit Function
            End If
        End If
    End If
    CellToText = Trim$(CStr(v))
End Function

' Last used row among the first `width` columns of ws (1 when the sheet is empty).
Private Function LastRowOf(ByVal ws As Worksheet, ByVal width As Long) As Long
    Dim c As Long
    Dim lr As Long
    Dim best As Long

    best = 1
    For c = 1 To width
        lr = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If lr > best Then best = lr
    Next c
    LastRowOf = best
End Function

' Reads rows 2 and below of a sheet into tbl(1 To rowCount, 1 To width).
' Fully blank rows are dropped. lastRowFixed > 0 fixes the last row read.
Private Sub ReadTable(ByVal sheetName As String, ByVal width As Long, ByVal dateColA As Long, _
                      ByVal dateColB As Long, ByVal lastRowFixed As Long, _
                      ByRef tbl() As String, ByRef rowCount As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim c As Long
    Dim n As Long
    Dim i As Long
    Dim blank As Boolean
    Dim t As String
    Dim isDateCol As Boolean
    Dim buf() As String

    rowCount = 0
    Erase tbl
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If lastRowFixed > 0 Then
        lastRow = lastRowFixed
    Else
        lastRow = LastRowOf(ws, width)
    End If
    If lastRow < 2 Then Exit Sub
    ReDim buf(1 To lastRow - 1, 1 To width)
    n = 0
    For r = 2 To lastRow
        blank = True
        For c = 1 To width
            isDateCol = (c = dateColA Or c = dateColB)
            t = CellToText(ws.Cells(r, c).Value2, isDateCol)
            buf(n + 1, c) = t
            If t <> "" Then blank = False
        Next c
        If Not blank Then n = n + 1
    Next r
    If n = 0 Then Exit Sub
    ReDim tbl(1 To n, 1 To width)
    For r = 1 To n
        For i = 1 To width
            tbl(r, i) = buf(r, i)
        Next i
    Next r
    rowCount = n
End Sub

Public Sub ReadSheetsToRaw()
    ReadTable "Staff", 5, 0, 0, 0, gRawStaff, gRawStaffRows
    ReadTable "Holidays", 3, 2, 0, 0, gRawHolidays, gRawHolidaysRows
    ReadTable "LastYear", 4, 2, 0, 0, gRawLastYear, gRawLastYearRows
    ReadTable "Leave", 4, 2, 3, 0, gRawLeave, gRawLeaveRows
    ReadTable "Settings", 2, 0, 0, 13, gRawSettings, gRawSettingsRows
End Sub

' ------------------------------------------------------------------
' Building inputs
' ------------------------------------------------------------------

Private Function PieceMask(ByVal word As String) As Long
    Dim w As String

    w = Norm(word)
    If w = "DAY" Then
        PieceMask = M_DAY
    ElseIf w = "EVENING" Then
        PieceMask = M_EVE
    ElseIf w = "NIGHT" Then
        PieceMask = M_NIGHT
    ElseIf w = "ALL" Then
        PieceMask = M_ALL
    Else
        PieceMask = 0
    End If
End Function

' Port of inputs.build_inputs. The raw tables must have passed ValidateRaw.
Public Sub BuildInputs()
    Dim v As String
    Dim ok As Boolean
    Dim i As Long
    Dim r As Long
    Dim s As Long
    Dim d As Long
    Dim p As Long
    Dim pos As Long
    Dim j As Long
    Dim pieces() As String
    Dim w As String
    Dim mask As Long
    Dim hasCoag As Boolean
    Dim serial As Long
    Dim ex As Boolean
    Dim lv As Long
    Dim keyS As String
    Dim holName As String
    Dim fromSerial As Long
    Dim toSerial As Long

    ' settings
    ok = FindSetting("Year", v)
    gYear = CLng(Trim$(v))
    For i = 0 To 4
        ok = FindSetting(SettingsKey(i + 1), v)
        gCounts(i) = CLng(Trim$(v))
    Next i
    gRuleOn(3) = True
    gRuleOn(1) = RuleSwitch(1)
    gRuleOn(2) = RuleSwitch(2)
    gRuleOn(4) = RuleSwitch(4)
    gRuleOn(5) = RuleSwitch(5)
    gRuleOn(6) = RuleSwitch(6)
    gRuleOn(7) = RuleSwitch(7)
    BuildSlots

    ' staff
    gStaffCount = gRawStaffRows
    If gStaffCount > 0 Then
        ReDim gStaff(0 To gStaffCount - 1)
        ReDim gSenior(0 To gStaffCount - 1)
        ReDim gCoag(0 To gStaffCount - 1)
        ReDim gCannot(0 To gStaffCount - 1)
    Else
        Erase gStaff
        Erase gSenior
        Erase gCoag
        Erase gCannot
    End If
    For r = 1 To gRawStaffRows
        s = r - 1
        gStaff(s) = Trim$(gRawStaff(r, 1))
        gSenior(s) = (Norm(gRawStaff(r, 3)) = "Y")
        hasCoag = False
        pieces = Split(gRawStaff(r, 4), ",")
        For p = LBound(pieces) To UBound(pieces)
            If Norm(pieces(p)) = "SPECIAL COAG" Then hasCoag = True
        Next p
        gCoag(s) = hasCoag
        mask = 0
        If gRuleOn(6) Then
            pieces = Split(gRawStaff(r, 5), ",")
            For p = LBound(pieces) To UBound(pieces)
                w = Norm(pieces(p))
                If w = "DAY" Or w = "EVENING" Or w = "NIGHT" Then
                    mask = mask Or PieceMask(w)
                End If
            Next p
        End If
        gCannot(s) = mask
    Next r

    ' holidays sorted ascending by serial (insertion sort, stable)
    gDateCount = gRawHolidaysRows
    If gDateCount > 0 Then
        ReDim gDates(0 To gDateCount - 1)
        ReDim gHoliday(0 To gDateCount - 1)
    Else
        Erase gDates
        Erase gHoliday
    End If
    For r = 1 To gRawHolidaysRows
        serial = SerialFromText(Trim$(gRawHolidays(r, 2)))
        pos = r - 1
        Do While pos > 0
            If gDates(pos - 1) > serial Then
                pos = pos - 1
            Else
                Exit Do
            End If
        Loop
        For j = r - 1 To pos + 1 Step -1
            gDates(j) = gDates(j - 1)
            gHoliday(j) = gHoliday(j - 1)
        Next j
        gDates(pos) = serial
        gHoliday(pos) = Trim$(gRawHolidays(r, 1))
    Next r

    ' exempt and leave
    If gStaffCount > 0 And gDateCount > 0 Then
        ReDim gExempt(0 To gStaffCount - 1, 0 To gDateCount - 1)
        ReDim gLeave(0 To gStaffCount - 1, 0 To gDateCount - 1)
    Else
        Erase gExempt
        Erase gLeave
    End If
    For s = 0 To gStaffCount - 1
        keyS = Norm(gStaff(s))
        For d = 0 To gDateCount - 1
            holName = Norm(gHoliday(d))
            ex = False
            If gRuleOn(4) Then
                For i = 1 To gRawLastYearRows
                    If Norm(gRawLastYear(i, 3)) = keyS And Norm(gRawLastYear(i, 1)) = holName Then
                        If Left$(Norm(gRawLastYear(i, 4)), 7) <> "STANDBY" Then
                            ex = True
                        End If
                    End If
                Next i
            End If
            lv = 0
            If gRuleOn(5) Then
                For i = 1 To gRawLeaveRows
                    If Norm(gRawLeave(i, 1)) = keyS Then
                        fromSerial = SerialFromText(Trim$(gRawLeave(i, 2)))
                        toSerial = SerialFromText(Trim$(gRawLeave(i, 3)))
                        If fromSerial <= gDates(d) And gDates(d) <= toSerial Then
                            lv = lv Or PieceMask(gRawLeave(i, 4))
                        End If
                    End If
                Next i
            End If
            gExempt(s, d) = ex
            gLeave(s, d) = lv
        Next d
    Next s
End Sub

' True when the setting "Rule <n>" is Y.
Private Function RuleSwitch(ByVal n As Long) As Boolean
    Dim v As String
    Dim ok As Boolean

    ok = FindSetting("Rule " & CStr(n), v)
    RuleSwitch = (Norm(v) = "Y")
End Function

' ------------------------------------------------------------------
' Writing
' ------------------------------------------------------------------

Private Sub ClearBelowHeader(ByVal ws As Worksheet, ByVal lastCol As Long)
    Dim rng As Range

    Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, lastCol))
    rng.ClearContents
    rng.ClearComments
    rng.Interior.ColorIndex = xlNone
End Sub

' Rules text recorded for the approval at cell (d, k), or "" when none is recorded.
Private Function ApprovalRulesFor(ByVal d As Long, ByVal k As Long) As String
    Dim a As Long

    ApprovalRulesFor = ""
    For a = 0 To gApprCount - 1
        If gApprD(a) = d And gApprK(a) = k Then
            ApprovalRulesFor = gApprRules(a)
            Exit Function
        End If
    Next a
End Function

Public Sub WriteRoster()
    Dim ws As Worksheet
    Dim d As Long
    Dim k As Long
    Dim code As Long
    Dim r As Long
    Dim c As Long
    Dim lastCol As Long
    Dim rules As String
    Dim note As String
    Dim cell As Range

    Set ws = ThisWorkbook.Worksheets("Roster")
    lastCol = ws.Columns.Count

    ' clear everything below the header and every slot column of the header row
    ClearBelowHeader ws, lastCol
    ws.Range(ws.Cells(1, 3), ws.Cells(1, lastCol)).ClearContents

    ws.Cells(1, 1).Value = "Holiday"
    ws.Cells(1, 2).Value = "Date"
    For k = 0 To gSlotCount - 1
        ws.Cells(1, 3 + k).Value = gSlotName(k)
    Next k

    If gDateCount < 1 Then Exit Sub

    ws.Range(ws.Cells(2, 1), ws.Cells(1 + gDateCount, 1)).NumberFormat = "@"
    ws.Range(ws.Cells(2, 2), ws.Cells(1 + gDateCount, 2)).NumberFormat = "yyyy-mm-dd"
    If gSlotCount > 0 Then
        ws.Range(ws.Cells(2, 3), ws.Cells(1 + gDateCount, 2 + gSlotCount)).NumberFormat = "@"
    End If

    For d = 0 To gDateCount - 1
        r = 2 + d
        ws.Cells(r, 1).Value = gHoliday(d)
        ws.Cells(r, 2).Value2 = gDates(d)
        For k = 0 To gSlotCount - 1
            c = 3 + k
            code = gGrid(d, k)
            Set cell = ws.Cells(r, c)
            If code >= 0 Then
                cell.Value = gStaff(code)
            ElseIf code = EMPTY_CELL Then
                cell.Value = ""
            Else
                cell.Value = "GAP"
                cell.Interior.Color = RGB(255, 199, 206)
            End If
            If gApproved(d, k) Then
                cell.Interior.Color = RGB(255, 235, 156)
                rules = ApprovalRulesFor(d, k)
                If rules = "" Then
                    note = "Approved exception: no rule broken"
                Else
                    note = "Approved exception: rule " & rules
                End If
                cell.AddComment note
            End If
        Next k
    Next d
End Sub

Public Sub WriteSummary()
    Dim ws As Worksheet
    Dim s As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets("Summary")
    ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 5)).ClearContents
    ws.Cells(1, 1).Value = "Staff name"
    ws.Cells(1, 2).Value = "Active slots"
    ws.Cells(1, 3).Value = "Eligible dates"
    ws.Cells(1, 4).Value = "Ratio"
    ws.Cells(1, 5).Value = "Standby slots"

    ws.Range(ws.Cells(2, 1), ws.Cells(2 + gStaffCount, 1)).NumberFormat = "@"
    ws.Range(ws.Cells(2, 4), ws.Cells(2 + gStaffCount, 4)).NumberFormat = "@"
    For s = 0 To gStaffCount - 1
        r = 2 + s
        ws.Cells(r, 1).Value = gStaff(s)
        ws.Cells(r, 2).Value2 = gActiveCount(s)
        ws.Cells(r, 3).Value2 = gEligible(s)
        ws.Cells(r, 4).Value = modReview.RatioText(gActiveCount(s), gEligible(s))
        ws.Cells(r, 5).Value2 = gStandbyCount(s)
    Next s
    r = 2 + gStaffCount
    ws.Cells(r, 1).Value = "Spread"
    ws.Cells(r, 4).Value = modReview.SpreadText()
End Sub

Public Sub WriteApprovals()
    Dim ws As Worksheet
    Dim a As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets("Approvals")
    ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 4)).ClearContents
    ws.Cells(1, 1).Value = "Date"
    ws.Cells(1, 2).Value = "Slot"
    ws.Cells(1, 3).Value = "Staff name"
    ws.Cells(1, 4).Value = "Rules broken"
    If gApprCount < 1 Then Exit Sub

    ws.Range(ws.Cells(2, 1), ws.Cells(1 + gApprCount, 1)).NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(2, 2), ws.Cells(1 + gApprCount, 4)).NumberFormat = "@"
    For a = 0 To gApprCount - 1
        r = 2 + a
        ws.Cells(r, 1).Value2 = gDates(gApprD(a))
        ws.Cells(r, 2).Value = gSlotName(gApprK(a))
        ws.Cells(r, 3).Value = gStaff(gApprS(a))
        ws.Cells(r, 4).Value = gApprRules(a)
    Next a
End Sub

' ------------------------------------------------------------------
' Reloading a result from the sheets
' ------------------------------------------------------------------

' Text of a cell for comparison: trimmed text, "" for empty or error cells.
Private Function CellText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    CellText = CellToText(ws.Cells(r, c).Value2, False)
End Function

' Serial of a date cell (number, or yyyy-mm-dd text), or -1 when it is neither.
Private Function CellSerial(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Long
    Dim v As Variant
    Dim dbl As Double
    Dim t As String

    CellSerial = -1
    v = ws.Cells(r, c).Value2
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Then
        dbl = CDbl(v)
        If dbl >= 0 And dbl <= 2958465 Then CellSerial = CLng(Int(dbl))
        Exit Function
    End If
    t = Trim$(CStr(v))
    If IsDateText(t) Then CellSerial = SerialFromText(t)
End Function

Private Function DateIndex(ByVal serial As Long) As Long
    Dim d As Long

    DateIndex = -1
    For d = 0 To gDateCount - 1
        If gDates(d) = serial Then
            DateIndex = d
            Exit Function
        End If
    Next d
End Function

Private Function SlotIndex(ByVal t As String) As Long
    Dim k As Long
    Dim key As String

    SlotIndex = -1
    key = Norm(t)
    For k = 0 To gSlotCount - 1
        If StrComp(Norm(gSlotName(k)), key, vbBinaryCompare) = 0 Then
            SlotIndex = k
            Exit Function
        End If
    Next k
End Function

Private Function CodeFromRuleText(ByVal t As String) As Long
    Dim key As String

    key = Norm(t)
    If key = Norm(RuleText(GAP_SENIOR)) Then
        CodeFromRuleText = GAP_SENIOR
    ElseIf key = Norm(RuleText(GAP_COAG)) Then
        CodeFromRuleText = GAP_COAG
    Else
        CodeFromRuleText = GAP
    End If
End Function

Private Function GroupFromText(ByVal t As String) As Long
    Dim key As String

    key = Norm(t)
    If key = Norm(GroupName(1)) Then
        GroupFromText = 1
    Else
        GroupFromText = 0
    End If
End Function

' Rebuilds the result state from the Roster, Approvals and Review sheets.
' BuildInputs must have run. Returns "" on success, else a message.
Public Function LoadResultFromSheets() As String
    Dim wsR As Worksheet
    Dim wsV As Worksheet
    Dim wsA As Worksheet
    Dim d As Long
    Dim k As Long
    Dim s As Long
    Dim r As Long
    Dim t As String
    Dim lastRow As Long
    Dim n As Long
    Dim a As Long
    Dim ser As Long
    Dim cnt As Long
    Dim blank As Boolean
    Dim revN As Long
    Dim revDate() As Long
    Dim revGroup() As String
    Dim revSlot() As String
    Dim revRule() As String
    Dim code As Long
    Dim found As Boolean
    Dim i As Long
    Dim kk As Long
    Dim ss As Long

    LoadResultFromSheets = ""
    If gDateCount < 1 Or gSlotCount < 1 Then
        LoadResultFromSheets = "The roster has no dates or slots. Run Generate roster again."
        Exit Function
    End If

    Set wsR = ThisWorkbook.Worksheets("Roster")
    Set wsV = ThisWorkbook.Worksheets("Review")
    Set wsA = ThisWorkbook.Worksheets("Approvals")

    ' ---- check Roster headers and dates ----
    For k = 0 To gSlotCount - 1
        If StrComp(Norm(CellText(wsR, 1, 3 + k)), Norm(gSlotName(k)), vbBinaryCompare) <> 0 Then
            LoadResultFromSheets = "The Roster slot columns do not match Settings. Run Generate roster again."
            Exit Function
        End If
    Next k
    If CellText(wsR, 1, 3 + gSlotCount) <> "" Then
        LoadResultFromSheets = "The Roster slot columns do not match Settings. Run Generate roster again."
        Exit Function
    End If
    For d = 0 To gDateCount - 1
        If CellSerial(wsR, 2 + d, 2) <> gDates(d) Then
            LoadResultFromSheets = "The Roster dates do not match the Holidays sheet. Run Generate roster again."
            Exit Function
        End If
    Next d
    If CellText(wsR, 2 + gDateCount, 1) <> "" Or CellText(wsR, 2 + gDateCount, 2) <> "" Then
        LoadResultFromSheets = "The Roster dates do not match the Holidays sheet. Run Generate roster again."
        Exit Function
    End If

    ' ---- read Review rows (date, group, slot, rule) ----
    lastRow = LastRowOf(wsV, 4)
    n = lastRow - 1
    If n < 1 Then n = 1
    ReDim revDate(1 To n)
    ReDim revGroup(1 To n)
    ReDim revSlot(1 To n)
    ReDim revRule(1 To n)
    revN = 0
    For r = 2 To lastRow
        blank = True
        For i = 1 To 4
            If CellText(wsV, r, i) <> "" Then blank = False
        Next i
        If Not blank Then
            revN = revN + 1
            revDate(revN) = CellSerial(wsV, r, 1)
            revGroup(revN) = CellText(wsV, r, 2)
            revSlot(revN) = CellText(wsV, r, 3)
            revRule(revN) = CellText(wsV, r, 4)
        End If
    Next r

    ' ---- size the result arrays ----
    ReDim gGrid(0 To gDateCount - 1, 0 To gSlotCount - 1)
    ReDim gApproved(0 To gDateCount - 1, 0 To gSlotCount - 1)
    ReDim gCutShort(0 To gDateCount - 1)
    ReDim gNodes(0 To gDateCount - 1)
    ReDim gUnmetCount(0 To gDateCount - 1)
    ReDim gUnmetGroup(0 To gDateCount - 1, 0 To 3)
    ReDim gUnmetCode(0 To gDateCount - 1, 0 To 3)
    If gStaffCount > 0 Then
        ReDim gActiveCount(0 To gStaffCount - 1)
        ReDim gStandbyCount(0 To gStaffCount - 1)
        ReDim gEligible(0 To gStaffCount - 1)
    Else
        Erase gActiveCount
        Erase gStandbyCount
        Erase gEligible
    End If

    ' ---- grid ----
    For d = 0 To gDateCount - 1
        For k = 0 To gSlotCount - 1
            t = CellText(wsR, 2 + d, 3 + k)
            If t = "" Then
                gGrid(d, k) = EMPTY_CELL
            ElseIf Norm(t) = "GAP" Then
                code = GAP
                For i = 1 To revN
                    If revDate(i) = gDates(d) And revSlot(i) <> "" Then
                        If SlotIndex(revSlot(i)) = k Then
                            code = CodeFromRuleText(revRule(i))
                            Exit For
                        End If
                    End If
                Next i
                gGrid(d, k) = code
            Else
                s = StaffIndex(t)
                If s < 0 Then
                    LoadResultFromSheets = "Roster cell " & t & " on " & TextFromSerial(gDates(d)) & _
                        " is not on the Staff sheet. Run Generate roster again."
                    Exit Function
                End If
                gGrid(d, k) = s
            End If
        Next k
    Next d

    ' ---- unmet requirements and cut-short flags ----
    For i = 1 To revN
        If revSlot(i) = "" Then
            d = DateIndex(revDate(i))
            If d >= 0 Then
                If Norm(revRule(i)) = Norm("Search cut short") Then
                    gCutShort(d) = True
                Else
                    cnt = gUnmetCount(d)
                    If cnt < 4 Then
                        gUnmetGroup(d, cnt) = GroupFromText(revGroup(i))
                        gUnmetCode(d, cnt) = CodeFromRuleText(revRule(i))
                        gUnmetCount(d) = cnt + 1
                    End If
                End If
            End If
        End If
    Next i

    ' ---- approvals ----
    lastRow = LastRowOf(wsA, 4)
    cnt = 0
    For r = 2 To lastRow
        blank = True
        For i = 1 To 4
            If CellText(wsA, r, i) <> "" Then blank = False
        Next i
        If Not blank Then cnt = cnt + 1
    Next r
    gApprCount = cnt
    If cnt > 0 Then
        ReDim gApprD(0 To cnt - 1)
        ReDim gApprK(0 To cnt - 1)
        ReDim gApprS(0 To cnt - 1)
        ReDim gApprRules(0 To cnt - 1)
    Else
        Erase gApprD
        Erase gApprK
        Erase gApprS
        Erase gApprRules
    End If
    a = 0
    For r = 2 To lastRow
        blank = True
        For i = 1 To 4
            If CellText(wsA, r, i) <> "" Then blank = False
        Next i
        If Not blank Then
            d = DateIndex(CellSerial(wsA, r, 1))
            kk = SlotIndex(CellText(wsA, r, 2))
            ss = StaffIndex(CellText(wsA, r, 3))
            If d < 0 Or kk < 0 Or ss < 0 Then
                LoadResultFromSheets = "Approvals row " & CStr(r - 1) & " does not match the roster. Run Generate roster again."
                Exit Function
            End If
            gApprD(a) = d
            gApprK(a) = kk
            gApprS(a) = ss
            gApprRules(a) = CellText(wsA, r, 4)
            gApproved(d, kk) = True
            a = a + 1
        End If
    Next r

    ' ---- recount and eligibility ----
    For d = 0 To gDateCount - 1
        For k = 0 To gSlotCount - 1
            s = gGrid(d, k)
            If s >= 0 Then
                If gSlotActive(k) Then
                    gActiveCount(s) = gActiveCount(s) + 1
                Else
                    gStandbyCount(s) = gStandbyCount(s) + 1
                End If
            End If
        Next k
    Next d
    modRules.ComputeEligible
End Function
