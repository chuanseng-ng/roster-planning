Attribute VB_Name = "modTypes"
Option Explicit

' Shared constants, state and helpers. Mirrors python/roster/model.py.
' Kind constants carry a K_ prefix because Day is a VBA function name.

Public Const K_DAY As Long = 0
Public Const K_EVE As Long = 1
Public Const K_NIGHT As Long = 2
Public Const K_SBY_DE As Long = 3
Public Const K_SBY_N As Long = 4

Public Const M_DAY As Long = 1
Public Const M_EVE As Long = 2
Public Const M_NIGHT As Long = 4
Public Const M_ALL As Long = 8

Public Const GAP As Long = -1
Public Const GAP_SENIOR As Long = -2
Public Const GAP_COAG As Long = -3
Public Const EMPTY_CELL As Long = -9

Public Const NODE_CAP As Long = 200000

' ---- settings and slots ----
Public gYear As Long
Public gCounts(0 To 4) As Long
Public gRuleOn(1 To 7) As Boolean
Public gSlotCount As Long
Public gSlotKind() As Long
Public gSlotGroup() As Long
Public gSlotActive() As Boolean
Public gSlotName() As String
Public gSlotMask() As Long

' ---- staff ----
Public gStaffCount As Long
Public gStaff() As String
Public gSenior() As Boolean
Public gCoag() As Boolean
Public gCannot() As Long

' ---- dates ----
Public gDateCount As Long
Public gDates() As Long
Public gHoliday() As String
Public gExempt() As Boolean
Public gLeave() As Long

' ---- result ----
Public gGrid() As Long
Public gApproved() As Boolean
Public gActiveCount() As Long
Public gStandbyCount() As Long
Public gEligible() As Long
Public gCutShort() As Boolean
Public gNodes() As Long
Public gUnmetCount() As Long
Public gUnmetGroup() As Long
Public gUnmetCode() As Long

' ---- approvals ----
Public gApprCount As Long
Public gApprD() As Long
Public gApprK() As Long
Public gApprS() As Long
Public gApprRules() As String

' ---- raw tables: (1 To rows, 1 To cols), companion row counts ----
Public gRawStaff() As String
Public gRawStaffRows As Long
Public gRawHolidays() As String
Public gRawHolidaysRows As Long
Public gRawLastYear() As String
Public gRawLastYearRows As Long
Public gRawLeave() As String
Public gRawLeaveRows As Long
Public gRawSettings() As String
Public gRawSettingsRows As Long

' Name of the i-th settings key, i = 0 to 11 (Python SETTINGS_KEYS).
Public Function SettingsKey(ByVal i As Long) As String
    Select Case i
        Case 0
            SettingsKey = "Year"
        Case 1
            SettingsKey = "Count Day"
        Case 2
            SettingsKey = "Count Evening"
        Case 3
            SettingsKey = "Count Night"
        Case 4
            SettingsKey = "Count Standby D/E"
        Case 5
            SettingsKey = "Count Standby N"
        Case 6
            SettingsKey = "Rule 1"
        Case 7
            SettingsKey = "Rule 2"
        Case 8
            SettingsKey = "Rule 4"
        Case 9
            SettingsKey = "Rule 5"
        Case 10
            SettingsKey = "Rule 6"
        Case 11
            SettingsKey = "Rule 7"
        Case Else
            SettingsKey = ""
    End Select
End Function

Public Function KindName(ByVal kind As Long) As String
    Select Case kind
        Case K_DAY
            KindName = "Day"
        Case K_EVE
            KindName = "Evening"
        Case K_NIGHT
            KindName = "Night"
        Case K_SBY_DE
            KindName = "Standby D/E"
        Case K_SBY_N
            KindName = "Standby N"
        Case Else
            KindName = ""
    End Select
End Function

' Python KIND_GROUP: 0 = Day + evening, 1 = Night.
Public Function KindGroup(ByVal kind As Long) As Long
    Select Case kind
        Case K_NIGHT, K_SBY_N
            KindGroup = 1
        Case Else
            KindGroup = 0
    End Select
End Function

' Python KIND_ACTIVE: standby kinds are not active.
Public Function KindActive(ByVal kind As Long) As Boolean
    Select Case kind
        Case K_SBY_DE, K_SBY_N
            KindActive = False
        Case Else
            KindActive = True
    End Select
End Function

' Python KIND_MASK: shifts whose leave / cannot-do block this kind.
Public Function KindMask(ByVal kind As Long) As Long
    Select Case kind
        Case K_DAY
            KindMask = M_DAY
        Case K_EVE
            KindMask = M_EVE
        Case K_NIGHT
            KindMask = M_NIGHT
        Case K_SBY_DE
            KindMask = M_DAY + M_EVE
        Case K_SBY_N
            KindMask = M_NIGHT
        Case Else
            KindMask = 0
    End Select
End Function

Public Function GroupName(ByVal g As Long) As String
    If g = 0 Then
        GroupName = "Day + evening"
    ElseIf g = 1 Then
        GroupName = "Night"
    Else
        GroupName = ""
    End If
End Function

Public Function RuleText(ByVal code As Long) As String
    Select Case code
        Case GAP
            RuleText = "No eligible staff"
        Case GAP_SENIOR
            RuleText = "No senior"
        Case GAP_COAG
            RuleText = "No special coag"
        Case Else
            RuleText = ""
    End Select
End Function

Public Function Norm(ByVal t As String) As String
    Norm = UCase$(Trim$(t))
End Function

' True when t (after trimming spaces) is yyyy-mm-dd, year 1900 or later,
' and a real calendar date. Port of validate.is_valid_date.
Public Function IsDateText(ByVal t As String) As Boolean
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim yy As Long
    Dim mm As Long
    Dim dd As Long
    Dim ds As Date

    IsDateText = False
    s = Trim$(t)
    If Len(s) <> 10 Then Exit Function
    For i = 1 To 10
        ch = Mid$(s, i, 1)
        If i = 5 Or i = 8 Then
            If ch <> "-" Then Exit Function
        Else
            If InStr(1, "0123456789", ch, vbBinaryCompare) < 1 Then Exit Function
        End If
    Next i
    yy = CLng(Left$(s, 4))
    mm = CLng(Mid$(s, 6, 2))
    dd = CLng(Mid$(s, 9, 2))
    If yy < 1900 Then Exit Function
    If mm < 1 Then Exit Function
    If mm > 12 Then Exit Function
    If dd < 1 Then Exit Function
    If dd > 31 Then Exit Function
    ds = DateSerial(yy, mm, dd)
    If Day(ds) <> dd Then Exit Function
    IsDateText = True
End Function

' yyyy-mm-dd text to a whole-number serial. The text must pass IsDateText.
Public Function SerialFromText(ByVal t As String) As Long
    Dim s As String
    s = Trim$(t)
    SerialFromText = CLng(DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Right$(s, 2))))
End Function

Public Function TextFromSerial(ByVal n As Long) As String
    TextFromSerial = Format$(CDate(n), "yyyy-mm-dd")
End Function

' Index into gStaff of the person whose Norm matches t, or -1.
Public Function StaffIndex(ByVal t As String) As Long
    Dim key As String
    Dim i As Long

    StaffIndex = -1
    key = Norm(t)
    For i = 0 To gStaffCount - 1
        If StrComp(Norm(gStaff(i)), key, vbBinaryCompare) = 0 Then
            StaffIndex = i
            Exit Function
        End If
    Next i
End Function

' Builds the slot arrays from gCounts, columns in kind order. A slot is named
' "<kind name> <i>" when its count is above 1, else "<kind name>".
Public Sub BuildSlots()
    Dim kind As Long
    Dim i As Long
    Dim total As Long
    Dim n As Long

    total = 0
    For kind = 0 To 4
        total = total + gCounts(kind)
    Next kind
    gSlotCount = total
    If total = 0 Then
        Erase gSlotKind
        Erase gSlotGroup
        Erase gSlotActive
        Erase gSlotName
        Erase gSlotMask
        Exit Sub
    End If
    ReDim gSlotKind(0 To total - 1)
    ReDim gSlotGroup(0 To total - 1)
    ReDim gSlotActive(0 To total - 1)
    ReDim gSlotName(0 To total - 1)
    ReDim gSlotMask(0 To total - 1)
    n = 0
    For kind = 0 To 4
        For i = 1 To gCounts(kind)
            gSlotKind(n) = kind
            gSlotGroup(n) = KindGroup(kind)
            gSlotActive(n) = KindActive(kind)
            gSlotMask(n) = KindMask(kind)
            If gCounts(kind) > 1 Then
                gSlotName(n) = KindName(kind) & " " & CStr(i)
            Else
                gSlotName(n) = KindName(kind)
            End If
            n = n + 1
        Next i
    Next kind
End Sub
