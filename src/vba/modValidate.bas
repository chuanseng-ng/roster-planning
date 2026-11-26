Attribute VB_Name = "modValidate"
Option Explicit

' Input validation. Port of python/roster/validate.py.
' Messages are joined with vbLf, in the same order and with the same text.

Private Sub AddMsg(ByRef msgs As String, ByVal m As String)
    If Len(msgs) = 0 Then
        msgs = m
    Else
        msgs = msgs & vbLf & m
    End If
End Sub

' Exactly 4 digits (after trimming spaces), value 1900 to 9999.
Private Function IsValidYear(ByVal t As String) As Boolean
    Dim s As String
    Dim i As Long

    IsValidYear = False
    s = Trim$(t)
    If Len(s) <> 4 Then Exit Function
    For i = 1 To 4
        If InStr(1, "0123456789", Mid$(s, i, 1), vbBinaryCompare) < 1 Then Exit Function
    Next i
    If CLng(s) >= 1900 Then IsValidYear = True
End Function

Private Function IsPlainAscii(ByVal t As String) As Boolean
    Dim i As Long
    Dim c As Long

    IsPlainAscii = True
    For i = 1 To Len(t)
        c = AscW(Mid$(t, i, 1))
        If c < 32 Or c > 126 Then
            IsPlainAscii = False
            Exit Function
        End If
    Next i
End Function

Private Function ListHas(ByRef items() As String, ByVal n As Long, ByVal t As String) As Boolean
    Dim i As Long

    ListHas = False
    For i = 1 To n
        If StrComp(items(i), t, vbBinaryCompare) = 0 Then
            ListHas = True
            Exit Function
        End If
    Next i
End Function

Private Function IsCountKey(ByVal i As Long) As Boolean
    IsCountKey = (i >= 1 And i <= 5)
End Function

Private Function IsCannotWord(ByVal w As String) As Boolean
    IsCannotWord = (StrComp(w, "DAY", vbBinaryCompare) = 0 Or _
                    StrComp(w, "EVENING", vbBinaryCompare) = 0 Or _
                    StrComp(w, "NIGHT", vbBinaryCompare) = 0)
End Function

Private Function IsShiftWord(ByVal w As String) As Boolean
    IsShiftWord = (StrComp(w, "ALL", vbBinaryCompare) = 0 Or _
                   StrComp(w, "DAY", vbBinaryCompare) = 0 Or _
                   StrComp(w, "EVENING", vbBinaryCompare) = 0 Or _
                   StrComp(w, "NIGHT", vbBinaryCompare) = 0)
End Function

' Finds the value text of the settings row for key. Returns False when the key
' is not present (Python returns None).
Public Function FindSetting(ByVal key As String, ByRef found As String) As Boolean
    Dim i As Long
    Dim k As String

    FindSetting = False
    found = ""
    k = Norm(key)
    For i = 1 To gRawSettingsRows
        If StrComp(Norm(gRawSettings(i, 1)), k, vbBinaryCompare) = 0 Then
            found = gRawSettings(i, 2)
            FindSetting = True
            Exit Function
        End If
    Next i
End Function

Public Function ValidateRaw() As String
    Dim msgs As String
    Dim yearOk As Boolean
    Dim yearVal As Long
    Dim i As Long
    Dim r As Long
    Dim key As String
    Dim v As String
    Dim t As String
    Dim nm As String
    Dim w As String
    Dim pieces() As String
    Dim p As Long
    Dim staffNames() As String
    Dim staffNameCount As Long
    Dim seen() As String
    Dim seenCount As Long
    Dim holidayNames() As String
    Dim holidayNameCount As Long
    Dim seenDates() As String
    Dim seenDateCount As Long
    Dim dateText As String
    Dim fromOk As Boolean
    Dim toOk As Boolean
    Dim n As Long

    msgs = ""

    ' Settings
    yearOk = False
    yearVal = 0
    For i = 0 To 11
        key = SettingsKey(i)
        If Not FindSetting(key, v) Then
            AddMsg msgs, "Settings: missing key " & key
        ElseIf i = 0 Then
            If IsValidYear(v) Then
                yearOk = True
                yearVal = CLng(Trim$(v))
            Else
                AddMsg msgs, "Settings: Year must be a whole number"
            End If
        ElseIf IsCountKey(i) Then
            t = Trim$(v)
            If Not (Len(t) = 1 And InStr(1, "0123456789", t, vbBinaryCompare) >= 1) Then
                AddMsg msgs, "Settings: " & key & " must be a whole number from 0 to 9"
            End If
        Else
            If StrComp(Norm(v), "Y", vbBinaryCompare) <> 0 And StrComp(Norm(v), "N", vbBinaryCompare) <> 0 Then
                AddMsg msgs, "Settings: " & key & " must be Y or N"
            End If
        End If
    Next i

    ' Staff
    n = gRawStaffRows
    If n < 1 Then n = 1
    ReDim staffNames(1 To n)
    ReDim seen(1 To n)
    staffNameCount = 0
    seenCount = 0
    For r = 1 To gRawStaffRows
        nm = Trim$(gRawStaff(r, 1))
        If nm = "" Then
            AddMsg msgs, "Staff row " & CStr(r) & ": name is blank"
        Else
            If ListHas(seen, seenCount, Norm(nm)) Then
                AddMsg msgs, "Staff row " & CStr(r) & ": duplicate name " & nm
            Else
                seenCount = seenCount + 1
                seen(seenCount) = Norm(nm)
                staffNameCount = staffNameCount + 1
                staffNames(staffNameCount) = Norm(nm)
            End If
            If Not IsPlainAscii(nm) Then
                AddMsg msgs, "Staff row " & CStr(r) & ": name must use plain ASCII characters"
            End If
        End If
        If StrComp(Norm(gRawStaff(r, 3)), "Y", vbBinaryCompare) <> 0 And StrComp(Norm(gRawStaff(r, 3)), "N", vbBinaryCompare) <> 0 Then
            AddMsg msgs, "Staff row " & CStr(r) & ": Senior must be Y or N"
        End If
        pieces = Split(gRawStaff(r, 5), ",")
        For p = LBound(pieces) To UBound(pieces)
            w = Trim$(pieces(p))
            If w <> "" Then
                If Not IsCannotWord(Norm(w)) Then
                    AddMsg msgs, "Staff row " & CStr(r) & ": unknown Cannot do value " & w
                End If
            End If
        Next p
    Next r

    ' Holidays
    n = gRawHolidaysRows
    If n < 1 Then n = 1
    ReDim holidayNames(1 To n)
    ReDim seenDates(1 To n)
    holidayNameCount = 0
    seenDateCount = 0
    For r = 1 To gRawHolidaysRows
        nm = Trim$(gRawHolidays(r, 1))
        If nm = "" Then
            AddMsg msgs, "Holidays row " & CStr(r) & ": holiday name is blank"
        Else
            holidayNameCount = holidayNameCount + 1
            holidayNames(holidayNameCount) = Norm(nm)
            If Not IsPlainAscii(nm) Then
                AddMsg msgs, "Holidays row " & CStr(r) & ": holiday name must use plain ASCII characters"
            End If
        End If
        dateText = Trim$(gRawHolidays(r, 2))
        If Not IsDateText(dateText) Then
            AddMsg msgs, "Holidays row " & CStr(r) & ": invalid date"
        Else
            If yearOk Then
                If CLng(Left$(dateText, 4)) <> yearVal Then
                    AddMsg msgs, "Holidays row " & CStr(r) & ": date is outside year " & CStr(yearVal)
                End If
            End If
            If ListHas(seenDates, seenDateCount, dateText) Then
                AddMsg msgs, "Holidays row " & CStr(r) & ": duplicate date " & dateText
            Else
                seenDateCount = seenDateCount + 1
                seenDates(seenDateCount) = dateText
            End If
        End If
    Next r

    ' LastYear
    For r = 1 To gRawLastYearRows
        If Not ListHas(staffNames, staffNameCount, Norm(gRawLastYear(r, 3))) Then
            AddMsg msgs, "LastYear row " & CStr(r) & ": unknown staff name " & Trim$(gRawLastYear(r, 3))
        End If
        If Not ListHas(holidayNames, holidayNameCount, Norm(gRawLastYear(r, 1))) Then
            AddMsg msgs, "LastYear row " & CStr(r) & ": holiday name " & Trim$(gRawLastYear(r, 1)) & " is not on the Holidays sheet"
        End If
        If Trim$(gRawLastYear(r, 4)) = "" Then
            AddMsg msgs, "LastYear row " & CStr(r) & ": Slot is blank"
        End If
    Next r

    ' Leave
    For r = 1 To gRawLeaveRows
        If Not ListHas(staffNames, staffNameCount, Norm(gRawLeave(r, 1))) Then
            AddMsg msgs, "Leave row " & CStr(r) & ": unknown staff name " & Trim$(gRawLeave(r, 1))
        End If
        fromOk = IsDateText(gRawLeave(r, 2))
        toOk = IsDateText(gRawLeave(r, 3))
        If Not fromOk Or Not toOk Then
            AddMsg msgs, "Leave row " & CStr(r) & ": invalid date"
        Else
            If StrComp(Trim$(gRawLeave(r, 2)), Trim$(gRawLeave(r, 3)), vbBinaryCompare) > 0 Then
                AddMsg msgs, "Leave row " & CStr(r) & ": From date is after To date"
            End If
        End If
        If Not IsShiftWord(Norm(gRawLeave(r, 4))) Then
            AddMsg msgs, "Leave row " & CStr(r) & ": unknown Shift " & Trim$(gRawLeave(r, 4))
        End If
    Next r

    ValidateRaw = msgs
End Function
