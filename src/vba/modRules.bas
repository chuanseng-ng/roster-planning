Attribute VB_Name = "modRules"
Option Explicit

' Eligibility rules and fairness order. Port of python/roster/rules.py and
' python/roster/fairness.py. Rule lists are comma-separated ascending, e.g. "3,5".

Private Sub AddRule(ByRef ruleList As String, ByVal r As Long)
    If Len(ruleList) = 0 Then
        ruleList = CStr(r)
    Else
        ruleList = ruleList & "," & CStr(r)
    End If
End Sub

' Rules among 3,4,5,6,7 that stop staff s taking slot k on date index d, ascending.
' The cell (d, k) itself is ignored.
Public Function BlockingRules(ByVal s As Long, ByVal d As Long, ByVal k As Long) As String
    Dim out As String
    Dim j As Long
    Dim lv As Long
    Dim mask As Long
    Dim kd As Long

    out = ""
    mask = gSlotMask(k)
    ' Rule 3: already in another cell of this date.
    For j = 0 To gSlotCount - 1
        If j <> k Then
            If gGrid(d, j) = s Then
                AddRule out, 3
                Exit For
            End If
        End If
    Next j
    ' Rule 4: exempt, for active slots only.
    If gSlotActive(k) Then
        If gExempt(s, d) Then
            AddRule out, 4
        End If
    End If
    ' Rule 5: leave.
    lv = gLeave(s, d)
    If (lv And M_ALL) <> 0 Or (lv And mask) <> 0 Then
        AddRule out, 5
    End If
    ' Rule 6: cannot-do flags.
    If (gCannot(s) And mask) <> 0 Then
        AddRule out, 6
    End If
    ' Rule 7: no day/evening work straight after a night.
    If gRuleOn(7) Then
        If d > 0 Then
            If gDates(d - 1) = gDates(d) - 1 Then
                kd = gSlotKind(k)
                If kd <> K_NIGHT And kd <> K_SBY_N Then
                    For j = 0 To gSlotCount - 1
                        If gSlotKind(j) = K_NIGHT Then
                            If gGrid(d - 1, j) = s Then
                                AddRule out, 7
                                Exit For
                            End If
                        End If
                    Next j
                End If
            End If
        End If
    End If
    BlockingRules = out
End Function

Public Function HasRule(ByVal ruleList As String, ByVal r As Long) As Boolean
    HasRule = (InStr(1, "," & ruleList & ",", "," & CStr(r) & ",", vbBinaryCompare) > 0)
End Function

' The list without rule r, order kept.
Public Function WithoutRule(ByVal ruleList As String, ByVal r As Long) As String
    Dim parts() As String
    Dim i As Long
    Dim out As String

    out = ""
    If Len(ruleList) > 0 Then
        parts = Split(ruleList, ",")
        For i = 0 To UBound(parts)
            If StrComp(parts(i), CStr(r), vbBinaryCompare) <> 0 Then
                If Len(out) = 0 Then
                    out = parts(i)
                Else
                    out = out & "," & parts(i)
                End If
            End If
        Next i
    End If
    WithoutRule = out
End Function

' True when staff s on a night slot k of date d would collide with an approved
' day/evening/standby D/E cell of s on the next calendar day (rule 7).
Public Function ForwardNightBlock(ByVal s As Long, ByVal d As Long, ByVal k As Long) As Boolean
    Dim j As Long
    Dim kd As Long

    ForwardNightBlock = False
    If Not gRuleOn(7) Then Exit Function
    If gSlotKind(k) <> K_NIGHT Then Exit Function
    If d + 1 >= gDateCount Then Exit Function
    If gDates(d + 1) <> gDates(d) + 1 Then Exit Function
    For j = 0 To gSlotCount - 1
        kd = gSlotKind(j)
        If kd <> K_NIGHT And kd <> K_SBY_N Then
            If gApproved(d + 1, j) Then
                If gGrid(d + 1, j) = s Then
                    ForwardNightBlock = True
                    Exit Function
                End If
            End If
        End If
    Next j
End Function

' Fills gEligible: per staff, the number of dates that are not exempt and not
' full-day leave. gEligible must already be sized (staff count above 0).
Public Sub ComputeEligible()
    Dim s As Long
    Dim d As Long
    Dim n As Long

    If gStaffCount < 1 Then Exit Sub
    For s = 0 To gStaffCount - 1
        n = 0
        For d = 0 To gDateCount - 1
            If Not gExempt(s, d) Then
                If (gLeave(s, d) And M_ALL) = 0 Then
                    n = n + 1
                End If
            End If
        Next d
        gEligible(s) = n
    Next s
End Sub

' Is a1/e1 strictly lower than a2/e2; requires e1, e2 > 0.
Public Function RatioLess(ByVal a1 As Long, ByVal e1 As Long, ByVal a2 As Long, ByVal e2 As Long) As Boolean
    RatioLess = (a1 * e2 < a2 * e1)
End Function

' True when staff x must be tried strictly before staff y (Python _before).
Private Function CandBefore(ByVal x As Long, ByVal y As Long) As Boolean
    Dim zx As Boolean
    Dim zy As Boolean
    Dim nx As String
    Dim ny As String

    zx = (gEligible(x) = 0)
    zy = (gEligible(y) = 0)
    If zx <> zy Then
        CandBefore = zy
        Exit Function
    End If
    If Not zx Then
        If RatioLess(gActiveCount(x), gEligible(x), gActiveCount(y), gEligible(y)) Then
            CandBefore = True
            Exit Function
        End If
        If RatioLess(gActiveCount(y), gEligible(y), gActiveCount(x), gEligible(x)) Then
            CandBefore = False
            Exit Function
        End If
    End If
    If gStandbyCount(x) <> gStandbyCount(y) Then
        CandBefore = (gStandbyCount(x) < gStandbyCount(y))
        Exit Function
    End If
    nx = Norm(gStaff(x))
    ny = Norm(gStaff(y))
    If StrComp(nx, ny, vbBinaryCompare) <> 0 Then
        CandBefore = (StrComp(nx, ny, vbBinaryCompare) < 0)
        Exit Function
    End If
    CandBefore = False
End Function

' Staff indexes in the order they are tried: staff with no eligible dates last,
' then lowest active/eligible ratio, then fewest standby, then name. Stable
' insertion sort. The array is resized here (unallocated when there is no staff).
Public Sub CandidateOrder(ByRef order() As Long)
    Dim s As Long
    Dim i As Long
    Dim j As Long
    Dim cur As Long

    If gStaffCount < 1 Then
        Erase order
        Exit Sub
    End If
    ReDim order(0 To gStaffCount - 1)
    For s = 0 To gStaffCount - 1
        order(s) = s
    Next s
    For i = 1 To gStaffCount - 1
        cur = order(i)
        j = i - 1
        Do
            If j < 0 Then Exit Do
            If Not CandBefore(cur, order(j)) Then Exit Do
            order(j + 1) = order(j)
            j = j - 1
        Loop
        order(j + 1) = cur
    Next i
End Sub
