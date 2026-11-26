Attribute VB_Name = "modHolidays"
Option Explicit

' Long-weekend expansion on the Holidays sheet. Port of ops.expand_weekends.

Private Function IsWholeYear(ByVal t As String) As Boolean
    Dim i As Long

    IsWholeYear = False
    If Len(t) <> 4 Then Exit Function
    For i = 1 To 4
        If InStr(1, "0123456789", Mid$(t, i, 1), vbBinaryCompare) < 1 Then Exit Function
    Next i
    If CLng(t) >= 1900 Then IsWholeYear = True
End Function

Private Function InList(ByRef items() As Long, ByVal n As Long, ByVal x As Long) As Boolean
    Dim i As Long

    InList = False
    For i = 1 To n
        If items(i) = x Then
            InList = True
            Exit Function
        End If
    Next i
End Function

' Adds the Saturday/Sunday/Monday/Friday neighbours of each holiday that is not
' already an added row, then sorts the data rows by date (stable).
Public Sub ExpandHolidaySheet()
    Dim ws As Worksheet
    Dim ys As String
    Dim yr As Long
    Dim n As Long
    Dim cap As Long
    Dim i As Long
    Dim j As Long
    Dim outName() As String
    Dim outAdd() As String
    Dim outSer() As Long
    Dim outN As Long
    Dim present() As Long
    Dim presentN As Long
    Dim skipped() As Long
    Dim skippedN As Long
    Dim wd As Long
    Dim o1 As Long
    Dim o2 As Long
    Dim off As Long
    Dim cand As Long
    Dim tn As String
    Dim ta As String
    Dim ts As Long
    Dim msg As String

    modIO.ReadSheetsToRaw
    ys = ""
    If Not modValidate.FindSetting("Year", ys) Then
        modMain.ShowMessages "Long weekends", "Settings: missing key Year"
        Exit Sub
    End If
    ys = Trim$(ys)
    If Not IsWholeYear(ys) Then
        modMain.ShowMessages "Long weekends", "Settings: Year must be a whole number"
        Exit Sub
    End If
    yr = CLng(ys)
    n = gRawHolidaysRows
    If n < 1 Then
        modMain.ShowMessages "Long weekends", "The Holidays sheet has no rows."
        Exit Sub
    End If
    For i = 1 To n
        If Not IsDateText(gRawHolidays(i, 2)) Then
            modMain.ShowMessages "Long weekends", "Holidays: invalid date for " & gRawHolidays(i, 1) & _
                ". Dates must be yyyy-mm-dd."
            Exit Sub
        End If
    Next i

    cap = n * 3
    ReDim outName(1 To cap)
    ReDim outAdd(1 To cap)
    ReDim outSer(1 To cap)
    ReDim present(1 To cap)
    ReDim skipped(1 To cap)
    For i = 1 To n
        outName(i) = Trim$(gRawHolidays(i, 1))
        outSer(i) = SerialFromText(gRawHolidays(i, 2))
        outAdd(i) = Trim$(gRawHolidays(i, 3))
        present(i) = outSer(i)
    Next i
    outN = n
    presentN = n
    skippedN = 0

    For i = 1 To n
        If Norm(gRawHolidays(i, 3)) <> "Y" Then
            wd = Weekday(CDate(outSer(i)), vbMonday)
            o1 = 0
            o2 = 0
            Select Case wd
                Case 1
                    o1 = -2
                    o2 = -1
                Case 5, 6
                    o1 = 1
                    o2 = 2
                Case 7
                    o1 = -1
                    o2 = 1
                Case Else
                    o1 = 0
                    o2 = 0
            End Select
            If o1 <> 0 Then
                For j = 1 To 2
                    If j = 1 Then
                        off = o1
                    Else
                        off = o2
                    End If
                    cand = outSer(i) + off
                    If Not InList(present, presentN, cand) Then
                        If Not InList(skipped, skippedN, cand) Then
                            If Year(CDate(cand)) <> yr Then
                                skippedN = skippedN + 1
                                skipped(skippedN) = cand
                            Else
                                presentN = presentN + 1
                                present(presentN) = cand
                                outN = outN + 1
                                outName(outN) = outName(i)
                                outSer(outN) = cand
                                outAdd(outN) = "Y"
                            End If
                        End If
                    End If
                Next j
            End If
        End If
    Next i

    ' stable insertion sort by date
    For i = 2 To outN
        tn = outName(i)
        ts = outSer(i)
        ta = outAdd(i)
        j = i - 1
        Do While j >= 1
            If outSer(j) > ts Then
                outName(j + 1) = outName(j)
                outSer(j + 1) = outSer(j)
                outAdd(j + 1) = outAdd(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        outName(j + 1) = tn
        outSer(j + 1) = ts
        outAdd(j + 1) = ta
    Next i

    Set ws = ThisWorkbook.Worksheets("Holidays")
    With ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 3))
        .ClearContents
        .ClearComments
        .Interior.ColorIndex = xlNone
    End With
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + outN, 1)).NumberFormat = "@"
    ws.Range(ws.Cells(2, 2), ws.Cells(1 + outN, 2)).NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(2, 3), ws.Cells(1 + outN, 3)).NumberFormat = "@"
    For i = 1 To outN
        ws.Cells(1 + i, 1).Value = outName(i)
        ws.Cells(1 + i, 2).Value2 = outSer(i)
        ws.Cells(1 + i, 3).Value = outAdd(i)
    Next i

    msg = "Added " & CStr(outN - n) & " row(s)."
    If skippedN > 0 Then
        msg = msg & vbLf & "Skipped (outside " & CStr(yr) & "): "
        For i = 1 To skippedN
            If i > 1 Then msg = msg & ", "
            msg = msg & TextFromSerial(skipped(i))
        Next i
    End If
    modMain.ShowMessages "Long weekends", msg
End Sub
