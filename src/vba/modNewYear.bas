Attribute VB_Name = "modNewYear"
Option Explicit

' Start a new year: current roster becomes Last year, data sheets are cleared,
' Year goes up by one. Port of ops.roster_to_lastyear plus sheet handling.

Private Sub ClearDataRows(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, ws.Columns.Count))
    rng.ClearContents
    rng.ClearComments
    rng.Interior.ColorIndex = xlNone
End Sub

' Row on the Settings sheet holding the Year key, or 0.
Private Function YearRow() As Long
    Dim ws As Worksheet
    Dim r As Long

    YearRow = 0
    Set ws = ThisWorkbook.Worksheets("Settings")
    For r = 2 To 13
        If StrComp(Norm(CStr(ws.Cells(r, 1).Value)), "YEAR", vbBinaryCompare) = 0 Then
            YearRow = r
            Exit Function
        End If
    Next r
End Function

Public Sub DoStartNewYear()
    Dim msg As String
    Dim wsL As Worksheet
    Dim wsS As Worksheet
    Dim yRow As Long
    Dim d As Long
    Dim k As Long
    Dim n As Long
    Dim r As Long

    If MsgBox("Start a new year? The current roster is saved as Last year, then Holidays, Leave, " & _
              "Roster, Review, Summary and Approvals are cleared and Year goes up by one.", _
              vbYesNo + vbQuestion, "PH Roster Planner") <> vbYes Then Exit Sub

    modIO.ReadSheetsToRaw
    msg = modValidate.ValidateRaw()
    If Len(msg) > 0 Then
        modMain.ShowMessages "Input problems", msg
        Exit Sub
    End If
    modIO.BuildInputs
    msg = modIO.LoadResultFromSheets()
    If Len(msg) > 0 Then
        modMain.ShowMessages "New year", msg
        Exit Sub
    End If
    yRow = YearRow()
    If yRow = 0 Then
        modMain.ShowMessages "New year", "Settings: missing key Year"
        Exit Sub
    End If

    ' Last year rows: filled active slots in date, slot order.
    Set wsL = ThisWorkbook.Worksheets("LastYear")
    ClearDataRows "LastYear"
    n = 0
    For d = 0 To gDateCount - 1
        For k = 0 To gSlotCount - 1
            If gGrid(d, k) >= 0 Then
                If gSlotActive(k) Then n = n + 1
            End If
        Next k
    Next d
    If n > 0 Then
        wsL.Range(wsL.Cells(2, 1), wsL.Cells(1 + n, 1)).NumberFormat = "@"
        wsL.Range(wsL.Cells(2, 2), wsL.Cells(1 + n, 2)).NumberFormat = "yyyy-mm-dd"
        wsL.Range(wsL.Cells(2, 3), wsL.Cells(1 + n, 4)).NumberFormat = "@"
        r = 1
        For d = 0 To gDateCount - 1
            For k = 0 To gSlotCount - 1
                If gGrid(d, k) >= 0 Then
                    If gSlotActive(k) Then
                        r = r + 1
                        wsL.Cells(r, 1).Value = gHoliday(d)
                        wsL.Cells(r, 2).Value2 = gDates(d)
                        wsL.Cells(r, 3).Value = gStaff(gGrid(d, k))
                        wsL.Cells(r, 4).Value = gSlotName(k)
                    End If
                End If
            Next k
        Next d
    End If

    ClearDataRows "Holidays"
    ClearDataRows "Leave"
    ClearDataRows "Roster"
    ClearDataRows "Review"
    ClearDataRows "Summary"
    ClearDataRows "Approvals"

    Set wsS = ThisWorkbook.Worksheets("Settings")
    wsS.Cells(yRow, 2).Value2 = gYear + 1
    modMain.ShowMessages "New year", "Done. Year is now " & CStr(gYear + 1) & ". Last year holds the roster just ended."
End Sub
