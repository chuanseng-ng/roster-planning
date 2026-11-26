Attribute VB_Name = "modMain"
Option Explicit

' Button entry points. Each one hides screen updating, runs the work in a
' helper and shows a message for any runtime error.

Private Const APP_TITLE As String = "PH Roster Planner"

' Long text goes to the Review sheet (column I); short text to a MsgBox.
Public Sub ShowMessages(ByVal title As String, ByVal text As String)
    Dim ws As Worksheet
    Dim lines() As String
    Dim i As Long
    Dim n As Long

    If Len(text) <= 900 Then
        MsgBox text, vbInformation, title
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets("Review")
    ws.Range(ws.Cells(1, 9), ws.Cells(ws.Rows.Count, 9)).ClearContents
    lines = Split(text, vbLf)
    n = UBound(lines) + 1
    ws.Range(ws.Cells(1, 9), ws.Cells(1 + n, 9)).NumberFormat = "@"
    ws.Cells(1, 9).Value = "Messages"
    For i = 0 To n - 1
        ws.Cells(2 + i, 9).Value = lines(i)
    Next i
    MsgBox title & ": " & CStr(n) & " message line(s) are on the Review sheet, column I (header Messages).", _
           vbInformation, title
End Sub

' True when any of the first `width` columns has an entry below the header.
Private Function SheetHasData(ByVal sheetName As String, ByVal width As Long) As Boolean
    Dim ws As Worksheet
    Dim c As Long

    SheetHasData = False
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For c = 1 To width
        If ws.Cells(ws.Rows.Count, c).End(xlUp).Row >= 2 Then
            SheetHasData = True
            Exit Function
        End If
    Next c
End Function

Private Sub DoGenerate()
    Dim msg As String
    Dim chk As String

    modIO.ReadSheetsToRaw
    msg = modValidate.ValidateRaw()
    If Len(msg) > 0 Then
        ShowMessages "Input problems", msg
        Exit Sub
    End If
    If SheetHasData("Approvals", 4) Then
        If MsgBox("Existing approvals will be discarded. Continue?", vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub
    End If
    modIO.BuildInputs
    modEngine.NewResult
    modEngine.SolveYear 0
    modReview.BuildReview
    modIO.WriteRoster
    modReview.WriteReview
    modIO.WriteSummary
    modIO.WriteApprovals
    chk = modReview.CheckRules()
    If Len(chk) > 0 Then
        ShowMessages "Rule check", chk
    Else
        ShowMessages APP_TITLE, "Roster generated. " & CStr(modReview.gRevCount) & " item(s) to review."
    End If
End Sub

Public Sub GenerateRoster()
    Dim en As Long
    Dim ed As String

    On Error GoTo Fail
    Application.ScreenUpdating = False
    DoGenerate
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    en = Err.Number
    ed = Err.Description
    Application.ScreenUpdating = True
    MsgBox "Error " & CStr(en) & ": " & ed, vbExclamation, APP_TITLE
End Sub

Public Sub ApplyApprovals()
    Dim en As Long
    Dim ed As String

    On Error GoTo Fail
    Application.ScreenUpdating = False
    modApprovals.DoApplyApprovals
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    en = Err.Number
    ed = Err.Description
    Application.ScreenUpdating = True
    MsgBox "Error " & CStr(en) & ": " & ed, vbExclamation, APP_TITLE
End Sub

Public Sub ExpandLongWeekends()
    Dim en As Long
    Dim ed As String

    On Error GoTo Fail
    Application.ScreenUpdating = False
    modHolidays.ExpandHolidaySheet
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    en = Err.Number
    ed = Err.Description
    Application.ScreenUpdating = True
    MsgBox "Error " & CStr(en) & ": " & ed, vbExclamation, APP_TITLE
End Sub

Public Sub StartNewYear()
    Dim en As Long
    Dim ed As String

    On Error GoTo Fail
    Application.ScreenUpdating = False
    modNewYear.DoStartNewYear
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    en = Err.Number
    ed = Err.Description
    Application.ScreenUpdating = True
    MsgBox "Error " & CStr(en) & ": " & ed, vbExclamation, APP_TITLE
End Sub

Public Sub RunSelfTest()
    Dim en As Long
    Dim ed As String

    On Error GoTo Fail
    Application.ScreenUpdating = False
    modSelfTest.RunSelfTestInteractive
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    en = Err.Number
    ed = Err.Description
    Application.ScreenUpdating = True
    MsgBox "Error " & CStr(en) & ": " & ed, vbExclamation, APP_TITLE
End Sub
