Attribute VB_Name = "PensionReportingAutomation"
Option Explicit

Public Sub RefreshPensionDashboard()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ThisWorkbook.RefreshAll
    Application.CalculateFull
    Worksheets("Dashboard").Calculate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Pension dashboard refresh requested successfully.", vbInformation, "Pension Reporting"
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Refresh failed: " & Err.Description, vbExclamation, "Pension Reporting"
End Sub

Public Sub ValidatePensionHoldings()
    Dim ws As Worksheet, lastRow As Long, r As Long, issues As Long
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Worksheets("Holdings")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, "A").Value)) = "" Or Trim(CStr(ws.Cells(r, "B").Value)) = "" _
          Or Trim(CStr(ws.Cells(r, "C").Value)) = "" Or Val(ws.Cells(r, "G").Value) <= 0 _
          Or Val(ws.Cells(r, "H").Value) <= 0 Or Val(ws.Cells(r, "I").Value) <= 0 Then
            ws.Cells(r, "P").Value = "CHECK": issues = issues + 1
        Else
            ws.Cells(r, "P").Value = "OK"
        End If
    Next r
    MsgBox "Validation complete. Rows requiring review: " & issues, vbInformation, "Data Quality Check"
    Exit Sub
ErrorHandler:
    MsgBox "Validation failed: " & Err.Description, vbExclamation, "Data Quality Check"
End Sub

Public Sub ExportDashboardToPDF()
    Dim ws As Worksheet, filePath As String
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook before exporting the PDF.", vbExclamation, "Export Dashboard": Exit Sub
    End If
    filePath = ThisWorkbook.Path & Application.PathSeparator & "Pension_Dashboard.pdf"
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, Quality:=xlQualityStandard
    MsgBox "Dashboard exported to:" & vbCrLf & filePath, vbInformation, "Export Dashboard"
    Exit Sub
ErrorHandler:
    MsgBox "PDF export failed: " & Err.Description, vbExclamation, "Export Dashboard"
End Sub

Public Sub CreateMonthlySnapshot()
    Dim wsDash As Worksheet, wsLog As Worksheet, nextRow As Long
    On Error GoTo ErrorHandler
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("Snapshot_Log")
    On Error GoTo ErrorHandler
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLog.Name = "Snapshot_Log"
        wsLog.Range("A1:F1").Value = Array("Timestamp", "Reporting Date", "Market Value (€)", "Previous Value (€)", "Change in Value (€)", "Return (%)")
        wsLog.Rows(1).Font.Bold = True
    End If
    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    wsLog.Cells(nextRow, "A").Value = Now
    wsLog.Cells(nextRow, "B").Value = wsDash.Range("B4").Value
    wsLog.Cells(nextRow, "C").Value = wsDash.Range("B5").Value
    wsLog.Cells(nextRow, "D").Value = wsDash.Range("B6").Value
    wsLog.Cells(nextRow, "E").Value = wsDash.Range("B7").Value
    wsLog.Cells(nextRow, "F").Value = wsDash.Range("B8").Value
    wsLog.Cells(nextRow, "A").NumberFormat = "dd-mmm-yyyy hh:mm"
    wsLog.Cells(nextRow, "B").NumberFormat = "dd-mmm-yyyy"
    wsLog.Range("C" & nextRow & ":E" & nextRow).NumberFormat = "€#,##0.00;[Red]-€#,##0.00"
    wsLog.Cells(nextRow, "F").NumberFormat = "0.00%;[Red]-0.00%"
    MsgBox "Dashboard snapshot saved to Snapshot_Log.", vbInformation, "Monthly Snapshot"
    Exit Sub
ErrorHandler:
    MsgBox "Snapshot failed: " & Err.Description, vbExclamation, "Monthly Snapshot"
End Sub
