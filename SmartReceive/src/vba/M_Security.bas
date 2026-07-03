Attribute VB_Name = "M_Security"
Option Explicit

' --- Security & Protection ---

Private Const PWD As String = "SR2026" ' Default password for warehouse use

Public Sub ProtectWorkbook()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Unlock data entry tables if necessary, otherwise lock everything
        ws.Protect Password:=PWD, UserInterfaceOnly:=True
    Next ws

    ThisWorkbook.Protect Password:=PWD
    LogAction "Security", "Workbook Protected"
End Sub

Public Sub UnprotectWorkbook()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=PWD
    Next ws

    ThisWorkbook.Unprotect Password:=PWD
    LogAction "Security", "Workbook Unprotected"
End Sub

Public Sub HideTechnicalSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case SH_DASHBOARD, SH_SCAN_WHOLE, SH_OPEN_BOXES, SH_MISSING
                ws.Visible = xlSheetVisible
            Case Else
                ws.Visible = xlSheetVeryHidden
        End Select
    Next ws
End Sub
