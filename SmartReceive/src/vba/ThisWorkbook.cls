' --- ThisWorkbook Module ---
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    M_Security.ProtectWorkbook
    M_Security.HideTechnicalSheets

    ThisWorkbook.Sheets("Dashboard").Select
    ShowInfo "SmartReceive Pro Loaded. Please enter Operator name."
End Sub

Public Sub OnStartClick()
    ThisWorkbook.Sheets("Scanner UI").Select
End Sub
