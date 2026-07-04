Attribute VB_Name = "ThisWorkbook"
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    ProtectProject
    ThisWorkbook.Sheets("Dashboard").Select
    MsgBox "SmartReceive Pro Ready. Enter Operator Name.", vbInformation
End Sub
