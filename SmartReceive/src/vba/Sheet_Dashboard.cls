Attribute VB_Name = "Sheet1"
Option Explicit

' Sheet 1 is Dashboard
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub

    Select Case Target.Address
        Case "$B$6" ' START SCANNING
            ThisWorkbook.Sheets("Scanner UI").Select
        Case "$B$8" ' IMPORT PL
            ImportSupplierPL
        Case "$B$10" ' REPORTS
            GenerateWMSExport
    End Select
End Sub
