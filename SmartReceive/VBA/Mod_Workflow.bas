Attribute VB_Name = "Mod_Workflow"
Option Explicit

' Segregation Assistant: Shows SKUs for a partial carton
Public Sub ShowSegregationInfo(ByVal cartonID As String)
    Dim wsPL As Worksheet: Set wsPL = Sheets(SH_ORIGINAL_PL)
    Dim wsSeg As Worksheet: Set wsSeg = Sheets(SH_SEGREGATION)
    Dim tblPL As ListObject: Set tblPL = wsPL.ListObjects(TBL_ORIGINAL_PL)
    Dim row As Range
    Dim nextRow As Long

    wsSeg.Cells.ClearContents
    wsSeg.Range("A1:C1").Value = Array("SKU", "Qty", "Destination Store")
    wsSeg.Range("A1:C1").Font.Bold = True
    nextRow = 2

    For Each row In tblPL.DataBodyRange.Rows
        If row.Cells(1, 1).Value = cartonID Then
            wsSeg.Cells(nextRow, 1).Value = row.Cells(1, 2).Value ' SKU
            wsSeg.Cells(nextRow, 2).Value = row.Cells(1, 4).Value ' Qty
            wsSeg.Cells(nextRow, 3).Value = row.Cells(1, 5).Value ' Store
            nextRow = nextRow + 1
        End If
    Next row

    wsSeg.Activate
    wsSeg.Columns("A:C").AutoFit
End Sub

' Open Boxes Tracking
Public Sub MarkBoxAsOpened(ByVal cartonID As String)
    Dim ws As Worksheet: Set ws = Sheets(SH_OPEN_BOXES)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(lastRow, 1).Value = cartonID
    ws.Cells(lastRow, 2).Value = Now
    ws.Cells(lastRow, 3).Value = STATUS_IN_PROGRESS
    ShowSegregationInfo cartonID
End Sub

' Discrepancy Reporting
Public Sub GenerateReceptionReport()
    Dim wsPL As Worksheet: Set wsPL = Sheets(SH_ORIGINAL_PL)
    Dim wsRec As Worksheet: Set wsRec = Sheets(SH_DATA_RECEPTION)
    Dim wsMiss As Worksheet: Set wsMiss = Sheets(SH_MISSING)
    Dim tblPL As ListObject: Set tblPL = wsPL.ListObjects(TBL_ORIGINAL_PL)
    Dim plRow As Range, nextRow As Long

    wsMiss.Cells.ClearContents
    wsMiss.Range("A1:E1").Value = Array("CartonID", "SKU", "Expected", "Actual", "Discrepancy")
    wsMiss.Range("A1:E1").Font.Bold = True
    nextRow = 2

    For Each plRow In tblPL.DataBodyRange.Rows
        Dim cartonID As String, sku As String, expected As Double, actual As Double
        cartonID = plRow.Cells(1, 1).Value
        sku = plRow.Cells(1, 2).Value
        expected = plRow.Cells(1, 4).Value

        ' Use WorksheetFunction.SumIfs for actual quantity from reception
        actual = Application.WorksheetFunction.SumIfs(wsRec.Columns("D"), wsRec.Columns("A"), cartonID, wsRec.Columns("B"), sku)

        If actual <> expected Then
            wsMiss.Cells(nextRow, 1).Value = cartonID
            wsMiss.Cells(nextRow, 2).Value = sku
            wsMiss.Cells(nextRow, 3).Value = expected
            wsMiss.Cells(nextRow, 4).Value = actual
            wsMiss.Cells(nextRow, 5).Value = actual - expected
            nextRow = nextRow + 1
        End If
    Next plRow

    RefreshDashboard
    MsgBox "Reception report generated. Found " & (nextRow - 2) & " discrepancies.", vbInformation
End Sub

Public Sub ExportToWMS()
    Dim wsRec As Worksheet: Set wsRec = Sheets(SH_DATA_RECEPTION)
    Dim filePath As String: filePath = ThisWorkbook.Path & "\WMS_Export_" & Format(Now, "YYYYMMDD_HHMMSS") & ".csv"
    Dim fNum As Integer: fNum = FreeFile
    Dim row As Long

    Open filePath For Output As #fNum
    Print #fNum, "CartonID,SKU,Qty,Store,Timestamp"
    For row = 2 To wsRec.Cells(wsRec.Rows.Count, 1).End(xlUp).Row
        Print #fNum, wsRec.Cells(row, 1).Value & "," & _
                     wsRec.Cells(row, 2).Value & "," & _
                     wsRec.Cells(row, 4).Value & "," & _
                     wsRec.Cells(row, 5).Value & "," & _
                     wsRec.Cells(row, 8).Value
    Next row
    Close #fNum
    MsgBox "WMS Export created: " & filePath, vbInformation
End Sub

Public Sub RefreshDashboard()
    Dim wsDash As Worksheet: Set wsDash = Sheets(SH_DASHBOARD)
    wsDash.Range("B2").Value = "Total Cartons"
    wsDash.Range("C2").Value = Sheets(SH_ORIGINAL_PL).ListObjects(TBL_ORIGINAL_PL).ListRows.Count
    wsDash.Range("B3").Value = "Processed"
    wsDash.Range("C3").Value = Sheets(SH_DATA_RECEPTION).Cells(Rows.Count, 1).End(xlUp).Row - 1
End Sub
