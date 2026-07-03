Attribute VB_Name = "Mod_Scanner"
Option Explicit

Public Sub HandleScan(ByVal barcode As String)
    If barcode = "" Then Exit Sub
    Select Case ActiveSheet.Name
        Case SH_FIRST_SCAN: ProcessFirstScan barcode
        Case SH_SCAN_WHOLE: ProcessWholeScan barcode
        Case SH_SCAN_PARTIAL: ProcessPartialScan barcode
    End Select
End Sub

Private Sub ProcessFirstScan(ByVal barcode As String)
    Dim classification As String
    classification = GetCartonClassification(barcode)
    If classification = "" Then
        MsgBox "Carton " & barcode & " not found!", vbCritical
        Exit Sub
    End If

    CurrentCartonID = barcode
    LogScan SH_FIRST_SCAN, barcode, "SUCCESS"

    If classification = CLASS_WHOLE Then
        Sheets(SH_SCAN_WHOLE).Activate
    Else
        Sheets(SH_OPEN_BOXES).Activate
        MarkBoxAsOpened barcode
    End If
End Sub

Private Sub ProcessWholeScan(ByVal barcode As String)
    If barcode <> CurrentCartonID Then
        MsgBox "Wrong Carton! Expected: " & CurrentCartonID, vbCritical
        Exit Sub
    End If
    TransferCartonToReception barcode, SH_DATA_RECEPTION
    LogScan SH_SCAN_WHOLE, barcode, "RECEIVED WHOLE"
    MsgBox "Carton " & barcode & " Received.", vbInformation
    Sheets(SH_FIRST_SCAN).Activate
End Sub

Private Sub ProcessPartialScan(ByVal barcode As String)
    If ValidateQuantity(barcode) Then
        AggregateSKUCount barcode
        LogScan SH_SCAN_PARTIAL, barcode, "SKU SCANNED"
    Else
        MsgBox "Quantity Limit Exceeded for SKU: " & barcode, vbCritical
    End If
End Sub

Private Function ValidateQuantity(ByVal sku As String) As Boolean
    Dim expected As Double, actual As Double
    expected = GetExpectedQty(CurrentCartonID, sku)
    actual = GetCurrentScanQty(sku)
    ValidateQuantity = (actual + 1 <= expected)
End Function

Private Function GetExpectedQty(ByVal cartonID As String, ByVal sku As String) As Double
    Dim tbl As ListObject: Set tbl = Sheets(SH_ORIGINAL_PL).ListObjects(TBL_ORIGINAL_PL)
    Dim row As ListRow
    Dim idxID As Long: idxID = GetColIndex(tbl, "CartonID")
    Dim idxSKU As Long: idxSKU = GetColIndex(tbl, "SKU")
    Dim idxQty As Long: idxQty = GetColIndex(tbl, "Qty")

    GetExpectedQty = 0
    For Each row In tbl.ListRows
        If row.Range(1, idxID).Value = cartonID And row.Range(1, idxSKU).Value = sku Then
            GetExpectedQty = row.Range(1, idxQty).Value
            Exit Function
        End If
    Next row
End Function

Private Function GetCurrentScanQty(ByVal sku As String) As Double
    Dim ws As Worksheet: Set ws = Sheets(SH_SCAN_PARTIAL)
    Dim found As Range
    Set found = ws.Columns("A").Find(sku, LookAt:=xlWhole)
    If Not found Is Nothing Then
        GetCurrentScanQty = found.Offset(0, 1).Value
    Else
        GetCurrentScanQty = 0
    End If
End Function

Public Sub LogScan(ByVal sheetName As String, ByVal barcode As String, ByVal result As String)
    Dim ws As Worksheet: Set ws = Sheets(sheetName)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(lastRow, 1).Value = barcode
    ws.Cells(lastRow, 2).Value = Now
    ws.Cells(lastRow, 3).Value = result
    ws.Cells(lastRow, 4).Value = Application.UserName
End Sub

Public Sub TransferCartonToReception(ByVal cartonID As String, ByVal destSheetName As String)
    Dim wsPL As Worksheet: Set wsPL = Sheets(SH_ORIGINAL_PL)
    Dim wsDest As Worksheet: Set wsDest = Sheets(destSheetName)
    Dim tblPL As ListObject: Set tblPL = wsPL.ListObjects(TBL_ORIGINAL_PL)
    Dim row As Range

    For Each row In tblPL.DataBodyRange.Rows
        If row.Cells(1, 1).Value = cartonID Then
            Dim lastRow As Long: lastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
            row.Copy wsDest.Range("A" & lastRow)
            wsDest.Cells(lastRow, 8).Value = Now
        End If
    Next row
End Sub

Public Sub AggregateSKUCount(ByVal sku As String)
    Dim ws As Worksheet: Set ws = Sheets(SH_SCAN_PARTIAL)
    Dim found As Range
    Set found = ws.Columns("A").Find(sku, LookAt:=xlWhole)
    If Not found Is Nothing Then
        found.Offset(0, 1).Value = found.Offset(0, 1).Value + 1
    Else
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        ws.Cells(lastRow, 1).Value = sku
        ws.Cells(lastRow, 2).Value = 1
    End If
End Sub

Public Sub FinalizePartialCarton()
    Dim wsScan As Worksheet: Set wsScan = Sheets(SH_SCAN_PARTIAL)
    Dim wsRec As Worksheet: Set wsRec = Sheets(SH_DATA_RECEPTION)
    Dim lastRowScan As Long, i As Long

    lastRowScan = wsScan.Cells(wsScan.Rows.Count, "A").End(xlUp).Row
    If lastRowScan < 2 Then Exit Sub

    For i = 2 To lastRowScan
        Dim lastRowRec As Long: lastRowRec = wsRec.Cells(wsRec.Rows.Count, "A").End(xlUp).Row + 1
        wsRec.Cells(lastRowRec, 1).Value = CurrentCartonID
        wsRec.Cells(lastRowRec, 2).Value = wsScan.Cells(i, 1).Value ' SKU
        wsRec.Cells(lastRowRec, 4).Value = wsScan.Cells(i, 2).Value ' Qty
        wsRec.Cells(lastRowRec, 8).Value = Now
    Next i

    wsScan.Range("A2:B" & lastRowScan).ClearContents
    MsgBox "Carton " & CurrentCartonID & " finalized.", vbInformation
    CurrentCartonID = ""
    Sheets(SH_FIRST_SCAN).Activate
End Sub

Public Function GetCartonClassification(ByVal cartonID As String) As String
    Dim tbl As ListObject: Set tbl = Sheets(SH_ORIGINAL_PL).ListObjects(TBL_ORIGINAL_PL)
    Dim found As Range
    Set found = tbl.ListColumns("CartonID").DataBodyRange.Find(cartonID, LookAt:=xlWhole)
    If Not found Is Nothing Then
        GetCartonClassification = found.Offset(0, tbl.ListColumns("Class").Index - 1).Value
    Else
        GetCartonClassification = ""
    End If
End Function
