Attribute VB_Name = "modScanner"
Option Explicit

' ---
' Module: modScanner
' Purpose: Handles Barcode Scan events and Validations
' ---

Public Sub ProcessOpenBox(barcode As String, condition As String, comments As String)
    ' Validation: Missing/Damaged require comments
    If (condition = "MISSING" Or condition = "DAMAGED") And Trim(comments) = "" Then
        modUtils.ShowError "Comments are mandatory for " & condition & " items."
        Exit Sub
    End If

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("Open Boxes").ListObjects("tblOpenBoxes")

    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    newRow.Range(1, 1).Value = barcode
    newRow.Range(1, 4).Value = condition
    newRow.Range(1, 5).Value = comments
End Sub

Public Sub ProcessFirstScan(barcode As String)
    On Error GoTo ErrorHandler

    If Not IsValidCarton(barcode) Then
        modUtils.ShowError "Unknown Carton: " & barcode
        Exit Sub
    End If

    If IsDuplicateScan(barcode, "tblFirstScan") Then
        modUtils.ShowError "Duplicate Scan detected for: " & barcode
        Exit Sub
    End If

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("First Scan").ListObjects("tblFirstScan")

    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    newRow.Range(1, 1).Value = Now
    newRow.Range(1, 2).Value = barcode
    newRow.Range(1, 3).Value = "RECEIVED"

    Exit Sub

ErrorHandler:
    modUtils.HandleError "ProcessFirstScan"
End Sub

Public Sub ProcessWholeScan(barcode As String)
    On Error GoTo ErrorHandler

    If Not IsWholeCarton(barcode) Then
        modUtils.ShowError "Not a Whole Carton or already processed: " & barcode
        Exit Sub
    End If

    If IsDuplicateScan(barcode, "tblScanWhole") Then
        modUtils.ShowError "Carton already palletized: " & barcode
        Exit Sub
    End If

    Dim store As String, pallet As String
    store = GetCartonStore(barcode)
    pallet = GetCurrentPallet(store)

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("SCAN WHOLE").ListObjects("tblScanWhole")

    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    newRow.Range(1, 1).Value = Now
    newRow.Range(1, 2).Value = barcode
    newRow.Range(1, 3).Value = store
    newRow.Range(1, 4).Value = pallet

    modUI.ShowAssignment store, pallet

    Exit Sub

ErrorHandler:
    modUtils.HandleError "ProcessWholeScan"
End Sub

Public Sub ProcessPartialScan(repackedBox As String, skuBarcode As String)
    On Error GoTo ErrorHandler

    ' Validation: Check if SKU exists in PL and not exceeding quantity
    If Not IsValidSKU(skuBarcode) Then
        modUtils.ShowError "Unknown SKU: " & skuBarcode
        Exit Sub
    End If

    Dim totalExpected As Long, totalScanned As Long
    totalExpected = GetExpectedSKUQty(skuBarcode)
    totalScanned = GetScannedSKUQty(skuBarcode)

    If totalScanned >= totalExpected Then
        modUtils.ShowError "Excess Quantity Rejected! SKU " & skuBarcode & " already fully received (" & totalExpected & " expected)."
        Exit Sub
    End If

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("Scan Partial").ListObjects("tblScanPartial")

    Dim row As ListRow
    Dim found As Boolean
    found = False

    If Not tbl.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To tbl.ListRows.Count
            If tbl.DataBodyRange(i, 2).Value = repackedBox And tbl.DataBodyRange(i, 4).Value = skuBarcode Then
                tbl.DataBodyRange(i, 5).Value = tbl.DataBodyRange(i, 5).Value + 1
                found = True
                Exit For
            End If
        Next i
    End If

    If Not found Then
        Set row = tbl.ListRows.Add
        row.Range(1, 1).Value = Now
        row.Range(1, 2).Value = repackedBox
        row.Range(1, 4).Value = skuBarcode
        row.Range(1, 5).Value = 1
    End If

    Exit Sub

ErrorHandler:
    modUtils.HandleError "ProcessPartialScan"
End Sub

' Helper Functions

Private Function IsValidCarton(barcode As String) As Boolean
    IsValidCarton = modUtils.ValueExistsInTable("tblOriginalPL", "Carton_ID", barcode)
End Function

Private Function IsDuplicateScan(barcode As String, tableName As String) As Boolean
    IsDuplicateScan = modUtils.ValueExistsInTable(tableName, "Carton_ID", barcode)
End Function

Private Function IsWholeCarton(barcode As String) As Boolean
    IsWholeCarton = modUtils.ValueExistsInTable("tblDataWhole", "Carton_ID", barcode)
End Function

Private Function IsValidSKU(sku As String) As Boolean
    IsValidSKU = modUtils.ValueExistsInTable("tblOriginalPL", "SKU", sku)
End Function

Private Function GetExpectedSKUQty(sku As String) As Long
    ' Sum of Qty for this SKU in Original PL
    Dim tbl As ListObject: Set tbl = Range("tblOriginalPL").ListObject
    Dim i As Long, total As Long: total = 0
    For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange(i, 3).Value = sku Then total = total + tbl.DataBodyRange(i, 5).Value
    Next i
    GetExpectedSKUQty = total
End Function

Private Function GetScannedSKUQty(sku As String) As Long
    ' Sum of Qty for this SKU in Scan Partial
    Dim tbl As ListObject: Set tbl = Range("tblScanPartial").ListObject
    Dim i As Long, total As Long: total = 0
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            If tbl.DataBodyRange(i, 4).Value = sku Then total = total + tbl.DataBodyRange(i, 5).Value
        Next i
    End If
    GetScannedSKUQty = total
End Function

Private Function GetCartonStore(barcode As String) As String
    GetCartonStore = modUtils.LookupValue("tblDataWhole", "Carton_ID", barcode, "Store")
End Function

Private Function GetCurrentPallet(store As String) As String
    GetCurrentPallet = modUI.GetActivePallet(store)
End Function
