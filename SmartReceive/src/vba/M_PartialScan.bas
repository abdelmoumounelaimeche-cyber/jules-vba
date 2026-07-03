Attribute VB_Name = "M_PartialScan"
Option Explicit

' --- Partial Scan Logic ---
' Aggregates quantities scanned for repacked boxes.

Public Sub ScanItemToBox(ByVal BoxID As String, ByVal SKU As String, ByVal Qty As Long)
    On Error GoTo ErrHandler

    Dim loScan As ListObject: Set loScan = GetTable(TBL_SCAN_PARTIAL)

    ' Check if this SKU already exists in this box scan session to aggregate
    Dim found As Boolean: found = False
    Dim r As ListRow
    For Each r In loScan.ListRows
        If r.Range(1, loScan.ListColumns("Box ID").Index).Value = BoxID And _
           r.Range(1, loScan.ListColumns("SKU").Index).Value = SKU Then

            r.Range(1, loScan.ListColumns("Scanned Qty").Index).Value = _
                r.Range(1, loScan.ListColumns("Scanned Qty").Index).Value + Qty
            r.Range(1, loScan.ListColumns("Timestamp").Index).Value = Now
            found = True
            Exit For
        End If
    Next r

    If Not found Then
        Dim newRow As ListRow: Set newRow = loScan.ListRows.Add
        With newRow
            .Range(1, loScan.ListColumns("Box ID").Index).Value = BoxID
            .Range(1, loScan.ListColumns("SKU").Index).Value = SKU
            .Range(1, loScan.ListColumns("Scanned Qty").Index).Value = Qty
            .Range(1, loScan.ListColumns("Timestamp").Index).Value = Now
            .Range(1, loScan.ListColumns("Operator").Index).Value = CurrentOperator
        End With
    End If

    Exit Sub
ErrHandler:
    ShowError "Error in Partial Scan: " & Err.Description
End Sub

Public Sub FinalizeBox(ByVal BoxID As String)
    ' Move data from tblScanPartial to tblDataAfterReception
    ' Mark box as "Closed" in tblOpenBoxes
    Dim loScan As ListObject: Set loScan = GetTable(TBL_SCAN_PARTIAL)
    Dim loAfter As ListObject: Set loAfter = GetTable(TBL_DATA_AFTER)

    Dim i As Long
    For i = loScan.ListRows.Count To 1 Step -1
        Dim r As ListRow: Set r = loScan.ListRows(i)
        If r.Range(1, loScan.ListColumns("Box ID").Index).Value = BoxID Then
            ' Add to after reception
            Dim nr As ListRow: Set nr = loAfter.ListRows.Add
            nr.Range(1, loAfter.ListColumns("Box/Carton ID").Index).Value = BoxID
            nr.Range(1, loAfter.ListColumns("SKU").Index).Value = r.Range(1, loScan.ListColumns("SKU").Index).Value
            nr.Range(1, loAfter.ListColumns("Final Qty").Index).Value = r.Range(1, loScan.ListColumns("Scanned Qty").Index).Value
            nr.Range(1, loAfter.ListColumns("Type").Index).Value = "Partial"
            nr.Range(1, loAfter.ListColumns("Reception Date").Index).Value = Now

            ' Delete from scan table
            r.Delete
        End If
    Next i

    ' Update status in Open Boxes
    Dim loOpen As ListObject: Set loOpen = GetTable(TBL_OPEN_BOXES)
    Dim rOpen As ListRow: Set rOpen = FindInTable(loOpen, "Box Name", BoxID)
    If Not rOpen Is Nothing Then
        rOpen.Range(1, loOpen.ListColumns("Status").Index).Value = "Closed"
    End If
End Sub
