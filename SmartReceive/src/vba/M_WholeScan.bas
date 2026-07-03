Attribute VB_Name = "M_WholeScan"
Option Explicit

' --- Whole Scan Logic ---
' Processes cartons classified as "Whole". Moves data to final reception table.

Public Sub ProcessWholeScan(ByVal CartonID As String)
    On Error GoTo ErrHandler
    LogAction "Whole Scan", "Scanning Carton: " & CartonID

    Dim loPL As ListObject: Set loPL = GetTable(TBL_PL)
    Dim loDataWhole As ListObject: Set loDataWhole = GetTable(TBL_DATA_WHOLE)
    Dim loAfter As ListObject: Set loAfter = GetTable(TBL_DATA_AFTER)

    ' Check if already scanned
    If Not FindInTable(loDataWhole, "Carton ID", CartonID) Is Nothing Then
        ShowError "Carton already scanned: " & CartonID
        Exit Sub
    End If

    ' Get details from PL
    Dim rPL As ListRow
    Set rPL = FindInTable(loPL, "Carton ID", CartonID)

    If rPL Is Nothing Then
        ShowError "Carton not found in PL: " & CartonID
        Exit Sub
    End If

    If rPL.Range(1, loPL.ListColumns("Type").Index).Value <> "Whole" Then
        ShowError "Carton is NOT a Whole carton. Use Partial Scan."
        Exit Sub
    End If

    ' Add to data Whole
    Dim newRow As ListRow: Set newRow = loDataWhole.ListRows.Add
    With newRow
        .Range(1, loDataWhole.ListColumns("Carton ID").Index).Value = CartonID
        .Range(1, loDataWhole.ListColumns("SKU").Index).Value = rPL.Range(1, loPL.ListColumns("SKU").Index).Value
        .Range(1, loDataWhole.ListColumns("Qty").Index).Value = rPL.Range(1, loPL.ListColumns("Expected Qty").Index).Value
        .Range(1, loDataWhole.ListColumns("Store").Index).Value = rPL.Range(1, loPL.ListColumns("Store").Index).Value
        .Range(1, loDataWhole.ListColumns("Scan Timestamp").Index).Value = Now
    End With

    ' Also update final reception table
    Set newRow = loAfter.ListRows.Add
    With newRow
        .Range(1, loAfter.ListColumns("Box/Carton ID").Index).Value = CartonID
        .Range(1, loAfter.ListColumns("SKU").Index).Value = rPL.Range(1, loPL.ListColumns("SKU").Index).Value
        .Range(1, loAfter.ListColumns("Final Qty").Index).Value = rPL.Range(1, loPL.ListColumns("Expected Qty").Index).Value
        .Range(1, loAfter.ListColumns("Store").Index).Value = rPL.Range(1, loPL.ListColumns("Store").Index).Value
        .Range(1, loAfter.ListColumns("Type").Index).Value = "Whole"
        .Range(1, loAfter.ListColumns("Reception Date").Index).Value = Now
    End With

    LogAction "Whole Scan", "Carton " & CartonID & " completed."

    Exit Sub
ErrHandler:
    ShowError "Error in Whole Scan: " & Err.Description
End Sub
