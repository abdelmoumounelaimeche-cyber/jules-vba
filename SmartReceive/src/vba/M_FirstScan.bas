Attribute VB_Name = "M_FirstScan"
Option Explicit

' --- First Scan Logic ---
' Validates if carton belongs to shipment and determines if it's Whole or Partial.

Public Sub ProcessFirstScan(ByVal CartonID As String)
    On Error GoTo ErrHandler
    LogAction "First Scan", "Scanning Carton: " & CartonID

    Dim loPL As ListObject
    Set loPL = GetTable(TBL_PL)

    Dim rPL As ListRow
    Set rPL = FindInTable(loPL, "Carton ID", CartonID)

    If rPL Is Nothing Then
        ShowError "Unknown Carton: " & CartonID
        Exit Sub
    End If

    Dim cType As String
    cType = rPL.Range(1, loPL.ListColumns("Type").Index).Value

    ' Record the scan
    Dim loFirst As ListObject
    Set loFirst = GetTable(TBL_FIRST_SCAN)
    Dim newRow As ListRow
    Set newRow = loFirst.ListRows.Add

    With newRow
        .Range(1, loFirst.ListColumns("Timestamp").Index).Value = Now
        .Range(1, loFirst.ListColumns("Carton ID").Index).Value = CartonID
        .Range(1, loFirst.ListColumns("Operator").Index).Value = CurrentOperator
        .Range(1, loFirst.ListColumns("Classification").Index).Value = cType
    End With

    If cType = "Whole" Then
        ' ShowInfo "Whole Carton - Proceed to Whole Scan"
        ' In UI, this would trigger a navigation to Whole Scan screen
    Else
        ' ShowInfo "Partial Carton - Proceed to Open Box / Segregation"
    End If

    Exit Sub

ErrHandler:
    ShowError "Error in First Scan: " & Err.Description
End Sub
