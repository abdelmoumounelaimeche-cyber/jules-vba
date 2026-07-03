Attribute VB_Name = "M_Segregation"
Option Explicit

' --- Segregation Assistant ---

Public Sub AssignToRepackBox(ByVal SourceCarton As String, ByVal SKU As String, ByVal Qty As Long, ByVal DestBox As String)
    On Error GoTo ErrHandler

    Dim loSeg As ListObject: Set loSeg = GetTable(TBL_SEGREGATION)
    Dim newRow As ListRow: Set newRow = loSeg.ListRows.Add

    With newRow
        .Range(1, loSeg.ListColumns("Source Carton").Index).Value = SourceCarton
        .Range(1, loSeg.ListColumns("SKU").Index).Value = SKU
        .Range(1, loSeg.ListColumns("Qty").Index).Value = Qty
        .Range(1, loSeg.ListColumns("Destination Box").Index).Value = DestBox
        .Range(1, loSeg.ListColumns("Status").Index).Value = "Segregated"
        .Range(1, loSeg.ListColumns("Operator").Index).Value = CurrentOperator
    End With

    ' Update Open Boxes status
    M_OpenBoxes.UpdateBoxCount DestBox

    Exit Sub
ErrHandler:
    ShowError "Error in Segregation: " & Err.Description
End Sub

Public Function GetSuggestedBox(ByVal SKU As String) As String
    ' Logic to suggest a box based on Store of the SKU from PL
    Dim loPL As ListObject: Set loPL = GetTable(TBL_PL)
    Dim r As ListRow: Set r = FindInTable(loPL, "SKU", SKU)

    Dim storeName As String
    If Not r Is Nothing Then
        storeName = r.Range(1, loPL.ListColumns("Store").Index).Value
        GetSuggestedBox = "BOX_" & storeName
    Else
        GetSuggestedBox = "BOX_UNKNOWN"
    End If
End Function
