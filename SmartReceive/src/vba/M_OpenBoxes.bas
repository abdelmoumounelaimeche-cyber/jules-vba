Attribute VB_Name = "M_OpenBoxes"
Option Explicit

' --- Open Boxes Module ---
' Tracks the lifecycle of repacked boxes.

Public Sub CreateNewBox(ByVal BoxName As String)
    Dim lo As ListObject: Set lo = GetTable(TBL_OPEN_BOXES)

    If Not FindInTable(lo, "Box Name", BoxName) Is Nothing Then
        ShowError "Box name already exists: " & BoxName
        Exit Sub
    End If

    Dim nr As ListRow: Set nr = lo.ListRows.Add
    nr.Range(1, lo.ListColumns("Box Name").Index).Value = BoxName
    nr.Range(1, lo.ListColumns("Current SKU Count").Index).Value = 0
    nr.Range(1, lo.ListColumns("Status").Index).Value = "Open"
End Sub

Public Sub UpdateBoxCount(ByVal BoxName As String)
    Dim lo As ListObject: Set lo = GetTable(TBL_OPEN_BOXES)
    Dim r As ListRow: Set r = FindInTable(lo, "Box Name", BoxName)

    If Not r Is Nothing Then
        ' Count SKUs in segregation table for this box
        Dim loSeg As ListObject: Set loSeg = GetTable(TBL_SEGREGATION)
        Dim count As Long: count = 0
        Dim sr As ListRow
        For Each sr In loSeg.ListRows
            If sr.Range(1, loSeg.ListColumns("Destination Box").Index).Value = BoxName Then
                count = count + 1
            End If
        Next sr
        r.Range(1, lo.ListColumns("Current SKU Count").Index).Value = count
    End If
End Sub
