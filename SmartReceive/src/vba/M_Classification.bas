Attribute VB_Name = "M_Classification"
Option Explicit

' This module handles specific logic for classifying cartons beyond simple counts if needed.

Public Sub ReclassifyAll()
    ' Force a refresh of the whole/partial status
    M_PL_Import.CalculateCartonTypes
End Sub

Public Function GetCartonType(ByVal CartonID As String) As String
    Dim lo As ListObject
    Set lo = GetTable(TBL_PL)
    Dim r As ListRow
    Set r = FindInTable(lo, "Carton ID", CartonID)

    If Not r Is Nothing Then
        GetCartonType = r.Range(1, lo.ListColumns("Type").Index).Value
    Else
        GetCartonType = "Unknown"
    End If
End Function
