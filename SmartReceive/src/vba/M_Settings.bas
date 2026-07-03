Attribute VB_Name = "M_Settings"
Option Explicit

' --- Settings Management ---

Public Function GetSetting(ByVal ParamName As String) As String
    Dim lo As ListObject: Set lo = GetTable(TBL_SETTINGS)
    Dim r As ListRow: Set r = FindInTable(lo, "Parameter", ParamName)

    If Not r Is Nothing Then
        GetSetting = r.Range(1, lo.ListColumns("Value").Index).Value
    Else
        GetSetting = ""
    End If
End Function

Public Sub SaveSetting(ByVal ParamName As String, ByVal Val As String)
    Dim lo As ListObject: Set lo = GetTable(TBL_SETTINGS)
    Dim r As ListRow: Set r = FindInTable(lo, "Parameter", ParamName)

    If Not r Is Nothing Then
        r.Range(1, lo.ListColumns("Value").Index).Value = Val
    Else
        Dim nr As ListRow: Set nr = lo.ListRows.Add
        nr.Range(1, lo.ListColumns("Parameter").Index).Value = ParamName
        nr.Range(1, lo.ListColumns("Value").Index).Value = Val
    End If
End Sub
