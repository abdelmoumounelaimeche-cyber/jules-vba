Attribute VB_Name = "Mod_Settings"
Option Explicit

Public Function GetSetting(ByVal key As String) As String
    Dim tbl As ListObject
    Dim foundRange As Range
    On Error Resume Next
    Set tbl = Sheets(SH_SETTINGS).ListObjects(TBL_SETTINGS)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    Set foundRange = tbl.ListColumns("Key").DataBodyRange.Find(key, LookAt:=xlWhole)
    If Not foundRange Is Nothing Then
        GetSetting = foundRange.Offset(0, 1).Value
    End If
End Function

Public Sub SetSetting(ByVal key As String, ByVal value As String)
    Dim tbl As ListObject: Set tbl = Sheets(SH_SETTINGS).ListObjects(TBL_SETTINGS)
    Dim foundRange As Range
    Set foundRange = tbl.ListColumns("Key").DataBodyRange.Find(key, LookAt:=xlWhole)
    If Not foundRange Is Nothing Then
        foundRange.Offset(0, 1).Value = value
    Else
        Dim newRow As ListRow: Set newRow = tbl.ListRows.Add
        newRow.Range(1, 1).Value = key
        newRow.Range(1, 2).Value = value
    End If
End Sub
