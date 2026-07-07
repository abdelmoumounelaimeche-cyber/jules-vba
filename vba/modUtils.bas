Attribute VB_Name = "modUtils"
Option Explicit

' ---
' Module: modUtils
' Purpose: Shared utilities, table operations, and error handling
' ---

Public Sub ClearTable(tbl As ListObject)
    If Not tbl.DataBodyRange Is Nothing Then
        ' ClearContents preserves formatting and formulas better than Delete
        tbl.DataBodyRange.ClearContents
        ' If we need to remove extra rows:
        If tbl.ListRows.Count > 1 Then
            Dim i As Long
            For i = tbl.ListRows.Count To 2 Step -1
                tbl.ListRows(i).Delete
            Next i
        End If
    End If
End Sub

Public Function GetTableRowCount(tableName As String) As Long
    On Error Resume Next
    GetTableRowCount = Range(tableName).ListObject.ListRows.Count
    If Err.Number <> 0 Then GetTableRowCount = 0
End Function

Public Function ValueExistsInTable(tableName As String, colName As String, val As Variant) As Boolean
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Range(tableName).ListObject
    If tbl Is Nothing Then Exit Function

    Dim colIdx As Long
    colIdx = tbl.ListColumns(colName).Index

    Dim found As Range
    Set found = tbl.ListColumns(colIdx).DataBodyRange.Find(What:=val, LookIn:=xlValues, LookAt:=xlWhole)

    ValueExistsInTable = Not found Is Nothing
End Function

' Optimized Lookup using Dictionary for performance on large datasets
Public Function GetTableDictionary(tableName As String, keyColName As String, valColName As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Range(tableName).ListObject
    If tbl Is Nothing Then Set GetTableDictionary = dict: Exit Function

    Dim data As Variant
    data = tbl.DataBodyRange.Value

    Dim kCol As Long, vCol As Long
    kCol = tbl.ListColumns(keyColName).Index
    vCol = tbl.ListColumns(valColName).Index

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Not dict.Exists(CStr(data(i, kCol))) Then
            dict.Add CStr(data(i, kCol)), data(i, vCol)
        End If
    Next i

    Set GetTableDictionary = dict
End Function

Public Function LookupValue(tableName As String, lookupCol As String, lookupVal As Variant, resultCol As String) As Variant
    Dim tbl As ListObject
    Set tbl = Range(tableName).ListObject

    Dim lCol As Long, rCol As Long
    lCol = tbl.ListColumns(lookupCol).Index
    rCol = tbl.ListColumns(resultCol).Index

    ' Linear search for small lookups/utility
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.DataBodyRange(i, lCol).Value) = CStr(lookupVal) Then
            LookupValue = tbl.DataBodyRange(i, rCol).Value
            Exit Function
        End If
    Next i
    LookupValue = ""
End Function

Public Sub UpdateSetting(sName As String, sValue As String)
    Dim tbl As ListObject
    Set tbl = Range("tblSettings").ListObject

    Dim i As Long
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            If tbl.DataBodyRange(i, 1).Value = sName Then
                tbl.DataBodyRange(i, 2).Value = sValue
                Exit Sub
            End If
        Next i
    End If

    Dim row As ListRow
    Set row = tbl.ListRows.Add
    row.Range(1, 1).Value = sName
    row.Range(1, 2).Value = sValue
End Sub

Public Sub ShowError(msg As String)
    MsgBox msg, vbCritical, "SmartReceive Error"
End Sub

Public Sub HandleError(procName As String)
    MsgBox "An error occurred in " & procName & ": " & Err.Description, vbCritical
End Sub
