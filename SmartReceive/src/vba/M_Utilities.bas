Attribute VB_Name = "M_Utilities"
Option Explicit

' Optimized Table Cache
Private TableCache As Object

Public Sub LogAction(ByVal Action As String, ByVal Details As String)
    Debug.Print Now & " | " & Action & " | " & Details
End Sub

Public Sub ShowError(ByVal Msg As String)
    MsgBox Msg, vbCritical, "Error"
End Sub

Public Sub ShowInfo(ByVal Msg As String)
    MsgBox Msg, vbInformation, "SmartReceive Pro"
End Sub

Public Function GetTable(ByVal TableName As String) As ListObject
    If TableCache Is Nothing Then Set TableCache = CreateObject("Scripting.Dictionary")

    If TableCache.Exists(TableName) Then
        Set GetTable = TableCache(TableName)
        Exit Function
    End If

    Dim lo As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(TableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set TableCache(TableName) = lo
            Set GetTable = lo
            Exit Function
        End If
    Next ws
End Function

Public Function FindInTable(ByVal lo As ListObject, ByVal ColumnName As String, ByVal SearchValue As Variant) As ListRow
    Dim colIndex As Long
    On Error Resume Next
    colIndex = lo.ListColumns(ColumnName).Index
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    ' Optimized using Application.Match
    Dim res As Variant
    res = Application.Match(SearchValue, lo.ListColumns(ColumnName).Range, 0)

    If Not IsError(res) Then
        ' Match returns index including header, so we subtract 1 to get row in DataBodyRange
        If res > 1 Then
            Set FindInTable = lo.ListRows(res - 1)
        End If
    End If
End Function

Public Sub OptimizePerformance(ByVal Status As Boolean)
    With Application
        .ScreenUpdating = Not Status
        .Calculation = IIf(Status, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not Status
    End With
End Sub
