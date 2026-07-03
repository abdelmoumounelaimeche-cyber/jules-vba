Attribute VB_Name = "M_Reports"
Option Explicit

' --- Reporting Logic ---

Public Sub GenerateWMSExport()
    On Error GoTo ErrHandler

    Dim loAfter As ListObject: Set loAfter = GetTable(TBL_DATA_AFTER)
    If loAfter Is Nothing Or loAfter.ListRows.Count = 0 Then
        ShowError "No data to export."
        Exit Sub
    End If

    Dim exportPath As String
    exportPath = M_Settings.GetSetting("WMS Export Path")
    If exportPath = "" Then
        exportPath = ThisWorkbook.Path & "\WMS_Export_" & Format(Now, "YYYYMMDD_HHMMSS") & ".csv"
    End If

    LogAction "Report", "Generating WMS Export: " & exportPath

    ' Actual export to CSV
    Dim fNum As Integer
    fNum = FreeFile
    Open exportPath For Output As #fNum

    ' Headers
    Dim i As Long, j As Long
    Dim line As String
    For j = 1 To loAfter.ListColumns.Count
        line = line & loAfter.ListColumns(j).Name & IIf(j < loAfter.ListColumns.Count, ",", "")
    Next j
    Print #fNum, line

    ' Data
    Dim data As Variant
    data = loAfter.DataBodyRange.Value
    For i = 1 To UBound(data, 1)
        line = ""
        For j = 1 To UBound(data, 2)
            line = line & data(i, j) & IIf(j < UBound(data, 2), ",", "")
        Next j
        Print #fNum, line
    Next i

    Close #fNum

    ShowInfo "WMS Export generated successfully at: " & exportPath

    Exit Sub
ErrHandler:
    ShowError "Error generating report: " & Err.Description
End Sub
