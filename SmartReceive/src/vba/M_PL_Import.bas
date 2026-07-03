Attribute VB_Name = "M_PL_Import"
Option Explicit

' --- PL Import Automation ---

Public Sub ImportSupplierPL()
    Dim fd As Object
    Dim selectedFile As Variant

    On Error GoTo ErrHandler

    ' Select file
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    With fd
        .Title = "Select Supplier PL CSV"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    OptimizePerformance True
    LogAction "PL Import", "Started with file: " & selectedFile

    ' Import logic using QueryTable for speed and reliability
    Dim lo As ListObject: Set lo = GetTable(TBL_PL)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete

    Dim ws As Worksheet: Set ws = lo.Parent
    Dim qt As QueryTable
    Dim connStr As String: connStr = "TEXT;" & selectedFile

    Set qt = ws.QueryTables.Add(Connection:=connStr, Destination:=lo.HeaderRowRange.Offset(1, 0).Cells(1, 1))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh
    End With

    ' Clean up QueryTable but keep data
    qt.Delete

    ' Automatically calculate Whole/Partial status
    CalculateCartonTypes

    LogAction "PL Import", "Completed"
    ShowInfo "Supplier PL Imported successfully."

CleanUp:
    OptimizePerformance False
    Exit Sub

ErrHandler:
    ShowError "Error importing PL: " & Err.Description
    Resume CleanUp
End Sub

Public Sub CalculateCartonTypes()
    Dim lo As ListObject
    Set lo = GetTable(TBL_PL)
    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    Dim r As ListRow
    Dim cartonID As String
    Dim skusInCarton As Object
    Set skusInCarton = CreateObject("Scripting.Dictionary")

    ' First pass: count SKUs per carton
    Dim data As Variant
    data = lo.DataBodyRange.Value

    Dim i As Long
    Dim cartonCol As Long: cartonCol = lo.ListColumns("Carton ID").Index

    For i = 1 To UBound(data, 1)
        cartonID = data(i, cartonCol)
        If Not skusInCarton.Exists(cartonID) Then
            skusInCarton.Add cartonID, 1
        Else
            skusInCarton(cartonID) = skusInCarton(cartonID) + 1
        End If
    Next i

    ' Second pass: assign type
    Dim typeCol As Long: typeCol = lo.ListColumns("Type").Index
    For i = 1 To UBound(data, 1)
        cartonID = data(i, cartonCol)
        If skusInCarton(cartonID) > 1 Then
            data(i, typeCol) = "Partial"
        Else
            data(i, typeCol) = "Whole"
        End If
    Next i

    lo.DataBodyRange.Value = data

    LogAction "Classification", "Completed for " & lo.ListRows.Count & " lines"
End Sub
