Attribute VB_Name = "M_Main"
Option Explicit

' --- GLOBALS ---
Public CurrentOperator As String
Public ActiveSourceCarton As String

' --- UI HELPERS ---
Public Sub ShowInfo(ByVal msg As String)
    MsgBox msg, vbInformation, "SmartReceive Pro"
End Sub

Public Sub ShowError(ByVal msg As String)
    MsgBox msg, vbCritical, "Error"
End Sub

Public Sub OptimizePerformance(ByVal status As Boolean)
    With Application
        .ScreenUpdating = Not status
        .Calculation = IIf(status, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not status
    End With
End Sub

' --- TABLE UTILITIES ---
Public Function GetTable(ByVal TableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(TableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set GetTable = lo
            Exit Function
        End If
    Next ws
End Function

Public Function FindInTable(ByVal lo As ListObject, ByVal ColumnName As String, ByVal SearchValue As Variant) As ListRow
    Dim res As Variant
    On Error Resume Next
    res = Application.Match(SearchValue, lo.ListColumns(ColumnName).Range, 0)
    On Error GoTo 0
    If Not IsError(res) Then
        If res > 1 Then Set FindInTable = lo.ListRows(res - 1)
    End If
End Function

' --- BUSINESS LOGIC ---
Public Sub ImportSupplierPL()
    Dim fd As Object: Set fd = Application.FileDialog(3)
    Dim selectedFile As Variant
    With fd
        .Title = "Select PL CSV"
        .Filters.Add "CSV Files", "*.csv"
        If .Show = -1 Then selectedFile = .SelectedItems(1) Else Exit Sub
    End With

    OptimizePerformance True
    Dim lo As ListObject: Set lo = GetTable("tblOriginalPL")
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete

    On Error GoTo ImportErr
    With lo.Parent.QueryTables.Add(Connection:="TEXT;" & selectedFile, Destination:=lo.HeaderRowRange.Offset(1, 0).Cells(1, 1))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh False
        .Delete
    End With

    CalculateCartonTypes
    OptimizePerformance False
    ShowInfo "Import Complete"
    Exit Sub

ImportErr:
    OptimizePerformance False
    ShowError "Import Failed: " & Err.Description
End Sub

Private Sub CalculateCartonTypes()
    Dim lo As ListObject: Set lo = GetTable("tblOriginalPL")
    If lo.ListRows.Count = 0 Then Exit Sub

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim data As Variant: data = lo.DataBodyRange.Value
    Dim i As Long

    For i = 1 To UBound(data, 1)
        Dim cid As String: cid = CStr(data(i, 1))
        dict(cid) = dict(cid) + 1
    Next i

    For i = 1 To UBound(data, 1)
        data(i, 7) = IIf(dict(CStr(data(i, 1))) > 1, "Partial", "Whole")
    Next i
    lo.DataBodyRange.Value = data
End Sub

' --- SCANNER LOGIC ---
Public Sub HandleScannerInput(ByVal barcode As String)
    Dim wsUI As Worksheet: Set wsUI = ThisWorkbook.Sheets("Scanner UI")
    CurrentOperator = ThisWorkbook.Sheets("Dashboard").Range("D4").Value

    If CurrentOperator = "" Then
        ShowError "Please enter Operator Name on Dashboard"
        Exit Sub
    End If

    ' Using D6 for Mode as per Workbook Generation
    Dim mode As String: mode = wsUI.Range("D6").Value
    wsUI.Range("D8").Value = "Processing..."

    Select Case mode
        Case "First Scan": ProcessFirstScan barcode
        Case "Whole Scan": ProcessWholeScan barcode
        Case "Segregation": ProcessSegregation barcode
    End Select

    Application.EnableEvents = False
    wsUI.Range("D4").Value = ""
    wsUI.Range("D8").Value = "Ready"
    Application.EnableEvents = True
End Sub

Private Sub ProcessFirstScan(ByVal barcode As String)
    Dim loPL As ListObject: Set loPL = GetTable("tblOriginalPL")
    Dim r As ListRow: Set r = FindInTable(loPL, "Carton ID", barcode)

    If r Is Nothing Then
        ShowError "Unknown Carton: " & barcode
        Exit Sub
    End If

    Dim cType As String: cType = r.Range(1, 7).Value
    Dim loFirst As ListObject: Set loFirst = GetTable("tblFirstScan")
    With loFirst.ListRows.Add
        .Range(1, 1).Value = Now
        .Range(1, 2).Value = barcode
        .Range(1, 3).Value = CurrentOperator
        .Range(1, 4).Value = cType
    End With
    ShowInfo "Carton " & barcode & " is " & cType
End Sub

Private Sub ProcessWholeScan(ByVal barcode As String)
    Dim loPL As ListObject: Set loPL = GetTable("tblOriginalPL")
    Dim rPL As ListRow: Set rPL = FindInTable(loPL, "Carton ID", barcode)

    If rPL Is Nothing Then
        ShowError "Unknown Carton: " & barcode
        Exit Sub
    End If

    If rPL.Range(1, 7).Value <> "Whole" Then
        ShowError "Carton is not WHOLE"
        Exit Sub
    End If

    Dim loWhole As ListObject: Set loWhole = GetTable("tblDataWhole")
    With loWhole.ListRows.Add
        .Range(1, 1).Value = barcode
        .Range(1, 2).Value = rPL.Range(1, 2).Value
        .Range(1, 3).Value = rPL.Range(1, 4).Value
        .Range(1, 4).Value = rPL.Range(1, 5).Value
        .Range(1, 5).Value = Now
    End With
    ShowInfo "Whole Carton Processed"
End Sub

Private Sub ProcessSegregation(ByVal barcode As String)
    Dim loPL As ListObject: Set loPL = GetTable("tblOriginalPL")
    Dim r As ListRow: Set r = FindInTable(loPL, "SKU", barcode)

    If Not r Is Nothing Then
        Dim store As String: store = r.Range(1, 5).Value
        ShowInfo "Place Item in BOX_" & store
    Else
        ShowError "Unknown SKU: " & barcode
    End If
End Sub

' --- SECURITY ---
Public Sub ProtectProject()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect "SR2026", UserInterfaceOnly:=True
    Next ws
End Sub

Public Sub GenerateWMSExport()
    Dim lo As ListObject: Set lo = GetTable("tblDataWhole")
    If lo.ListRows.Count = 0 Then
        ShowError "No data to export"
        Exit Sub
    End If

    Dim path As String: path = ThisWorkbook.path & "\WMS_Export_" & Format(Now, "YYYYMMDD") & ".csv"
    Dim fNum As Integer: fNum = FreeFile
    On Error Resume Next
    Open path For Output As #fNum
    If Err.Number <> 0 Then
        ShowError "Could not write to file. Check path."
        Exit Sub
    End If
    On Error GoTo 0

    Dim i As Long, j As Long
    For i = 1 To lo.ListRows.Count
        Dim line As String: line = ""
        For j = 1 To lo.ListColumns.Count
            line = line & lo.DataBodyRange(i, j).Value & IIf(j < lo.ListColumns.Count, ",", "")
        Next j
        Print #fNum, line
    Next i
    Close #fNum
    ShowInfo "Exported to " & path
End Sub
