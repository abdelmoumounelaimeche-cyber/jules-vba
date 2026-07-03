Attribute VB_Name = "Mod_UI"
Option Explicit

' To be placed in Sheet code for First Scan, Whole Scan, and Partial Scan
' -----------------------------------------------------------------------
' Private Sub Worksheet_Change(ByVal Target As Range)
'    If Not Intersect(Target, Range("B2")) Is Nothing Then
'        Application.EnableEvents = False
'        Mod_Scanner.HandleScan Target.Value
'        Target.ClearContents
'        Target.Select
'        Application.EnableEvents = True
'    End If
' End Sub
' -----------------------------------------------------------------------

Public Sub SetupScanSheet(ByVal wsName As String)
    Dim ws As Worksheet
    Set ws = Sheets(wsName)
    ws.Cells.Clear

    ' Header
    With ws.Range("A1:D1")
        .Value = Array("SCANNER INPUT", "", "OPERATOR:", Application.UserName)
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Scan Cell
    With ws.Range("B2")
        .Interior.Color = vbYellow
        .Borders.LineStyle = xlContinuous
        .Font.Size = 24
        .HorizontalAlignment = xlCenter
    End With

    ws.Range("A2").Value = "SCAN HERE >>"
    ws.Range("A2").Font.Size = 14

    ' Log Header
    ws.Range("A5:D5").Value = Array("Barcode", "Timestamp", "Result", "Operator")
    ws.Range("A5:D5").Font.Bold = True
    ws.Range("A5:D5").Interior.Color = RGB(200, 200, 200)

    If wsName = SH_SCAN_PARTIAL Then
        CreateButton ws, "FINALIZE CARTON", "FinalizePartialCarton", 400, 10
    End If

    ws.Columns("A:D").AutoFit
    ws.Columns("B").ColumnWidth = 30
End Sub

Public Sub SetupDashboard()
    Dim ws As Worksheet
    Set ws = Sheets(SH_DASHBOARD)
    ws.Cells.Clear

    ws.Range("B1").Value = "SmartReceive Pro - Dashboard"
    ws.Range("B1").Font.Size = 20
    ws.Range("B1").Font.Bold = True

    CreateButton ws, "Import PL", "ImportPackingList", 100, 50
    CreateButton ws, "Start First Scan", "ActivateFirstScan", 100, 110
    CreateButton ws, "Generate Report", "GenerateReceptionReport", 100, 170

    RefreshDashboard
End Sub

Public Sub SetupSegregationSheet()
    Dim ws As Worksheet
    Set ws = Sheets(SH_SEGREGATION)
    CreateButton ws, "PROCEED TO PARTIAL SCAN", "ActivatePartialScan", 400, 10
End Sub

Private Sub CreateButton(ws As Worksheet, caption As String, action As String, left As Double, top As Double)
    Dim btn As Button
    Set btn = ws.Buttons.Add(left, top, 200, 40)
    btn.caption = caption
    btn.OnAction = action
    btn.Font.Size = 12
    btn.Font.Bold = True
End Sub

Public Sub ActivateFirstScan()
    Sheets(SH_FIRST_SCAN).Activate
    Range("B2").Select
End Sub

Public Sub ActivatePartialScan()
    Sheets(SH_SCAN_PARTIAL).Activate
    Range("B2").Select
End Sub
