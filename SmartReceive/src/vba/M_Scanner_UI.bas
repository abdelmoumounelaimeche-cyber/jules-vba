Attribute VB_Name = "M_Scanner_UI"
Option Explicit

' --- Sheet-based Scanner Logic ---

Public Sub HandleScannerInput(ByVal barcode As String)
    If barcode = "" Then Exit Sub

    Dim wsUI As Worksheet: Set wsUI = ThisWorkbook.Sheets("Scanner UI")
    Dim mode As String: mode = wsUI.Range("C6").Value

    ' Capture operator from Dashboard
    CurrentOperator = ThisWorkbook.Sheets("Dashboard").Range("D4").Value
    If CurrentOperator = "" Then
        ShowError "Please enter Operator Name on Dashboard."
        Exit Sub
    End If

    wsUI.Range("C7").Value = "Processing..."

    Select Case mode
        Case "First Scan"
            M_FirstScan.ProcessFirstScan barcode
        Case "Whole Scan"
            M_WholeScan.ProcessWholeScan barcode
        Case "Segregation"
            HandleSegregation barcode
        Case "Partial Scan"
            Dim boxName As String: boxName = wsUI.Range("C8").Value
            If boxName = "" Then
                ShowError "Please scan/enter Box Name for Partial Scan."
            Else
                M_PartialScan.ScanItemToBox boxName, barcode, 1
            End If
    End Select

    ' Reset UI
    Application.EnableEvents = False
    wsUI.Range("C4").Value = ""
    wsUI.Range("C7").Value = "Ready"
    Application.EnableEvents = True
End Sub

Private Sub HandleSegregation(ByVal barcode As String)
    ' Workflow: Scan Carton -> Scan Item -> Show Suggestion
    If ActiveSourceCarton = "" Then
        If GetCartonType(barcode) = "Partial" Then
            ActiveSourceCarton = barcode
            ShowInfo "Carton " & barcode & " active. Now scan items."
        Else
            ShowError "Invalid carton for segregation: " & barcode
        End If
    Else
        ' It is an item
        Dim suggestedBox As String
        suggestedBox = M_Segregation.GetSuggestedBox(barcode)
        M_Segregation.AssignToRepackBox ActiveSourceCarton, barcode, 1, suggestedBox
        ShowInfo "Item " & barcode & " -> " & suggestedBox
    End If
End Sub
