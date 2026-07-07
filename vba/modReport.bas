Attribute VB_Name = "modReport"
Option Explicit

' ---
' Module: modReport
' Purpose: Report Generation and WMS Export
' ---

Public Sub GenerateWMSExport()
    On Error GoTo ErrorHandler

    Dim fileName As String
    fileName = ThisWorkbook.Path & "\WMS_Export_" & Format(Now, "YYYYMMDD_HHMMSS") & ".csv"

    Dim fNum As Integer
    fNum = FreeFile

    Open fileName For Output As #fNum

    ' Header
    Print #fNum, "Store,SKU,Qty,Type,Ref"

    ' Optimize using Dictionaries for lookups
    Dim dictWholeData As Object, dictWholeStore As Object
    Set dictWholeData = modUtils.GetTableDictionary("tblDataWhole", "Carton_ID", "SKU")
    Set dictWholeStore = modUtils.GetTableDictionary("tblDataWhole", "Carton_ID", "Store")

    ' 1. Export Whole Scans
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("SCAN WHOLE").ListObjects("tblScanWhole")

    If Not tbl.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To tbl.ListRows.Count
            Dim cID As String
            cID = CStr(tbl.DataBodyRange(i, 2).Value)

            Print #fNum, dictWholeStore(cID) & "," & _
                         dictWholeData(cID) & "," & _
                         "1," & "WHOLE," & cID ' Qty is 1 per carton in this simplified whole model
        Next i
    End If

    ' 2. Export Partial Scans
    Set tbl = ThisWorkbook.Worksheets("Scan Partial").ListObjects("tblScanPartial")
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            Print #fNum, tbl.DataBodyRange(i, 3).Value & "," & _
                         tbl.DataBodyRange(i, 4).Value & "," & _
                         tbl.DataBodyRange(i, 5).Value & "," & _
                         "PARTIAL," & tbl.DataBodyRange(i, 2).Value
        Next i
    End If

    Close #fNum

    MsgBox "WMS Export generated: " & fileName, vbInformation
    Exit Sub

ErrorHandler:
    Close #fNum
    modUtils.HandleError "GenerateWMSExport"
End Sub

Public Sub GenerateMissingReport()
    On Error GoTo ErrorHandler
    Dim wsMissing As Worksheet: Set wsMissing = ThisWorkbook.Worksheets("Missing")
    Dim tblMissing As ListObject: Set tblMissing = wsMissing.ListObjects("tblMissing")
    modUtils.ClearTable tblMissing

    Dim tblPL As ListObject: Set tblPL = Range("tblOriginalPL").ListObject
    Dim dictScanned As Object: Set dictScanned = CreateObject("Scripting.Dictionary")

    ' Build dictionary of all scanned cartons (Whole)
    Dim tblWholeScan As ListObject: Set tblWholeScan = Range("tblScanWhole").ListObject
    If Not tblWholeScan.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To tblWholeScan.ListRows.Count
            dictScanned(CStr(tblWholeScan.DataBodyRange(i, 2).Value)) = True
        Next i
    End If

    ' Loop through PL and find items not in dictScanned (for Whole) or partial items
    ' Simplified: Check every Carton_ID in PL
    For i = 1 To tblPL.ListRows.Count
        Dim cID As String: cID = CStr(tblPL.DataBodyRange(i, 2).Value)
        If Not dictScanned.Exists(cID) Then
            ' Add to missing table
            Dim newRow As ListRow: Set newRow = tblMissing.ListRows.Add
            newRow.Range(1, 1).Value = cID
            newRow.Range(1, 2).Value = tblPL.DataBodyRange(i, 3).Value ' SKU
            newRow.Range(1, 3).Value = tblPL.DataBodyRange(i, 5).Value ' Qty
            newRow.Range(1, 4).Value = "NOT SCANNED"
        End If
    Next i

    MsgBox "Missing Report Updated.", vbInformation
    Exit Sub
ErrorHandler:
    modUtils.HandleError "GenerateMissingReport"
End Sub
