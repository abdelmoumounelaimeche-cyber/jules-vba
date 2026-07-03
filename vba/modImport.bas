Attribute VB_Name = "modImport"
Option Explicit

' ---
' Module: modImport
' Purpose: Handles PL Import and Carton Classification
' ---

Public Sub ImportAndProcessPL()
    On Error GoTo ErrorHandler

    Dim filePath As Variant
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls; *.csv), *.xlsx; *.xls; *.csv", , "Select Supplier PL")

    If filePath = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1. Clear existing data
    modUtils.ClearTable ThisWorkbook.Worksheets("Original PL").ListObjects("tblOriginalPL")
    modUtils.ClearTable ThisWorkbook.Worksheets("Data Whole").ListObjects("tblDataWhole")
    modUtils.ClearTable ThisWorkbook.Worksheets("Partial").ListObjects("tblPartial")

    ' 2. Import Data from selected file
    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)

    ' Assume data is in first sheet
    Dim lastRow As Long
    lastRow = srcWb.Sheets(1).Cells(srcWb.Sheets(1).Rows.Count, "A").End(xlUp).Row

    ' Copy data to tblOriginalPL
    srcWb.Sheets(1).Range("A2:E" & lastRow).Copy
    ThisWorkbook.Worksheets("Original PL").ListObjects("tblOriginalPL").DataBodyRange(1, 1).PasteSpecial xlPasteValues

    srcWb.Close False

    ' 3. Process data using Arrays for speed
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("Original PL").ListObjects("tblOriginalPL")

    Dim data() As Variant
    data = tbl.DataBodyRange.Value

    Dim i As Long, j As Long
    For i = LBound(data, 1) To UBound(data, 1)
        For j = LBound(data, 2) To UBound(data, 2)
            data(i, j) = Trim(UCase(CStr(data(i, j))))
        Next j
    Next i

    tbl.DataBodyRange.Value = data

    ' 4. Classify Cartons
    ClassifyCartons tbl

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "PL Processed successfully!", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    modUtils.HandleError "ImportAndProcessPL"
End Sub

Private Sub ClassifyCartons(tblPL As ListObject)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim data As Variant
    data = tblPL.DataBodyRange.Value

    Dim i As Long
    Dim cartonID As String

    ' Count SKUs per Carton
    For i = 1 To UBound(data, 1)
        cartonID = data(i, 2) ' Carton_ID column
        If Not dict.Exists(cartonID) Then
            dict.Add cartonID, 1
        Else
            dict(cartonID) = dict(cartonID) + 1
        End If
    Next i

    ' Distribute to Tables
    Dim tblWhole As ListObject, tblPartial As ListObject
    Set tblWhole = ThisWorkbook.Worksheets("Data Whole").ListObjects("tblDataWhole")
    Set tblPartial = ThisWorkbook.Worksheets("Partial").ListObjects("tblPartial")

    For i = 1 To UBound(data, 1)
        cartonID = data(i, 2)
        Dim targetTbl As ListObject
        If dict(cartonID) = 1 Then
            Set targetTbl = tblWhole
        Else
            Set targetTbl = tblPartial
        End If

        Dim newRow As ListRow
        Set newRow = targetTbl.ListRows.Add
        ' Column mapping (Simplified for brevity)
        newRow.Range(1, 1).Value = data(i, 2) ' Carton_ID
        newRow.Range(1, 2).Value = data(i, 1) ' Store
        newRow.Range(1, 3).Value = data(i, 3) ' SKU
        newRow.Range(1, 4).Value = data(i, 5) ' Qty
        newRow.Range(1, 5).Value = "PENDING"
    Next i
End Sub
