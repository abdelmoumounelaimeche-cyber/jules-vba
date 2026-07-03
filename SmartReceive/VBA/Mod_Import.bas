Attribute VB_Name = "Mod_Import"
Option Explicit

Public Sub ImportPackingList()
    Dim filePath As Variant, wbSource As Workbook, wsDest As Worksheet, tblDest As ListObject
    Application.ScreenUpdating = False
    filePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", Title:="Select Packing List File")
    If filePath = False Then Exit Sub

    On Error GoTo ErrorHandler
    Set wbSource = Workbooks.Open(filePath)
    Set wsDest = ThisWorkbook.Sheets(SH_ORIGINAL_PL)
    Set tblDest = wsDest.ListObjects(TBL_ORIGINAL_PL)
    ClearTable TBL_ORIGINAL_PL

    Dim lastRow As Long
    lastRow = wbSource.Sheets(1).Cells(wbSource.Sheets(1).Rows.Count, "A").End(xlUp).Row
    wbSource.Sheets(1).Range("A2:F" & lastRow).Copy wsDest.Range("A2")
    wbSource.Close False
    ClassifyCartons
    Application.ScreenUpdating = True
    MsgBox "Import Successful.", vbInformation
    Exit Sub
ErrorHandler:
    If Not wbSource Is Nothing Then wbSource.Close False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub ClassifyCartons()
    Dim tbl As ListObject: Set tbl = ThisWorkbook.Sheets(SH_ORIGINAL_PL).ListObjects(TBL_ORIGINAL_PL)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim row As ListRow

    ' Use GetColIndex to avoid hardcoding
    Dim idxID As Long: idxID = GetColIndex(tbl, "CartonID")
    Dim idxSKU As Long: idxSKU = GetColIndex(tbl, "SKU")
    Dim idxClass As Long: idxClass = GetColIndex(tbl, "Class")

    For Each row In tbl.ListRows
        Dim id As String: id = row.Range(1, idxID).Value
        If Not dict.Exists(id) Then Set dict(id) = CreateObject("Scripting.Dictionary")
        dict(id)(row.Range(1, idxSKU).Value) = 1
    Next row

    For Each row In tbl.ListRows
        row.Range(1, idxClass).Value = IIf(dict(row.Range(1, idxID).Value).Count = 1, CLASS_WHOLE, CLASS_PARTIAL)
    Next row
End Sub
