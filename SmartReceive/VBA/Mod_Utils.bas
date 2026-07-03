Attribute VB_Name = "Mod_Utils"
Option Explicit

' To be placed in ThisWorkbook module
' -----------------------------------------------------------------------
' Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'    If Sh.Name = SH_FIRST_SCAN Or Sh.Name = SH_SCAN_WHOLE Or Sh.Name = SH_SCAN_PARTIAL Then
'        If Not Intersect(Target, Sh.Range("B2")) Is Nothing Then
'            Application.EnableEvents = False
'            Mod_Scanner.HandleScan Target.Value
'            Target.ClearContents
'            Target.Select
'            Application.EnableEvents = True
'        End If
'    End If
' End Sub
' -----------------------------------------------------------------------

Public Sub ToggleProtection(ByVal protect As Boolean, Optional ByVal password As String = "")
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If protect Then
            ws.Protect Password:=password, UserInterfaceOnly:=True, DrawingObjects:=False
        Else
            ws.Unprotect Password:=password
        End If
    Next ws
End Sub

Public Sub ClearTable(ByVal tableName As String)
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Range(tableName).ListObject
    On Error GoTo 0

    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.Delete
        End If
    End If
End Sub

Public Function GetColIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    GetColIndex = tbl.ListColumns(colName).Index
    On Error GoTo 0
End Function

Public Sub InitialSetup()
    Application.ScreenUpdating = False
    SetupScanSheet SH_FIRST_SCAN
    SetupScanSheet SH_SCAN_WHOLE
    SetupScanSheet SH_SCAN_PARTIAL
    SetupSegregationSheet
    SetupDashboard
    Application.ScreenUpdating = True
    MsgBox "Initial Setup Complete. Please ensure ThisWorkbook code is in place.", vbInformation
End Sub
