Attribute VB_Name = "WorksheetHandlers"
Option Explicit

' IMPORTANT: These subroutines must be placed inside the respective Worksheet Objects in the VBA Editor.

' ---
' Place in Worksheet "First Scan"
' ---
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Assuming scanner input is in cell E2
    If Not Intersect(Target, Me.Range("E2")) Is Nothing Then
        If Target.Value <> "" Then
            modScanner.ProcessFirstScan Target.Value
            Target.ClearContents
            Target.Activate
        End If
    End If
End Sub

' ---
' Place in Worksheet "SCAN WHOLE"
' ---
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Assuming scanner input is in cell E2
    If Not Intersect(Target, Me.Range("E2")) Is Nothing Then
        If Target.Value <> "" Then
            modScanner.ProcessWholeScan Target.Value
            Target.ClearContents
            Target.Activate
        End If
    End If
End Sub

' ---
' Place in Worksheet "Scan Partial"
' ---
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Assuming Box ID in E2, SKU scan in E4
    If Not Intersect(Target, Me.Range("E4")) Is Nothing Then
        If Target.Value <> "" Then
            modScanner.ProcessPartialScan Me.Range("E2").Value, Target.Value
            Target.ClearContents
            Target.Activate
        End If
    End If
End Sub
