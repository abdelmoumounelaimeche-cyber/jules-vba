' --- Scanner UI Sheet Module ---
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Detect scan in cell C4
    If Not Intersect(Target, Me.Range("C4")) Is Nothing Then
        If Me.Range("C4").Value <> "" Then
            M_Scanner_UI.HandleScannerInput Me.Range("C4").Value
        End If
    End If
End Sub
