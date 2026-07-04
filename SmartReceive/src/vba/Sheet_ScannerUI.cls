Attribute VB_Name = "Sheet2"
Option Explicit

' Sheet 2 is Scanner UI
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Target barcode input cell D4
    If Not Intersect(Target, Me.Range("D4")) Is Nothing Then
        If Me.Range("D4").Value <> "" Then
            HandleScannerInput Me.Range("D4").Value
        End If
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub

    Select Case Target.Address
        Case "$F$6" ' SWITCH MODE
            Dim modes As Variant: modes = Array("First Scan", "Whole Scan", "Segregation")
            Dim curr As String: curr = Me.Range("D6").Value
            Dim i As Integer
            For i = 0 To UBound(modes)
                If modes(i) = curr Then
                    Me.Range("D6").Value = modes((i + 1) Mod 3)
                    Exit For
                End If
            Next i
        Case "$F$8" ' CLEAR
            Me.Range("D4").Value = ""
    End Select
End Sub
