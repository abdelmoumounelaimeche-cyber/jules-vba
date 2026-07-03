Attribute VB_Name = "modUI"
Option Explicit

' ---
' Module: modUI
' Purpose: Dashboard, Navigation and Operator Interface
' ---

Public Sub UpdateDashboard()
    ' Logic to refresh KPI numbers on Dashboard
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")

    wsDash.Range("KPI_TotalCartons").Value = modUtils.GetTableRowCount("tblOriginalPL")
    wsDash.Range("KPI_WholeCartons").Value = modUtils.GetTableRowCount("tblDataWhole")
    wsDash.Range("KPI_ScannedWhole").Value = modUtils.GetTableRowCount("tblScanWhole")
    wsDash.Range("KPI_PartialCartons").Value = modUtils.GetTableRowCount("tblPartial")
    wsDash.Range("KPI_Missing").Value = modUtils.GetTableRowCount("tblMissing")

    ' Update progress bar (simulated by cell width or conditional formatting)
End Sub

Public Sub NavigateTo(sheetName As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(sheetName).Activate
    If Err.Number <> 0 Then
        MsgBox "Sheet " & sheetName & " not found.", vbExclamation
    End If
End Sub

Public Sub ShowAssignment(store As String, pallet As String)
    ' This would update a large "Assignment" card on the Whole Scan sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("SCAN WHOLE")

    ws.Range("UI_StoreDisplay").Value = store
    ws.Range("UI_PalletDisplay").Value = pallet

    ' Voice feedback could be added here if needed
End Sub

Public Function GetActivePallet(store As String) As String
    ' Logic to manage pallet numbers (C-1, C-2...) per store
    ' For now, retrieval from Settings or a dedicated pallet tracker table
    GetActivePallet = modUtils.LookupValue("tblSettings", "Setting_Name", "CURRENT_PALLET_" & store, "Setting_Value")
    If GetActivePallet = "" Then GetActivePallet = "C-1"
End Function

Public Sub NextPallet()
    ' Triggered by "NEXT PALLET" button
    Dim currentStore As String
    currentStore = ThisWorkbook.Worksheets("SCAN WHOLE").Range("UI_StoreDisplay").Value

    If currentStore = "" Then Exit Sub

    Dim currentPallet As String
    currentPallet = GetActivePallet(currentStore)

    ' Increment C-1 to C-2
    Dim num As Integer
    num = CInt(Mid(currentPallet, 3)) + 1

    Dim newPallet As String
    newPallet = "C-" & num

    ' Save back to settings
    modUtils.UpdateSetting "CURRENT_PALLET_" & currentStore, newPallet

    ' Refresh UI
    ShowAssignment currentStore, newPallet
End Sub

Public Sub UpdateSegregationCard(cartonID As String)
    ' Displays large info for operator in Segregation sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Segregation")

    ' Lookup details from tblPartial
    ws.Range("Card_CartonID").Value = cartonID
    ws.Range("Card_Store").Value = modUtils.LookupValue("tblPartial", "Carton_ID", cartonID, "Store")
    ws.Range("Card_SKU").Value = modUtils.LookupValue("tblPartial", "Carton_ID", cartonID, "SKU")
    ws.Range("Card_Qty").Value = modUtils.LookupValue("tblPartial", "Carton_ID", cartonID, "Qty")
End Sub
